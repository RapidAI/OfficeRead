package officeread

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"os"
	"strings"
	"testing"
)

func TestProbeLegacyWordContractionRepair(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/doc/000287.doc", Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"aren't", "didn't", "doesn't", "don't", "haven't", "isn't", "wasn't", "weren't", "wouldn't"} {
		if !strings.Contains(result.Text, want) {
			t.Fatalf("missing repaired contraction %q", want)
		}
	}
}

func TestProbe448024StrictImageOccurrences(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/docx/448024.docx", Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	for i, image := range result.Images {
		t.Logf("%d: %s (%d bytes)", i, image.Name, len(image.Data))
	}
}

func TestProbe00000867StrictText(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/docx/00000867.docx", Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, needle := range []string{"λ", "δ"} {
		for at := strings.Index(result.Text, needle); at >= 0; at = strings.Index(result.Text[at+len(needle):], needle) {
			end := at + 80
			if end > len(result.Text) {
				end = len(result.Text)
			}
			t.Logf("%q near %q", needle, result.Text[at:end])
			break
		}
	}
}

func TestProbePPT001012Structure(t *testing.T) {
	data, err := os.ReadFile("testdata/web-samples/samples/ppt/001012.ppt")
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	for i, slide := range pptActiveSlideContainers(streams) {
		var normal []string
		pptVisibleShapeTextInto(slide, 0, false, false, &normal)
		t.Logf("slide=%d normal=%q", i, normal)
		pptWalkRecords(slide, func(options, recordType uint16, payload []byte) {
			if recordType == 0xf00d || recordType == 0x0fa8 {
				t.Logf("slide=%d record options=%#x type=%#x size=%d preview=%q", i, options, recordType, len(payload), probeTrim(compressedUnicodeBytesToString(payload), 100))
			}
		})
	}
	if doc, ok := findLegacyStream(streams, "PowerPoint Document"); ok {
		for _, needle := range []string{"Entry Criteria", "All inputs", "Maintenance of configuration"} {
			at := strings.Index(string(doc.Data), needle)
			t.Logf("needle=%q stream-offset=%d", needle, at)
		}
		t.Logf("active persist=%v", pptActivePersistOffsets(streams))
		for _, off := range []int{6017, 3423, 4961, 37716} {
			if off+8 <= len(doc.Data) {
				t.Logf("persist offset=%d type=%#x size=%d", off, binary.LittleEndian.Uint16(doc.Data[off+2:]), binary.LittleEndian.Uint32(doc.Data[off+4:]))
			}
		}
		for at := 0; at+8 <= len(doc.Data); {
			op := binary.LittleEndian.Uint16(doc.Data[at:])
			ty := binary.LittleEndian.Uint16(doc.Data[at+2:])
			sz := int(binary.LittleEndian.Uint32(doc.Data[at+4:]))
			if sz < 0 || at+8+sz > len(doc.Data) { at++; continue }
			if ty == 0x03ee {
				payload := doc.Data[at+8:at+8+sz]
				if strings.Contains(string(payload), "Entry Criteria") { t.Logf("entry in slide container at=%d opts=%#x size=%d", at, op, sz) }
			}
			at += 8+sz
		}
	}
}

func TestProbePPTStrictCoverage(t *testing.T) {
	for _, name := range []string{"000295.ppt", "000305.ppt", "002447.ppt"} {
		data, err := os.ReadFile("testdata/web-samples/samples/ppt/" + name)
		if err != nil { t.Fatal(err) }
		streams, err := readOLEStreams(data)
		if err != nil { t.Fatal(err) }
		strict := extractPPTLegacyTextWithMode(streams, true)
		loose := extractPPTLegacyTextWithMode(streams, false)
		missing:=[]string{}
		seen:=map[string]bool{};for _,x:=range strict{seen[x]=true};for _,x:=range loose{if !seen[x]{missing=append(missing,x)}}
		if len(missing)>8{missing=missing[:8]}
		t.Logf("%s strictParts=%d strictWords=%d looseParts=%d looseWords=%d looseOnly=%q", name, len(strict), pptTextPartWords(strict), len(loose), pptTextPartWords(loose),missing)
	}
}

func TestProbeDOCXStrictPictureIDs(t *testing.T) {
	for _, name := range []string{"00000327.docx", "00000917.docx", "00002200.docx", "00004047.docx", "00004076.docx", "448024.docx"} {
		data, err := os.ReadFile("testdata/web-samples/samples/docx/" + name)
		if err != nil { t.Fatal(err) }
		zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
		if err != nil { t.Fatal(err) }
		files := map[string]*zip.File{}
		for _, f := range zr.File { files[f.Name] = f }
		document := ooxmlPartName(files, "word/document.xml")
		b, err := readZipFile(ooxmlFile(files, document))
		if err != nil { t.Fatal(err) }
		ids, err := docxStrictPictureRelationshipIDsInOrder(b)
		if err != nil { t.Fatal(err) }
		refs, err := docxImageRelationshipRefs(b)
		if err != nil { t.Fatal(err) }
		t.Logf("%s strict=%v visible=%v hidden=%v", name, ids, refs.Visible, refs.Hidden)
	}
}

func TestProbePPT001012SlidePersistAtoms(t *testing.T) {
	data, err := os.ReadFile("testdata/web-samples/samples/ppt/001012.ppt")
	if err != nil { t.Fatal(err) }
	streams, err := readOLEStreams(data)
	if err != nil { t.Fatal(err) }
	doc, _ := findLegacyStream(streams, "PowerPoint Document")
	offsets := pptActivePersistOffsets(streams)
	for id, at := range offsets {
		pos:=int(at)
		if pos+8 > len(doc.Data) || binary.LittleEndian.Uint16(doc.Data[pos+2:]) != 0x03ee { continue }
		sz:=int(binary.LittleEndian.Uint32(doc.Data[pos+4:])); if sz<0 || pos+8+sz>len(doc.Data){continue}
		payload:=doc.Data[pos+8:pos+8+sz]
		if strings.Contains(string(payload),"Entry Criteria") {
			t.Logf("slide persist=%d offset=%d size=%d",id,pos,sz)
			var scan func([]byte,int)
			scan=func(data []byte,depth int){if depth>12{return};for off:=0;off+8<=len(data);{op:=binary.LittleEndian.Uint16(data[off:]);ty:=binary.LittleEndian.Uint16(data[off+2:]);n:=int(binary.LittleEndian.Uint32(data[off+4:]));off+=8;if n<0||off+n>len(data){return};q:=data[off:off+n];if ty==0x03f3{t.Logf("slidepersist atom depth=%d options=%#x size=%d data=%x",depth,op,n,q)};if op&0xf==0xf{scan(q,depth+1)};off+=n}}
			scan(payload,0)
		}
	}
}



func TestProbe000948StrictImages(t *testing.T) {
	data, err := os.ReadFile("testdata/web-samples/samples/doc/000948-2.doc")
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	word, _ := findLegacyStream(streams, "WordDocument")
	table, _ := docSelectedTableStream(streams, word.Data)
	dataStream, _ := findLegacyStream(streams, "Data")
	t.Logf("main FC=%v CHPX=%#v", docMainStoryFCIntervals(word.Data, table.Data), docCHPXPictureLocationsWithRanges(word.Data, table.Data))
	t.Logf("inline=%d floating=%d all-floating=%d", len(docVisibleInlinePICFImages(streams, dataStream.Data)), len(docVisibleFloatingShapeImages(streams, dataStream.Data)), len(docAllFloatingShapeImages(streams, dataStream.Data)))
	t.Logf("PICF@0 header=%v degenerate=%v image=%v; PICF@21275 header=%v degenerate=%v image=%v", docPICFHeaderLooksValid(dataStream.Data, 0), docPICFIsDegenerateNonPicture(dataStream.Data, 0), func() bool { _, ok := docPICFImageAt(dataStream.Data, 0); return ok }(), docPICFHeaderLooksValid(dataStream.Data, 21275), docPICFIsDegenerateNonPicture(dataStream.Data, 21275), func() bool { _, ok := docPICFImageAt(dataStream.Data, 21275); return ok }())
	for i, entry := range docMainStoryFSPAEntries(word.Data, table.Data) {
		t.Logf("FSPA %d: %#v", i, entry)
	}
}

func TestProbe000113FieldControls(t *testing.T) {
	data, err := os.ReadFile("testdata/web-samples/samples/doc/000113-3.doc")
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	word, _ := findLegacyStream(streams, "WordDocument")
	table, _ := docSelectedTableStream(streams, word.Data)
	parts := wordMainStoryText(word.Data, table.Data)
	joined := strings.Join(parts, "\n")
	for _, needle := range []string{"HYPERLINK", "041207201"} {
		at := strings.Index(joined, needle)
		if at < 0 {
			t.Logf("%s absent", needle)
			continue
		}
		start := at - 40
		if start < 0 {
			start = 0
		}
		end := at + 180
		if end > len(joined) {
			end = len(joined)
		}
		t.Logf("%s raw=%q", needle, joined[start:end])
	}
	for _, value := range []struct{ name, text string }{{"clean", cleanVisibleText(joined)}, {"direct", stripWordFieldInstructions(joined)}} {
		at := strings.Index(value.text, "HYPERLINK")
		if at < 0 {
			t.Logf("%s: HYPERLINK absent", value.name)
			continue
		}
		end := at + 240
		if end > len(value.text) {
			end = len(value.text)
		}
		t.Logf("%s=%q", value.name, value.text[at:end])
	}
}

func TestProbeDOCStrictImageStructure(t *testing.T) {
	for _, name := range []string{"000913.doc", "000944.doc", "000950.doc"} {
		data, err := os.ReadFile("testdata/web-samples/samples/doc/" + name)
		if err != nil {
			t.Fatal(err)
		}
		streams, err := readOLEStreams(data)
		if err != nil {
			t.Fatal(err)
		}
		word, _ := findLegacyStream(streams, "WordDocument")
		table, _ := docSelectedTableStream(streams, word.Data)
		dataStream, _ := findLegacyStream(streams, "Data")
		var nonZero []docCHPXPictureLocation
		for _, pic := range docCHPXPictureLocationsWithRanges(word.Data, table.Data) {
			if pic.fcPic != 0 {
				nonZero = append(nonZero, pic)
			}
		}
		frames := docOfficeArtPictureFrames(table.Data)
		if len(frames) == 0 {
			frames = docOfficeArtPictureFrames(dataStream.Data)
		}
		t.Logf("%s main=%v inline=%d floating=%d allfloat=%d nonzero=%#v", name, docMainStoryFCIntervals(word.Data, table.Data), len(docVisibleInlinePICFImages(streams, dataStream.Data)), len(docVisibleFloatingShapeImages(streams, dataStream.Data)), len(docAllFloatingShapeImages(streams, dataStream.Data)), nonZero)
		for _, entry := range docMainStoryFSPAEntries(word.Data, table.Data) {
			frame, found := frames[entry.shapeID]
			if found {
				t.Logf("%s FSPA anchor=%d shape=%d frame=%#v found=%v", name, entry.anchor, entry.shapeID, frame, found)
			}
		}
		for _, pic := range nonZero {
			img, ok := docPICFImageAt(dataStream.Data, pic.fcPic)
			framesAtPICF := docOfficeArtPictureFrames(dataStream.Data[pic.fcPic:])
			t.Logf("%s PICF fc=%d image=%v ext=%s bytes=%d frames=%#v", name, pic.fcPic, ok, img.Ext, len(img.Data), framesAtPICF)
		}
	}
}

func TestProbeLowRecallMainStory(t *testing.T) {
	for _, name := range []string{"005715.doc", "005839.doc"} {
		data, err := os.ReadFile("testdata/web-samples/samples/doc/" + name)
		if err != nil {
			t.Fatal(err)
		}
		streams, err := readOLEStreams(data)
		if err != nil {
			t.Fatal(err)
		}
		word, _ := findLegacyStream(streams, "WordDocument")
		table, _ := docSelectedTableStream(streams, word.Data)
		strict := strings.Join(wordMainStoryText(word.Data, table.Data), "\n")
		all := strings.Join(wordPieceTableText(word.Data, table.Data), "\n")
		ccp := uint32(0)
		if len(word.Data) >= 0x50 {
			ccp = binary.LittleEndian.Uint32(word.Data[0x4c:])
		}
		fcCLX := int(binary.LittleEndian.Uint32(word.Data[0x1a2:]))
		lcbCLX := int(binary.LittleEndian.Uint32(word.Data[0x1a6:]))
		rawParts := parseWordCLXTextUntilCP(word.Data, table.Data[fcCLX:fcCLX+lcbCLX], ccp, true)
		fcMin, fcMac := uint32(0), uint32(0)
		if len(word.Data) >= 0x20 {
			fcMin = binary.LittleEndian.Uint32(word.Data[0x18:])
			fcMac = binary.LittleEndian.Uint32(word.Data[0x1c:])
		}
		result, err := Extract("testdata/web-samples/samples/doc/"+name, Options{StrictOfficeContent: true})
		if err != nil {
			t.Fatal(err)
		}
		rawLen := 0
		if len(rawParts) > 0 {
			rawLen = len(rawParts[0])
		}
		t.Logf("%s ccp=%d fc=%d..%d strict=%d raw=%d/%d extracted=%d all=%d preview=%q", name, ccp, fcMin, fcMac, len(strict), len(rawParts), rawLen, len(result.Text), len(all), probeTrim(strict, 300))
		logMainStoryPieces(t, word.Data, table.Data, ccp)
		if len(rawParts) > 0 {
			s := cleanTextNoMojibakeRepair(normalizeWordSpecialTextChars(normalizeWordOfficeContentControls(rawParts[0])))
			t.Logf("%s field-clean input=%d output=%d structured=%v", name, len(s), len(stripWordFieldInstructions(s)), wordHyperlinkFieldRE.MatchString(s) || wordNamedFieldRE.MatchString(s) || wordSimpleFieldRE.MatchString(s) || wordEmbedFieldRE.MatchString(s) || wordPrivateAddinFieldRE.MatchString(s))
		}
	}
}

func probeTrim(s string, n int) string {
	if len(s) <= n {
		return s
	}
	return s[:n] + "..."
}

func logMainStoryPieces(t *testing.T, word, table []byte, maxCP uint32) {
	fc := int(binary.LittleEndian.Uint32(word[0x1a2:]))
	lcb := int(binary.LittleEndian.Uint32(word[0x1a6:]))
	if fc < 0 || lcb < 0 || fc+lcb > len(table) {
		return
	}
	clx := table[fc : fc+lcb]
	for off := 0; off < len(clx); {
		if clx[off] == 1 {
			if off+3 > len(clx) {
				return
			}
			off += 3 + int(binary.LittleEndian.Uint16(clx[off+1:]))
			continue
		}
		if clx[off] != 2 || off+5 > len(clx) {
			off++
			continue
		}
		sz := int(binary.LittleEndian.Uint32(clx[off+1:]))
		if sz < 4 || off+5+sz > len(clx) {
			return
		}
		plc := clx[off+5 : off+5+sz]
		pieces := (len(plc) - 4) / 12
		pcd := (pieces + 1) * 4
		for i := 0; i < pieces; i++ {
			start, end := binary.LittleEndian.Uint32(plc[i*4:]), binary.LittleEndian.Uint32(plc[(i+1)*4:])
			if start >= maxCP {
				continue
			}
			if end > maxCP {
				end = maxCP
			}
			raw := binary.LittleEndian.Uint32(plc[pcd+i*8+2:])
			pos := int(raw & 0x3fffffff)
			compressed := raw&0x40000000 != 0
			if compressed {
				pos /= 2
			}
			var decoded string
			if compressed && pos >= 0 && int(end-start) <= len(word)-pos {
				decoded = decodeWordSingleByteTextWithMode(word[pos:pos+int(end-start)], wordPieceLegacyCodePage(word, plc, pieces, pcd), true)
			} else if !compressed && pos >= 0 && int(end-start)*2 <= len(word)-pos {
				decoded = wordPieceUTF16BytesToStringWithMode(word[pos:pos+int(end-start)*2], wordPieceLegacyCodePage(word, plc, pieces, pcd), true)
			}
			cleaned := strings.Join(cleanWordVisibleTextParts(normalizeWordOfficeContentControls(decoded)), "\n")
			direct := cleanVisibleText(normalizeWordOfficeContentControls(decoded))
			slow := cleanTextNoMojibakeRepair(normalizeWordSpecialTextChars(normalizeWordOfficeContentControls(decoded)))
			t.Logf("piece cp=%d..%d fc=%d compressed=%v raw=%d clean=%d direct=%d slow=%d preview=%q", start, end, pos, compressed, len(decoded), len(cleaned), len(direct), len(slow), probeTrim(decoded, 80))
		}
		off += 5 + sz
	}
}

func TestProbe005305StrictImages(t *testing.T) {
	data, err := os.ReadFile("testdata/web-samples/samples/doc/005305.doc")
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	word, _ := findLegacyStream(streams, "WordDocument")
	table, _ := docSelectedTableStream(streams, word.Data)
	dataStream, _ := findLegacyStream(streams, "Data")
	result, err := Extract("testdata/web-samples/samples/doc/005305.doc", Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	frames := docOfficeArtPictureFrames(table.Data)
	if len(frames) == 0 {
		frames = docOfficeArtPictureFrames(dataStream.Data)
	}
	t.Logf("images=%d data=%d fspa=%d frames=%d inline=%d floating=%d", len(result.Images), len(dataStream.Data), len(docMainStoryFSPAEntries(word.Data, table.Data)), len(frames), len(docVisibleInlinePICFImages(streams, dataStream.Data)), len(docVisibleFloatingShapeImages(streams, dataStream.Data)))
	for _, e := range docMainStoryFSPAEntries(word.Data, table.Data) {
		f, ok := frames[e.shapeID]
		t.Logf("fspa shape=%d anchor=%d frame=%#v found=%v", e.shapeID, e.anchor, f, ok)
	}
	for _, pic := range docCHPXPictureLocationsWithRanges(word.Data, table.Data) {
		if pic.fcPic == 0 {
			continue
		}
		img, ok := docPICFImageAt(dataStream.Data, pic.fcPic)
		mm := uint16(0)
		if pic.fcPic >= 0 && pic.fcPic+16 <= len(dataStream.Data) {
			mm = binary.LittleEndian.Uint16(dataStream.Data[pic.fcPic+14:])
		}
		framesAt := docOfficeArtPictureFrames(dataStream.Data[pic.fcPic:])
		lcb, header := uint32(0), uint16(0)
		if pic.fcPic >= 0 && pic.fcPic+6 <= len(dataStream.Data) {
			lcb = binary.LittleEndian.Uint32(dataStream.Data[pic.fcPic:])
			header = binary.LittleEndian.Uint16(dataStream.Data[pic.fcPic+4:])
		}
		t.Logf("pic=%d marker=%v ole=%v embedded=%v mm=%d lcb=%d header=%d valid=%v decoded=%v ext=%s bytes=%d frames=%#v", pic.fcPic, docCHPXRunContainsPictureMarker(word.Data, pic.startFC, pic.endFC), docCHPXRunIsOLEObject(word.Data, table.Data, pic.startFC, pic.endFC), docCHPXRunIsInlineEmbeddedObject(word.Data, table.Data, pic.startFC, pic.endFC), mm, lcb, header, docPICFHeaderLooksValid(dataStream.Data, pic.fcPic), ok, img.Ext, len(img.Data), framesAt)
	}
}
