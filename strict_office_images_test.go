package officeread

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"math"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestPPTXStrictImagesIncludeVMLPictureShapes(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree/></p:cSld><p:legacyDrawing r:id="rIdVML"/></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVML" Type="x" Target="../drawings/vmlDrawing1.vml"/></Relationships>`)
	addZip(t, zw, "ppt/drawings/vmlDrawing1.vml", `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape id="visible"><v:imagedata o:relid="rId1"/></v:shape><v:shape id="hidden" style="visibility:hidden"><v:imagedata o:relid="rId2"/></v:shape></xml>`)
	addZip(t, zw, "ppt/drawings/_rels/vmlDrawing1.vml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="../media/visible.png"/><Relationship Id="rId2" Type="x" Target="../media/hidden.jpg"/></Relationships>`)
	addZipBytes(t, zw, "ppt/media/visible.png", testPNG())
	addZipBytes(t, zw, "ppt/media/hidden.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "vml-pictures.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 1 || result.Images[0].Name != "visible.png" {
		t.Fatalf("strict PPTX VML images = %#v, want only visible.png", result.Images)
	}
}

func TestPPTXStrictImagesExcludeHTMLImportVMLStagingFrames(t *testing.T) {
	ids, err := pptxVMLPictureRelationshipIDOccurrences([]byte(`<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape id="HTMLText1" style="visibility:visible"><v:imagedata o:relid="rIdText"/></v:shape><v:shape id="HTMLHidden1"><v:imagedata o:relid="rIdHidden"/></v:shape><v:shape id="ordinary"><v:imagedata o:relid="rIdVisible"/></v:shape></xml>`))
	if err != nil {
		t.Fatal(err)
	}
	if len(ids) != 1 || ids[0] != "rIdVisible" {
		t.Fatalf("strict PPTX VML ids = %#v, want only ordinary visible picture", ids)
	}
}

func TestPPTX20457StrictImageCountMatchesPowerPointShapes(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/pptx/00020457.pptx", Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// PowerPoint exposes the eight DrawingML p:pic shapes on the slides. Its
	// companion VML part has 40 HTML importer staging frames, none of which is
	// a Slide.Shape.
	if len(result.Images) != 8 {
		t.Fatalf("strict PPTX images = %d, want 8 visible PowerPoint pictures", len(result.Images))
	}
}

func TestPPTXStrictImagesExcludeUnreferencedVMLCaches(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p"><p:cSld><p:spTree/></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVML" Type="x" Target="../drawings/vmlDrawing1.vml"/></Relationships>`)
	addZip(t, zw, "ppt/drawings/vmlDrawing1.vml", `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape id="Chart_x0020_cache"><v:imagedata o:relid="rId1"/></v:shape></xml>`)
	addZip(t, zw, "ppt/drawings/_rels/vmlDrawing1.vml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="../media/cache.png"/></Relationships>`)
	addZipBytes(t, zw, "ppt/media/cache.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-vml-cache.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 0 {
		t.Fatalf("strict PPTX VML cache images = %#v, want none", result.Images)
	}
}

func TestStrictPPTXRecoversOfficeVisiblePictureWithMissingRelationship(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/pptx/00022653.pptx", Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// PowerPoint exposes 122 Picture Shapes. Slide 36 contains one valid
	// p:pic whose rId11 relationship was omitted by its producer; Office
	// recovers image111.png, and strict extraction must do the same.
	if len(result.Images) != 122 {
		t.Fatalf("strict PPTX images = %d, want 122 PowerPoint Picture Shapes", len(result.Images))
	}
}

func TestDocxImageRelationshipRefsExcludeDocumentBackground(t *testing.T) {
	refs, err := docxImageRelationshipRefs([]byte(`
		<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml">
		  <w:background><v:fill r:id="background"/></w:background>
		  <w:body><w:p><w:r><w:drawing r:embed="picture"/></w:r></w:p></w:body>
		</w:document>`))
	if err != nil {
		t.Fatal(err)
	}
	if !refs.Hidden["background"] {
		t.Fatalf("background relationship must be hidden: %#v", refs)
	}
	if !refs.Visible["picture"] {
		t.Fatalf("ordinary body image must remain visible: %#v", refs)
	}
}

func TestParseSSTRecordsContinuesSharedString(t *testing.T) {
	records := [][]byte{{1, 0, 0, 0, 1, 0, 0, 0, 5, 0, 0, 'h', 'e', 'l'}, {0, 'l', 'o'}}
	got := parseSSTRecords(records)
	if len(got) != 1 || got[0] != "hello" {
		t.Fatalf("parseSSTRecords = %#v, want [hello]", got)
	}
}

func TestStrictLegacyXLSDoesNotExposeUnreferencedCarvedPicture(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/xls/005785.xls", Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 0 {
		t.Fatalf("strict XLS images = %d, want 0 (Excel Shapes)", len(result.Images))
	}
}

func TestStrictLegacyXLSFollowsContinuedOfficeArtPictures(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/xls/001109.xls", Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 29 {
		t.Fatalf("strict XLS images = %d, want 29 visible Excel Picture shapes", len(result.Images))
	}
}

func TestStrictLegacyXLSFollowsDrawingGroupAcrossBIFFRecords(t *testing.T) {
	// Excel stores this workbook-global DggContainer over multiple BIFF
	// records.  The 70 visible Picture shapes refer to its 30 BStore blips.
	result, err := Extract("testdata/web-samples/samples/xls/011121.xls", Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 70 {
		t.Fatalf("strict XLS images = %d, want 70 visible Excel Picture shapes", len(result.Images))
	}
}

func TestStrictLegacyXLSFollowsSingleOfficeArtPicture(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/xls/000054.xls", Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 1 {
		t.Fatalf("strict XLS images = %d, want 1 visible Excel Picture shape", len(result.Images))
	}
}

func TestStrictLegacyXLSUsesCustomDateFormat(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/xls/001474.xls", Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Jan-2001", "Dec-2005"} {
		if !strings.Contains(result.Text, want) {
			t.Fatalf("strict XLS text missing custom-formatted date %q: %q", want, result.Text)
		}
	}
}

func TestStrictLegacyXLSUsesCustomDateFormatAcrossYearBoundaries(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/xls/001478.xls", Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Jan-1987", "Dec-1987"} {
		if !strings.Contains(result.Text, want) {
			t.Fatalf("strict XLS text missing custom-formatted date %q", want)
		}
	}
}

func TestStrictLegacyXLSUsesYearOnlyCustomDateFormat(t *testing.T) {
	// Excel's custom "yyyy" format renders a serial as a year, not as the
	// underlying serial number.  This exact pattern occurs in the corpus
	// (004905.xls) and is part of the strict Office-visible text contract.
	formats := biffNumberFormats{custom: map[uint16]string{164: "yyyy"}}
	got, ok := biffFormattedNumberDisplayValue(35611, 164, formats)
	if !ok || got != "1997" {
		t.Fatalf("year-only custom date = %q, %v", got, ok)
	}
}

func TestStrictLegacyXLSUsesBuiltInPercentFormat(t *testing.T) {
	formats := biffNumberFormats{}
	for _, tc := range []struct {
		format uint16
		value  float64
		want   string
	}{
		{format: 9, value: 0.130523321, want: "13%"},
		{format: 10, value: 0.130523321, want: "13.05%"},
	} {
		got, ok := biffFormattedNumberDisplayValue(tc.value, tc.format, formats)
		if !ok || got != tc.want {
			t.Fatalf("built-in percent format %d = %q, %v; want %q", tc.format, got, ok, tc.want)
		}
	}
}

func TestStrictLegacyXLSUsesBuiltInScientificAndTimeFormats(t *testing.T) {
	for _, tc := range []struct {
		format uint16
		value  float64
		want   string
	}{
		{format: 11, value: 12345, want: "1.23E+4"},
		{format: 18, value: 0.5, want: "12:00 PM"},
		{format: 19, value: 0.5, want: "12:00:00 PM"},
		{format: 20, value: 1.5 / 24, want: "1:30"},
		{format: 45, value: 61.0 / 86400, want: "01:01"},
		{format: 46, value: 25.5 / 24, want: "25:30:00"},
		{format: 47, value: 61.2 / 86400, want: "01:01.2"},
		{format: 37, value: -1200, want: "(1,200)"},
		{format: 40, value: -1200.5, want: "(1,200.50)"},
	} {
		got, ok := biffFormattedNumberDisplayValue(tc.value, tc.format)
		if !ok || got != tc.want {
			t.Fatalf("built-in format %d = %q, %v; want %q", tc.format, got, ok, tc.want)
		}
	}
}

func TestStrictLegacyXLSUsesCustomPercentFormat(t *testing.T) {
	formats := biffNumberFormats{custom: map[uint16]string{164: "0.0%"}}
	got, ok := biffFormattedNumberDisplayValue(0.111, 164, formats)
	if !ok || got != "11.1%" {
		t.Fatalf("custom percent format = %q, %v; want 11.1%%", got, ok)
	}
}

func TestStrictLegacyXLSPrefersWorkbookStreamOverBook(t *testing.T) {
	// This real workbook retains a short stale Book stream next to its actual
	// Workbook stream.  The former has only 254 visible values while Office
	// opens the latter, which has 574.  Stream directory order is not a valid
	// authority for selecting the active workbook.
	b, err := os.ReadFile("testdata/web-samples/samples/xls/001898.xls")
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(b)
	if err != nil {
		t.Fatal(err)
	}
	book, workbook := 0, 0
	for _, s := range streams {
		if s.Name == "Book" {
			book = len(biffStrictOfficeText(s.Data))
		}
		if s.Name == "Workbook" {
			workbook = len(biffStrictOfficeText(s.Data))
		}
	}
	if book == 0 || workbook <= book {
		t.Fatalf("fixture does not demonstrate Book/Workbook distinction: Book=%d Workbook=%d", book, workbook)
	}
	got := len(biffStrictOfficeText(legacyWorkbookBytes(b, streams)))
	if got != workbook {
		t.Fatalf("strict BIFF values=%d, want Workbook stream values=%d (Book=%d)", got, workbook, book)
	}
	result, err := Extract("testdata/web-samples/samples/xls/001898.xls", Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Table 668", "Anchorage, AK", "Washington, DC"} {
		if !strings.Contains(result.Text, want) {
			t.Fatalf("strict XLS output chose stale Book stream; missing %q", want)
		}
	}
}

func TestLegacyXLSFormulaUsesCellNumberFormat(t *testing.T) {
	var formats biffNumberFormats
	formats.xf = []uint16{14}
	record := make([]byte, 14)
	binary.LittleEndian.PutUint64(record[6:], math.Float64bits(36906))
	if got, ok := biffFormulaDisplayValue(record, formats); !ok || got != "2001/1/15" {
		t.Fatalf("formatted formula = %q, %v", got, ok)
	}
}

func TestStrictLegacyXLSFormatsThousandsGroupedValues(t *testing.T) {
	if got := biffThousandsGroupedInt(-1890641); got != "-1,890,641" {
		t.Fatalf("grouped integer = %q", got)
	}
	got, ok := biffFormattedNumberDisplayValue(78495, 3)
	if !ok || got != "78,495" {
		t.Fatalf("format 3 = %q, %v", got, ok)
	}
	got, ok = biffFormattedNumberDisplayValue(38785+float64(12*60+34)/(24*60), 22)
	if !ok || got != "2006/3/9 12:34" {
		t.Fatalf("format 22 = %q, %v", got, ok)
	}
}

func TestStrictDOCXImagesExcludeHeaderFooterPictures(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/docx/00001714.docx", Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 0 {
		t.Fatalf("strict DOCX images = %d, want Word.Document Shapes count 0", len(result.Images))
	}
}

func TestStrictDOCXImagesExcludeDiagramResources(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/docx/223624.docx", Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// Word exposes the one InlineShape and two msoPicture Shapes. The document
	// package also contains ten diagram/graphic resources, which must not be
	// counted as Document.InlineShapes/Shapes pictures.
	if len(result.Images) != 3 {
		t.Fatalf("strict DOCX images = %d, want 3", len(result.Images))
	}
}

func TestStrictDOCXImagesKeepVMLGroupPictureChildren(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/docx/313486.docx", Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// Word COM exposes 239 inline DrawingML pictures, two picture children in
	// a floating VML group, and one floating DrawingML picture.
	if len(result.Images) != 242 {
		t.Fatalf("strict VML group picture occurrences = %d, want 242", len(result.Images))
	}
}

func TestRelationshipTargetMapKeepsMediaFilenameThatLooksLikeMetadata(t *testing.T) {
	targets, err := relationshipTargetMap([]byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="image" Target="media/image111.jpeg"/></Relationships>`))
	if err != nil {
		t.Fatal(err)
	}
	if got := targets["rIdImage"]; got != "media/image111.jpeg" {
		t.Fatalf("media relationship target = %q, want media/image111.jpeg", got)
	}
}

func TestStrictDOCXImagesDoNotDuplicateEmbeddedFallbackForExternalLink(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:drawing><pic:pic><pic:blipFill><a:blip r:embed="rIdEmbedded" r:link="rIdExternal"/></pic:blipFill></pic:pic></w:drawing></w:r></w:p></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEmbedded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/><Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/unavailable.png" TargetMode="External"/></Relationships>`)
	addZip(t, zw, "word/media/image1.png", string(testPNG()))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	file := filepath.Join(t.TempDir(), "external-linked-fallback.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 1 {
		t.Fatalf("strict linked fallback image occurrences = %d, want 1", len(result.Images))
	}
}

func TestStrictDOCXImagesExcludeTextBoxAndAlternateFallbackPictures(t *testing.T) {
	for _, tc := range []struct {
		file string
		want int
	}{
		// Word exposes a VML picture inside this floating text box as a Shape,
		// but it is not a Document.InlineShapes/Shapes picture occurrence.
		{file: "testdata/web-samples/samples/docx/00000387.docx", want: 0},
		// The package has five media placements. Two are an Office 2010 group
		// member and its downlevel VML AlternateContent fallback, neither of
		// which Word surfaces as a standalone picture shape.
		{file: "testdata/web-samples/samples/docx/00000471.docx", want: 3},
	} {
		result, err := Extract(tc.file, Options{StrictOfficeImages: true})
		if err != nil {
			t.Fatal(err)
		}
		if got := len(result.Images); got != tc.want {
			t.Fatalf("%s strict DOCX images = %d, want %d", tc.file, got, tc.want)
		}
	}
}

func TestStrictDOCImagesExcludeEmbeddedObjectPreview(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/doc/006489.doc", Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 0 {
		t.Fatalf("strict DOC images = %d, want Word document Shapes count 0", len(result.Images))
	}
}

func TestStrictDOCImagesExcludeUnreferencedDataStreamPictureFrames(t *testing.T) {
	for _, file := range []string{
		"testdata/web-samples/samples/doc/000004.doc",
		"testdata/web-samples/samples/doc/000005.doc",
	} {
		result, err := Extract(file, Options{StrictOfficeImages: true})
		if err != nil {
			t.Fatal(err)
		}
		// The Data streams contain structurally valid PictureFrame/BSE bytes,
		// but Word exposes neither an InlineShape nor a floating picture.
		if len(result.Images) != 0 {
			t.Fatalf("%s strict DOC images = %d, want 0", file, len(result.Images))
		}
	}
}

func TestStrictDOCImagesExcludeUnreferencedPictureStoreBlip(t *testing.T) {
	for _, file := range []string{
		"testdata/web-samples/samples/doc/000259.doc",
		"testdata/web-samples/samples/doc/000259-2.doc",
	} {
		result, err := Extract(file, Options{StrictOfficeImages: true})
		if err != nil {
			t.Fatal(err)
		}
		if len(result.Images) != 0 {
			t.Fatalf("%s strict DOC images = %d, want 0", file, len(result.Images))
		}
	}
}

func TestStrictDOCImagesExcludeInlineEmbeddedControlPictureFrame(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/doc/000259.doc", Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// Word exposes its only InlineShape as Type=10, rather than a picture.
	if len(result.Images) != 0 {
		t.Fatalf("strict DOC images = %d, want 0", len(result.Images))
	}
}

func TestDOCPIcfConsumesRemainingData(t *testing.T) {
	data := make([]byte, 100)
	binary.LittleEndian.PutUint32(data[10:], uint32(len(data)-10))
	if !docPICFConsumesRemainingData(data, 10) {
		t.Fatal("PICF consuming the remaining Data stream was not identified")
	}
	binary.LittleEndian.PutUint32(data[10:], 42)
	if docPICFConsumesRemainingData(data, 10) {
		t.Fatal("bounded PICF was misidentified as a remaining-stream store")
	}
	if docPICFIsDegenerateInlineControl(data, 10) {
		t.Fatal("empty PICF must not be identified as a degenerate inline control")
	}
}

func TestStrictDOCImagesKeepPictureStoreWithoutEmbeddedObjectPool(t *testing.T) {
	for _, file := range []string{
		"testdata/web-samples/samples/doc/000795-2.doc",
		"testdata/web-samples/samples/doc/001962-2.doc",
	} {
		result, err := Extract(file, Options{StrictOfficeImages: true})
		if err != nil {
			t.Fatal(err)
		}
		if len(result.Images) != 1 {
			t.Fatalf("%s strict DOC images = %d, want 1", file, len(result.Images))
		}
	}
}

func TestStrictDOCImagesResolvesZeroPICFLocationThroughBSE(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/doc/000126.doc", Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// Word exposes one InlineShape picture. Its sprmCPicLocation is zero, and
	// the PictureFrame resolves the raster through the following BSE rather than
	// carrying an embedded blip payload.
	if len(result.Images) != 1 {
		t.Fatalf("strict DOC images = %d, want 1", len(result.Images))
	}
	if result.Images[0].Ext != ".wmf" {
		t.Fatalf("strict DOC image extension = %q, want .wmf", result.Images[0].Ext)
	}
}

func TestStrictDOCImagesExcludeInlineHorizontalRuleControl(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/doc/006417.doc", Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// Word exposes 165 InlineShape pictures and four floating pictures. A tiny
	// PNG carried by its Type=7 horizontal-rule control must not become a 170th
	// strict image occurrence.
	if len(result.Images) != 169 {
		t.Fatalf("strict DOC images = %d, want 169", len(result.Images))
	}
}

func TestStrictDOCImagesKeepsUnanchoredDIBInlinePicture(t *testing.T) {
	for _, file := range []string{
		"testdata/web-samples/samples/doc/004422.doc",
		"testdata/web-samples/samples/doc/004422-2.doc",
	} {
		result, err := Extract(file, Options{StrictOfficeImages: true})
		if err != nil {
			t.Fatal(err)
		}
		// These files have no PlcfSpaMom floating-shape anchor. Their only
		// mm==2 PICF is therefore a genuine Word InlineShape picture, not a
		// floating shape's duplicate backing resource.
		if got := len(result.Images); got != 1 {
			t.Fatalf("%s strict DOC images = %d, want 1", file, got)
		}
	}
}

func TestStrictDOCImagesUsesTableStoreForFloatingPictureFrames(t *testing.T) {
	result, err := Extract("testdata/web-samples/samples/doc/004581.doc", Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// Its Data stream is empty, but three PlcfSpaMom anchors resolve through
	// the selected table stream's BStore. Word exposes all three Shapes,
	// including the repeated blip-1 placement.
	if got := len(result.Images); got != 3 {
		t.Fatalf("strict DOC table-store floating images = %d, want 3", got)
	}
}
