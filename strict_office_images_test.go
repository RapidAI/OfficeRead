package officeread

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"os"
	"path/filepath"
	"testing"
)

func TestPPTXStrictImagesIncludeVMLPictureShapes(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree/></p:cSld></p:sld>`)
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
