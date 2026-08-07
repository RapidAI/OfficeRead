package officeread

import (
	"archive/zip"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

func TestStrictDOCXSimpleDateFieldUsesCurrentWordValue(t *testing.T) {
	got, ok := wordSimpleDynamicFieldValue(" DATE \\* MERGEFORMAT ", time.Date(2026, time.August, 1, 0, 0, 0, 0, time.Local))
	if !ok || got != "8/1/2026" {
		t.Fatalf("simple DATE = %q, ok=%v", got, ok)
	}
}

func TestStrictDOCXComplexDateFieldUsesWordPicture(t *testing.T) {
	got, ok := wordDynamicFieldValue(`DATE  \@ "h时m分s秒"`, time.Date(2026, time.August, 1, 15, 31, 52, 0, time.Local))
	if !ok || got != "3时31分52秒" {
		t.Fatalf("complex DATE = %q, ok=%v", got, ok)
	}
}

func TestStrictDOCXComplexDateFieldUsesNumericWordPicture(t *testing.T) {
	got, ok := wordDynamicFieldValue(`DATE \@ "MM/DD/YY"`, time.Date(2026, time.August, 1, 15, 31, 52, 0, time.Local))
	if !ok || got != "08/01/26" {
		t.Fatalf("numeric complex DATE = %q, ok=%v", got, ok)
	}
}

func TestStrictDOCXComplexDateAndTimeFieldsUseLongMonthPicture(t *testing.T) {
	now := time.Date(2026, time.August, 2, 15, 31, 52, 0, time.Local)
	for _, instruction := range []string{`DATE \@ "MMMM d, yyyy"`, `TIME \@ "MMMM d, yyyy"`, `DATE \@ "d MMMM yyyy"`} {
		got, ok := wordDynamicFieldValue(instruction, now)
		if !ok {
			t.Fatalf("dynamic field %q was not recognized", instruction)
		}
		want := "August 2, 2026"
		if strings.Contains(instruction, "d MMMM") {
			want = "2 August 2026"
		}
		if got != want {
			t.Fatalf("dynamic field %q = %q, want %q", instruction, got, want)
		}
	}
}

func TestStrictDOCXContentKeepsSoftHyphenBoundaryAndOmitsInlineDrawing(t *testing.T) {
	b := []byte(`<w:document xmlns:w="urn:w" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body><w:p><w:r><w:t>Reloca</w:t><w:softHyphen/><w:t>tion</w:t></w:r><w:r><w:t>coo</w:t><w:drawing><wp:inline/></w:drawing><w:t>rdinate</w:t></w:r></w:p></w:body></w:document>`)
	got, err := visibleWordContentText(b)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(got, "Reloca tion") || !strings.Contains(got, "coordinate") {
		t.Fatalf("Word content boundaries = %q", got)
	}
}

func TestStrictDOCXContentOmitsSerializedRTFPicturePayload(t *testing.T) {
	b := []byte(`<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Before</w:t></w:r><w:r><w:t>{\*\shppict{\pict\pngblip\bin8 \x89PNGbinary}}</w:t></w:r><w:r><w:t>After</w:t></w:r></w:p></w:body></w:document>`)
	got, err := visibleWordContentText(b)
	if err != nil {
		t.Fatal(err)
	}
	if got != "BeforeAfter" {
		t.Fatalf("serialized RTF picture leaked into Word content: %q", got)
	}
}

func TestStrictDOCXContentOmitsLegacyWebFormMarkerStyles(t *testing.T) {
	b := []byte(`<w:document xmlns:w="urn:w"><w:body><w:p><w:pPr><w:pStyle w:val="z-TopofForm"/></w:pPr><w:r><w:t>Top of Form</w:t></w:r></w:p><w:p><w:r><w:t>Visible content</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="z-BottomofForm"/></w:pPr><w:r><w:t>Bottom of Form</w:t></w:r></w:p></w:body></w:document>`)
	got, err := visibleWordContentText(b)
	if err != nil {
		t.Fatal(err)
	}
	if got != "Visible content" {
		t.Fatalf("legacy web form markers leaked into Word content: %q", got)
	}
}

func TestStrictDOCXContentOmitsCollapsedHyperlinkCharacterStyle(t *testing.T) {
	b := []byte(`<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Visible hyperlink</w:t></w:r><w:r><w:rPr><w:rStyle w:val="acicollapsed1"/></w:rPr><w:t> hidden popup explanation</w:t></w:r></w:p></w:body></w:document>`)
	got, err := visibleWordContentText(b)
	if err != nil {
		t.Fatal(err)
	}
	if got != "Visible hyperlink" {
		t.Fatalf("collapsed hyperlink text leaked into Word content: %q", got)
	}
}

func TestStrictDOCXVMLGroupPictureMatchesWordGroupItem(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "00001526.docx")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// Word exposes the child VML picture through Shape.GroupItems even though
	// the top-level Shape is an msoGroup.  Strict extraction must preserve it.
	if got := len(result.Images); got != 1 {
		t.Fatalf("strict VML group images = %d, want 1", got)
	}
}

func TestStrictDOCXNestedVMLCanvasDoesNotBecomePictures(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "215384.docx")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// These canvas groups are Word msoCanvas Shapes rather than Picture Shapes;
	// only the independent floating picture belongs to Word.Shapes' picture set.
	if got := len(result.Images); got != 1 {
		t.Fatalf("strict VML canvas images = %d, want 1", got)
	}
}

func TestStrictDOCXIncludesPackageRootPictureTarget(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "dotnet__Open-XML-SDK__test_DocumentFormat.OpenXml.Tests.Assets_assets_TestDataStorage_O15Conformance_WD_CommentExTest_Comments-Sample-15-12-01_Comment041.docx")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if got := len(result.Images); got != 1 {
		t.Fatalf("strict package-root picture images = %d, want 1", got)
	}
}

func TestStrictDOCXSVGNonPictureIsExcluded(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "dotnet__Open-XML-SDK__test_DocumentFormat.OpenXml.Tests.Assets_assets_TestFiles_svg.docx")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// This SVG blip is exposed by this Word COM version as wdInlineShapeTextBox
	// rather than wdInlineShapePicture, so the strict image count is zero.
	if got := len(result.Images); got != 0 {
		t.Fatalf("strict SVG non-picture images = %d, want 0", got)
	}
}

func TestStrictDOCXGroupedPicturesKeepAllPictureOccurrences(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "00001310.docx")
	zr, err := zip.OpenReader(file)
	if err != nil {
		t.Fatal(err)
	}
	defer zr.Close()
	files := map[string]*zip.File{}
	for _, entry := range zr.File {
		files[entry.Name] = entry
	}
	document := ooxmlPartName(files, "word/document.xml")
	markup, err := readZipFile(ooxmlFile(files, document))
	if err != nil {
		t.Fatal(err)
	}
	ids, err := docxStrictPictureRelationshipIDsInOrder(markup)
	if err != nil {
		t.Fatal(err)
	}
	vmlIDs, err := docxVMLPictureRelationshipIDs(markup)
	if err != nil {
		t.Fatal(err)
	}
	if got := len(ids); got != 5 {
		t.Fatalf("strict grouped picture relationship occurrences = %d (%v), VML=%v, want 5", got, ids, vmlIDs)
	}
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// Word exposes five Picture Shapes here: three VML GroupItems and two
	// independent DrawingML pictures. Keep this sample as a guard against
	// chart-preview filtering accidentally suppressing real group pictures.
	if got := len(result.Images); got != 5 {
		t.Fatalf("strict grouped picture images = %d, want 5", got)
	}
}

func TestStrictDOCXFloatingVMLCanvasDoesNotBecomePictures(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "448024.docx")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// The canvas contains two VML image resources, but Word exposes it as an
	// msoCanvas rather than two Picture Shapes. The strict parser keeps the
	// six real main-story placements and suppresses both canvas children.
	if got := len(result.Images); got != 6 {
		t.Fatalf("strict DOCX floating VML canvas images = %d, want 6", got)
	}
}

func TestStrictDOCContentKeepsLegacyFieldResultAndFollowingText(t *testing.T) {
	input := "before \x13 HYPERLINK \"https://example.test\" \x14 visible link \x15 after"
	if got, want := stripLegacyWordFieldRanges(input), "before  visible link  after"; got != want {
		t.Fatalf("stripLegacyWordFieldRanges() = %q, want %q", got, want)
	}
}

func TestStrictDOCContentLeavesUnclosedLegacyFieldIntact(t *testing.T) {
	input := "before \x13 HYPERLINK \"https://example.test\" \x14 visible link after"
	if got, want := stripLegacyWordFieldRanges(input), "before  HYPERLINK \"https://example.test\"  visible link after"; got != want {
		t.Fatalf("unclosed field content was discarded: got %q, want %q", got, want)
	}
}

func TestStrictDOCTextHonorsDeclaredMainStoryCP(t *testing.T) {
	for _, name := range []string{"000151.doc", "002379.doc"} {
		file := filepath.Join("testdata", "web-samples", "samples", "doc", name)
		result, err := Extract(file, Options{StrictOfficeContent: true})
		if err != nil {
			t.Fatal(err)
		}
		// Word.Content is bounded by ccpText. These files retain text from later
		// stories in their complete piece tables; using that table wholesale
		// doubles the visible content compared with Word COM.
		if len(result.Text) > 700 {
			t.Fatalf("strict DOC %s escaped its main-story CP boundary: %d bytes", name, len(result.Text))
		}
	}
}

func TestStrictDOCTextRecoversEmptyMainStoryCLX(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "doc", "003157.doc")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	// The declared CLX carries no decodable main-story pieces, whereas Word
	// Content exposes the visible release text from the FIB-bounded range.
	if len(result.Text) < 1000 || !strings.Contains(result.Text, "Harpers Ferry") {
		t.Fatalf("strict DOC did not recover empty CLX main story: %d bytes", len(result.Text))
	}
}

func TestStrictDOCContentKeepsBareDateTableHeading(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "doc", "000113-2.doc")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	// The legacy Word table heading is literal visible text. Its DATE token is
	// not a field code and Word.Content.Text exposes it as part of the heading.
	if !strings.Contains(result.Text, "FINAL STATUS DATE of STATUS DETERMINATION") {
		t.Fatalf("strict DOC dropped visible DATE table heading")
	}
}

func TestStrictDOCContentRemovesHyperlinkWithCoexistingDateHeading(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "doc", "000113-3.doc")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(result.Text, "HYPERLINK") || strings.Contains(result.Text, "041207201_index.htm") {
		t.Fatal("strict DOC leaked a hidden hyperlink instruction")
	}
}

func TestStrictDOCContentOmitsAutoTextListFieldInstruction(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "doc", "001066.doc")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(result.Text, "AUTOTEXTLIST") {
		t.Fatal("strict DOC leaked AUTOTEXTLIST field instruction")
	}
}

func TestStrictDOCContentOmitsLowercaseSequenceFieldInstruction(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "doc", "001076-2.doc")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, unwanted := range []string{"seq Figure", "\\* Arabic"} {
		if strings.Contains(result.Text, unwanted) {
			t.Fatalf("strict DOC leaked sequence field instruction %q", unwanted)
		}
	}
}

func TestStrictDOCContentKeepsLongMainStoryWhenBareDateIsProse(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "doc", "005715.doc")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	// This large main story contains ordinary prose that happens to mention
	// DATE. It has no structurally delimited field around it. Treating that
	// word as an instruction made the legacy regex discard almost the entire
	// Word-visible report.
	if len(result.Text) < 60_000 || !strings.Contains(result.Text, "Air Travel Consumer Report") {
		t.Fatalf("strict DOC lost long visible main story: %d bytes", len(result.Text))
	}
}

func TestStrictDOCXContentCollapsesFormattingWhitespaceInsideSDT(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "dotnet__Open-XML-SDK__test_DocumentFormat.OpenXml.Tests.Assets_assets_TestDataStorage_v2FxTestFiles_ForTestCase_Bug242602_SDT_-_unknown.docx")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	// Word Content.Text collapses the formatting-only run boundary in this
	// content control: "SDT "+"Run" becomes one visible token.
	if !strings.Contains(result.Text, "sample paragraph with SDTRun inside") {
		t.Fatalf("strict DOCX did not mirror Word SDT run boundary: %q", result.Text)
	}
	if !strings.Contains(result.Text, "sample paragraph inside a SDT") {
		t.Fatalf("strict DOCX omitted SDT paragraph: %q", result.Text)
	}
}

func TestStrictDOCXContentKeepsTextAttachedWhitespaceInsideSDT(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "dotnet__Open-XML-SDK__test_DocumentFormat.OpenXml.Tests.Assets_assets_TestDataStorage_v2FxTestFiles_wordprocessing_content_control_SDT.docx")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	// Only formatting-only SDT runs are suppressed. The middle run here has
	// visible text, so Word keeps its XML-preserved word boundaries.
	if !strings.Contains(result.Text, "sample paragraph with SDT test Run inside") {
		t.Fatalf("strict DOCX lost visible SDT whitespace: %q", result.Text)
	}
	if got, want := strings.TrimSpace(result.Text), "This is a sample paragraph with SDT test Run inside.\nThis is a sample paragraph inside a SDT."; got != want {
		t.Fatalf("strict DOCX SDT text = %q, want Word-visible %q", got, want)
	}
}

func TestStrictDOCContentOmitsLegacyFormFieldPlaceholders(t *testing.T) {
	for _, name := range []string{"000131-2.doc", "000691-2.doc", "000693-2.doc"} {
		file := filepath.Join("testdata", "web-samples", "samples", "doc", name)
		result, err := Extract(file, Options{StrictOfficeContent: true})
		if err != nil {
			t.Fatal(err)
		}
		for _, placeholder := range []string{"FORMTEXT", "FORMCHECKBOX", "FORMDROPDOWN"} {
			if strings.Contains(result.Text, placeholder) {
				t.Fatalf("%s leaked invisible legacy form placeholder %q", name, placeholder)
			}
		}
	}
}

func TestStrictPPTXContentKeepsTemplateLabels(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "pptx", "00020964.pptx")
	result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"单击此处添加标题文字", "小标题", "报告人", "日期"} {
		if !strings.Contains(result.Text, want) {
			t.Fatalf("strict PPTX text missing %q: %q", want, result.Text)
		}
	}
}

func TestStrictPPTXContentIncludesGroupedShapeText(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "pptx", "00024860.pptx")
	result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(result.Text, "Total Consolidation") {
		t.Fatalf("strict PPTX text must include visible grouped text: %q", result.Text)
	}
}

func TestStrictPPTXContentIncludesTextBearingGraphicFrames(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "pptx", "00024860.pptx")
	result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(result.Text, "State of California") {
		t.Fatalf("strict PPTX text must include graphic-frame text: %q", result.Text)
	}
}

func TestStrictOfficeContentExcludesDOCXChartCache(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "LibreOffice__core__chart2_qa_extras_data_docx_data_point_inherited_color.docx")
	compatible, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(compatible.Text, "Column 1") {
		t.Fatalf("compatibility text lost chart cache: %q", compatible.Text)
	}
	strict, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if strict.Text != "" {
		t.Fatalf("strict document content = %q, want empty", strict.Text)
	}
}

func TestStrictOfficeContentExcludesEmbeddedChartWorkbook(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "LibreOffice__core__chart2_qa_extras_data_docx_3d-bar-label.docx")
	compatible, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(compatible.Text, "Series 3") {
		t.Fatalf("compatibility text lost embedded chart workbook: %q", compatible.Text)
	}
	strict, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if strict.Text != "" {
		t.Fatalf("strict document content = %q, want empty", strict.Text)
	}
}
