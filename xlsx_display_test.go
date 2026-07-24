package officeread

import (
	"encoding/binary"
	"path/filepath"
	"strings"
	"testing"
)

func TestXlsxGeneralNumberSuppressesStorageResidue(t *testing.T) {
	styles := xlsxCellStyles{numFmtIDs: []string{"0", "82"}}
	for _, style := range []int{0, 1} {
		if got := xlsxDisplayNumberForCell("42250557.5799999", style, styles); got != "42250557.58" {
			t.Fatalf("style %d display = %q, want %q", style, got, "42250557.58")
		}
	}
}

func TestXlsxGeneralNumberMatchesExcelSignificantDigits(t *testing.T) {
	styles := xlsxCellStyles{numFmtIDs: []string{"0"}}
	for input, want := range map[string]string{
		"0.14078153727589526":     "0.140782",
		"1023054.6259601419":      "1023055",
		"2.2871146412467191E-13":  "2.29E-13",
		"0.000022739604730348354": "2.27E-05",
	} {
		if got := xlsxDisplayNumberForCellWidth(input, 0, styles, 8); got != want {
			t.Fatalf("General display for %s = %q, want %q", input, got, want)
		}
	}
}

func TestStrictXLSXUsesVisibleDrawingImageOccurrences(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "xlsx", "00010274.xlsx")
	result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 5 {
		t.Fatalf("strict visible image occurrences = %d, want 5", len(result.Images))
	}
	for _, forbidden := range []string{"Comments", "http://", "https://"} {
		if strings.Contains(result.Text, forbidden) {
			t.Fatalf("strict workbook text unexpectedly contains %q", forbidden)
		}
	}
}

func TestStrictXLSXIncludesHiddenRowAndColumnUsedRangeValues(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "xlsx", "plutext__docx4j__xsd_docx4j_jaxb_packages.xlsx")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"wml.xsd or dml__ROOT", "pptx__ROOT.xsd"} {
		if !strings.Contains(result.Text, want) {
			t.Fatalf("strict Excel UsedRange content missing %q", want)
		}
	}
}

func TestXlsxCustomCurrencyDisplay(t *testing.T) {
	styles := xlsxCellStyles{
		numFmtIDs: []string{"164"},
		formats:   map[string]string{"164": `"$"#,##0.00`},
	}
	if got := xlsxDisplayNumberForCell("25.9", 0, styles); got != "$25.90" {
		t.Fatalf("currency display = %q, want %q", got, "$25.90")
	}
}

func TestXlsxCustomNumberRoundsLikeExcel(t *testing.T) {
	styles := xlsxCellStyles{
		numFmtIDs: []string{"164"},
		formats:   map[string]string{"164": "0.0"},
	}
	if got := xlsxDisplayNumberForCell("11111.25", 0, styles); got != "11111.3" {
		t.Fatalf("formatted number = %q, want %q", got, "11111.3")
	}
}

func TestXlsxBuiltInPercentDisplay(t *testing.T) {
	styles := xlsxCellStyles{numFmtIDs: []string{"10"}}
	if got := xlsxDisplayNumberForCell("2.5000000000000001E-2", 0, styles); got != "2.50%" {
		t.Fatalf("percent display = %q, want 2.50%%", got)
	}
}

func TestPptxStrictTextSeparatesFormattedRuns(t *testing.T) {
	data := []byte(`<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:sp><p:txBody><a:p><a:r><a:t>my</a:t></a:r><a:r><a:t>children</a:t></a:r></a:p></p:txBody></p:sp></p:sld>`)
	got, err := visiblePptxShapeText(data)
	if err != nil {
		t.Fatal(err)
	}
	if got != "my children" {
		t.Fatalf("formatted runs = %q, want %q", got, "my children")
	}
}

func TestPptxStrictTextKeepsSplitCitationNumber(t *testing.T) {
	data := []byte(`<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:sp><p:txBody><a:p><a:r><a:t>2007</a:t></a:r><a:r><a:t>1</a:t></a:r></a:p></p:txBody></p:sp></p:sld>`)
	got, err := visiblePptxShapeText(data)
	if err != nil {
		t.Fatal(err)
	}
	if got != "20071" {
		t.Fatalf("split citation = %q, want %q", got, "20071")
	}
}

func TestVisibleWebURLIsNotDiscardedAsBinaryNoise(t *testing.T) {
	const value = "www.mngeo.state.mn.us/workgroup/metadata/index.html"
	if got := cleanMarkdownVisibleText(value); got != value {
		t.Fatalf("visible URL = %q, want %q", got, value)
	}
	if !looksLikeWebURLText(value) {
		t.Fatal("visible URL was not recognized")
	}
}

func TestLegacyWordStrictModeKeepsTextAcrossPieceBoundary(t *testing.T) {
	word := []byte("anomaly")
	plc := make([]byte, 28)
	binary.LittleEndian.PutUint32(plc[0:], 0)
	binary.LittleEndian.PutUint32(plc[4:], 4)
	binary.LittleEndian.PutUint32(plc[8:], 7)
	// Compressed piece FCPRMs: 0x40000000 marks compressed text and the
	// remaining value is the byte offset shifted left once.
	binary.LittleEndian.PutUint32(plc[14:], 0x40000000)
	binary.LittleEndian.PutUint32(plc[22:], 0x40000008)
	parts := parseWordPieceTableTextUntilCP(word, plc, ^uint32(0), true)
	if got := strings.Join(parts, "\n"); got != "anomaly" {
		t.Fatalf("piece-boundary text = %q, want %q", got, "anomaly")
	}
}

func TestLegacyWordStrictModePreservesControlCharacterWordBoundary(t *testing.T) {
	// Word.Content.Text represents C0 field/object boundaries as whitespace.
	// Strict extraction must preserve that boundary rather than letting generic
	// text cleanup join the two visible fragments.
	word := []byte("packag\x01ing")
	plc := make([]byte, 16)
	binary.LittleEndian.PutUint32(plc[0:], 0)
	binary.LittleEndian.PutUint32(plc[4:], uint32(len(word)))
	binary.LittleEndian.PutUint32(plc[10:], 0x40000000)
	parts := parseWordPieceTableTextUntilCP(word, plc, ^uint32(0), true)
	if got := strings.Join(parts, "\n"); got != "packag ing" {
		t.Fatalf("control-character boundary = %q, want %q", got, "packag ing")
	}
}
