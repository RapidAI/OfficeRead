package officeread

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"os"
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

func TestStrictXlsxContentExcludesDrawingAndVMLControlText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" sheetId="1" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rId1"/><legacyDrawing r:id="rId2"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="../drawings/drawing1.xml"/><Relationship Id="rId2" Target="../drawings/vmlDrawing1.vml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:pic><xdr:txBody>drawing residue</xdr:txBody></xdr:pic></xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/vmlDrawing1.vml", `<xml xmlns:v="urn:vml" xmlns:x="urn:x"><v:shape><x:ClientData ObjectType="Pict"/></v:shape><v:textbox>vml residue</v:textbox></xml>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "drawing-only.xlsx")
	if err := os.WriteFile(path, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	result, err := Extract(path, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if result.Text != "Visible" {
		t.Fatalf("strict worksheet text = %q, want only the visible sheet name", result.Text)
	}
}

func TestStrictXlsxContentExcludesChartSheetCache(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Chart" sheetId="2" r:id="rId2"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="chartsheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>visible cell</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/chartsheets/sheet1.xml", `<chartsheet><drawing><v>chart cache 987654</v></drawing></chartsheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "chart-sheet.xlsx")
	if err := os.WriteFile(path, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	result, err := Extract(path, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if result.Text != "Data\nvisible cell" {
		t.Fatalf("strict workbook text = %q, want only worksheet name and cell", result.Text)
	}
}

func TestXlsxGeneralNumberUsesExcelElevenSignificantDigits(t *testing.T) {
	styles := xlsxCellStyles{numFmtIDs: []string{"0"}}
	if got := xlsxDisplayNumberForCell("142.85714285714286", 0, styles); got != "142.85714286" {
		t.Fatalf("General display = %q, want %q", got, "142.85714286")
	}
	if got := xlsxDisplayNumberForCell("120", 0, styles); got != "120" {
		t.Fatalf("General integer display = %q, want %q", got, "120")
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

func TestXlsxGeneralNumberWideColumnKeepsDecimalGeneralDisplay(t *testing.T) {
	styles := xlsxCellStyles{numFmtIDs: []string{"0"}}
	if got := xlsxDisplayNumberForCellWidth("142.85714285714286", 0, styles, 9); got != "142.85714286" {
		t.Fatalf("wide General display = %q, want %q", got, "142.85714286")
	}
	if got := xlsxDisplayNumberForCellWidth("120", 0, styles, 9); got != "120" {
		t.Fatalf("wide General integer = %q, want %q", got, "120")
	}
}

func TestXlsxGeneralNumberNarrowColumnHonorsElevenDigitCap(t *testing.T) {
	styles := xlsxCellStyles{numFmtIDs: []string{"0"}}
	if got := xlsxDisplayNumberForCellWidth("0.040140383383104", 0, styles, 11); got != "0.04014038" {
		t.Fatalf("narrow General significant-digit cap = %q, want %q", got, "0.04014038")
	}
}

func TestXlsxGeneralNumberWidthNineKeepsThreeDigitIntegerDecimal(t *testing.T) {
	styles := xlsxCellStyles{numFmtIDs: []string{"0"}}
	if got := xlsxDisplayNumberForCellWidth("120", 0, styles, 8); got != "120" {
		t.Fatalf("width-eight General integer = %q, want %q", got, "120")
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

func TestStrictXLSXExcludesVMLOnlyPictureFromExcelShapes(t *testing.T) {
	// These workbooks retain a VML image relationship, but Excel's worksheet
	// Shapes collection does not expose it as a Picture Shape. Strict Office
	// mode must not revive it merely because it is present in xl/media.
	for _, name := range []string{"00010505.xlsx", "00010542.xlsx", "00010568.xlsx"} {
		file := filepath.Join("testdata", "web-samples", "samples", "xlsx", name)
		result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
		if err != nil {
			t.Fatalf("%s: %v", name, err)
		}
		if len(result.Images) != 0 {
			t.Fatalf("%s strict visible picture occurrences = %d, want 0", name, len(result.Images))
		}
	}
}

func TestStrictXLSXExcludesVMLOLEPreviewPictures(t *testing.T) {
	// The VML shapes look like PictureFrames, but the worksheet explicitly
	// declares their IDs as MSPhotoEd OLE objects. Excel's Shapes collection
	// reports no msoPicture entries for them.
	file := filepath.Join("testdata", "web-samples", "samples", "xlsx", "00016528.xlsx")
	result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 0 {
		t.Fatalf("strict visible OLE-preview picture occurrences = %d, want 0", len(result.Images))
	}
}

func TestStrictXLSXIncludesHiddenRowUsedRangeValues(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "xlsx", "plutext__docx4j__xsd_docx4j_jaxb_packages.xlsx")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	// The fixture's asserted values are in hidden column A. Keep the legacy
	// strict UsedRange contract explicit: cells in hidden columns remain part
	// of the worksheet range source.
	for _, want := range []string{"wml.xsd or dml__ROOT", "pptx__ROOT.xsd"} {
		if !strings.Contains(result.Text, want) {
			t.Fatalf("strict Excel UsedRange content missing %q", want)
		}
	}
}

func TestStrictXLSXIncludesHiddenColumnAndRowCellText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet><cols><col min="1" max="1" hidden="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>hidden column</t></is></c><c r="B1" t="inlineStr"><is><t>visible column</t></is></c></row><row r="2" hidden="1"><c r="B2" t="inlineStr"><is><t>hidden row</t></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "hidden-column.xlsx")
	if err := os.WriteFile(path, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	result, err := Extract(path, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if result.Text != "Data\nhidden column\nvisible column\nhidden row" {
		t.Fatalf("strict hidden-column text = %q", result.Text)
	}
}

func TestStrictXLSXExcludesDataValidationPromptText(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "xlsx", "00015904.xlsx")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, forbidden := range []string{"ERROR DE DATO", "LISTA DE PROVEEDORES", "SELECCIONAR LA EMPRESA"} {
		if strings.Contains(result.Text, forbidden) {
			t.Fatalf("strict Excel text unexpectedly contains validation annotation %q", forbidden)
		}
	}
}

func TestStrictXLSXExcludesDataBarCellsWhoseValuesAreHidden(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "xlsx", "closedxml__closedxml__ClosedXML.Tests_Resource_Examples_ConditionalFormatting_CFDataBar.xlsx")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, forbidden := range []string{"\n1\n", "\n2\n", "\n3\n"} {
		if strings.Contains("\n"+result.Text+"\n", forbidden) {
			t.Fatalf("strict Excel text unexpectedly contains hidden data-bar value %q", forbidden)
		}
	}
}

func TestStrictXLSXExcludesIconSetCellsWhoseValuesAreHidden(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "xlsx", "closedxml__closedxml__ClosedXML.Tests_Resource_Examples_ConditionalFormatting_CFIconSet.xlsx")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, forbidden := range []string{"\n1\n", "\n2\n", "\n3\n"} {
		if strings.Contains("\n"+result.Text+"\n", forbidden) {
			t.Fatalf("strict Excel text unexpectedly contains hidden icon-set value %q", forbidden)
		}
	}
}

func TestStrictXLSXPreservesStopIfTrueCellBeforeHiddenIconSet(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "xlsx", "closedxml__closedxml__ClosedXML.Tests_Resource_Examples_ConditionalFormatting_CFStopIfTrue.xlsx")
	// Keep the parser-level predicate auditable as this fixture combines a
	// stopIfTrue comparison with a later showValue=0 icon set.
	zr, err := zip.OpenReader(file)
	if err != nil {
		t.Fatal(err)
	}
	defer zr.Close()
	var data []byte
	for _, entry := range zr.File {
		if entry.Name == "xl/worksheets/sheet1.xml" {
			data, err = readZipFile(entry)
			if err != nil {
				t.Fatal(err)
			}
			break
		}
	}
	if len(data) == 0 {
		t.Fatal("worksheet XML missing from fixture")
	}
	if hidden := xlsxHiddenConditionalFormatValueCells(data); hidden["A1"] || !hidden["A2"] || !hidden["A3"] || !hidden["A4"] {
		t.Fatalf("conditional visibility = %#v, want A1 visible and A2:A4 hidden", hidden)
	}
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains("\n"+result.Text+"\n", "\n6\n") {
		t.Fatalf("strict Excel text missing stopIfTrue-visible value: %q", result.Text)
	}
	for _, hidden := range []string{"\n1\n", "\n2\n", "\n3\n"} {
		if strings.Contains("\n"+result.Text+"\n", hidden) {
			t.Fatalf("strict Excel text unexpectedly contains icon-set value not matched by stopIfTrue: %q", result.Text)
		}
	}
}

func TestXlsxCellIsStopIfTrueMatch(t *testing.T) {
	for _, test := range []struct {
		value    float64
		operator string
		formulas []string
		want     bool
	}{
		{6, "greaterThan", []string{"5"}, true},
		{2, "greaterThan", []string{"5"}, false},
		{2, "between", []string{"1", "3"}, true},
		{4, "between", []string{"1", "3"}, false},
		{4, "notBetween", []string{"1", "3"}, true},
		{4, "greaterThan", []string{"A1"}, false},
	} {
		if got := xlsxCellIsStopIfTrueMatch(test.value, test.operator, test.formulas); got != test.want {
			t.Errorf("match(%v, %q, %v) = %v, want %v", test.value, test.operator, test.formulas, got, test.want)
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

func TestXlsxCustomFormatOverridesBuiltInCurrencyID(t *testing.T) {
	styles := xlsxCellStyles{
		numFmtIDs: []string{"6"},
		formats:   map[string]string{"6": `"$"#,##0_);[Red]\("$"#,##0\)`},
	}
	if got := xlsxDisplayNumberForCell("397402", 0, styles); got != "$397,402" {
		t.Fatalf("custom built-in currency display = %q, want %q", got, "$397,402")
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

func TestXlsxAccountingNumberFormatsMatchVisibleDigits(t *testing.T) {
	tests := []struct {
		format string
		value  string
		want   string
	}{
		{`_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`, "1705.8228000000001", "1,705.82"},
		{`_($* #,##0.00_);_($* \(#,##0.00\);_($* "-"??_);_(@_)`, "-72.188000000000002", "($72.19)"},
		{`"$"#,##0`, "16500", "$16,500"},
	}
	for _, tc := range tests {
		styles := xlsxCellStyles{numFmtIDs: []string{"164"}, formats: map[string]string{"164": tc.format}}
		if got := xlsxDisplayNumberForCell(tc.value, 0, styles); got != tc.want {
			t.Fatalf("format %q, value %s = %q, want %q", tc.format, tc.value, got, tc.want)
		}
	}
}

func TestXlsxAccountingFormatKeepsQuotedCurrencyLiteral(t *testing.T) {
	styles := xlsxCellStyles{
		numFmtIDs: []string{"44"},
		formats: map[string]string{
			"44": `"B/." #,##0.00`,
		},
	}
	if got := xlsxDisplayNumberForCell("300", 0, styles); got != "B/. 300.00" {
		t.Fatalf("quoted accounting currency display = %q, want %q", got, "B/. 300.00")
	}
}

func TestXlsxCustomPercentDisplay(t *testing.T) {
	styles := xlsxCellStyles{
		numFmtIDs: []string{"164"},
		formats:   map[string]string{"164": "0.0%"},
	}
	if got := xlsxDisplayNumberForCell("0.11137657804985834", 0, styles); got != "11.1%" {
		t.Fatalf("custom percent display = %q, want 11.1%%", got)
	}
}

func TestXlsxBuiltInPercentDisplay(t *testing.T) {
	styles := xlsxCellStyles{numFmtIDs: []string{"10"}}
	if got := xlsxDisplayNumberForCell("2.5000000000000001E-2", 0, styles); got != "2.50%" {
		t.Fatalf("percent display = %q, want 2.50%%", got)
	}
}

func TestXlsxBuiltInCommaNumberFormats(t *testing.T) {
	styles := xlsxCellStyles{numFmtIDs: []string{"37", "38", "39", "40"}}
	for _, tc := range []struct {
		value string
		style int
		want  string
	}{
		{"14077", 0, "14,077"},
		{"-14077", 1, "(14,077)"},
		{"14077.125", 2, "14,077.13"},
		{"-14077.125", 3, "(14,077.13)"},
	} {
		if got := xlsxDisplayNumberForCell(tc.value, tc.style, styles); got != tc.want {
			t.Fatalf("built-in comma style %d, value %s = %q, want %q", tc.style, tc.value, got, tc.want)
		}
	}
}

func TestXlsxBuiltInTimeDisplays(t *testing.T) {
	styles := xlsxCellStyles{numFmtIDs: []string{"20", "46", "45", "47"}}
	for _, tc := range []struct {
		value string
		style int
		want  string
	}{
		{"0.5", 0, "12:00"},
		{"1.5", 1, "36:00:00"},
		{"0.5", 2, "00:00"},
		{"0.5", 3, "00:00.0"},
	} {
		if got := xlsxDisplayNumberForCell(tc.value, tc.style, styles); got != tc.want {
			t.Fatalf("style %d, value %s = %q, want %q", tc.style, tc.value, got, tc.want)
		}
	}
}

func TestXlsxBuiltInDateDisplay(t *testing.T) {
	styles := xlsxCellStyles{numFmtIDs: []string{"14", "15", "16", "17"}}
	for _, tc := range []struct {
		style int
		want  string
	}{
		{style: 0, want: "1/1/24"},
		{style: 1, want: "1-Jan-24"},
		{style: 2, want: "1-Jan"},
		{style: 3, want: "Jan-24"},
	} {
		if got := xlsxDisplayNumberForCell("45292", tc.style, styles); got != tc.want {
			t.Fatalf("built-in date style %d = %q, want %q", tc.style, got, tc.want)
		}
	}
}

func TestXlsxBuiltInAMPMTimeDisplays(t *testing.T) {
	styles := xlsxCellStyles{numFmtIDs: []string{"18", "19"}}
	for _, tc := range []struct {
		value string
		style int
		want  string
	}{
		{"0.5", 0, "12:00 PM"},
		{"0.005150462962963", 1, "12:07:25 AM"},
	} {
		if got := xlsxDisplayNumberForCell(tc.value, tc.style, styles); got != tc.want {
			t.Fatalf("style %d, value %s = %q, want %q", tc.style, tc.value, got, tc.want)
		}
	}
}

func TestXlsxCustomDateDisplay(t *testing.T) {
	styles := xlsxCellStyles{
		numFmtIDs: []string{"164"},
		formats:   map[string]string{"164": "dd\\-mmm\\-yy"},
	}
	if got := xlsxDisplayNumberForCell("41014", 0, styles); got != "15-Apr-12" {
		t.Fatalf("custom date display = %q, want 15-Apr-12", got)
	}
}

func TestXlsxCustomWeekdayDateDisplay(t *testing.T) {
	styles := xlsxCellStyles{
		numFmtIDs: []string{"164"},
		formats:   map[string]string{"164": "ddd\\-dd\\-mmm\\-yy"},
	}
	if got := xlsxDisplayNumberForCell("40240", 0, styles); got != "Wed-03-Mar-10" {
		t.Fatalf("weekday date display = %q, want %q", got, "Wed-03-Mar-10")
	}
}

func TestXlsxCustomNumericMonthFirstDateDisplay(t *testing.T) {
	styles := xlsxCellStyles{
		numFmtIDs: []string{"164", "165"},
		formats:   map[string]string{"164": "m/d/yy;@", "165": "m/d/yyyy;@"},
	}
	if got := xlsxDisplayNumberForCell("27906", 0, styles); got != "5/26/76" {
		t.Fatalf("numeric month-first date display = %q, want %q", got, "5/26/76")
	}
	if got := xlsxDisplayNumberForCell("27906", 1, styles); got != "5/26/1976" {
		t.Fatalf("numeric long month-first date display = %q, want %q", got, "5/26/1976")
	}
}

func TestXlsxCustomTimeDisplay(t *testing.T) {
	styles := xlsxCellStyles{
		numFmtIDs: []string{"164", "165", "166"},
		formats: map[string]string{
			"164": "h:mm:ss;@",
			"165": "hh:mm:ss;@",
			"166": "[h]:mm:ss",
		},
	}
	if got := xlsxDisplayNumberForCell("0.03472222222222221", 0, styles); got != "0:50:00" {
		t.Fatalf("custom h:mm:ss display = %q, want %q", got, "0:50:00")
	}
	if got := xlsxDisplayNumberForCell("0.5", 1, styles); got != "12:00:00" {
		t.Fatalf("custom hh:mm:ss display = %q, want %q", got, "12:00:00")
	}
	if got := xlsxDisplayNumberForCell("1.5", 2, styles); got != "36:00:00" {
		t.Fatalf("custom elapsed time display = %q, want %q", got, "36:00:00")
	}
}

func TestXlsxCustomElapsedMinutesAndAMPMTimeDisplay(t *testing.T) {
	styles := xlsxCellStyles{
		numFmtIDs: []string{"164", "165"},
		formats: map[string]string{
			"164": "[mm]:ss",
			"165": "[$-F400]h:mm:ss\\ AM/PM",
		},
	}
	if got := xlsxDisplayNumberForCell("0.018831018518518518", 0, styles); got != "27:07" {
		t.Fatalf("custom elapsed minutes display = %q, want %q", got, "27:07")
	}
	if got := xlsxDisplayNumberForCell("0.5", 1, styles); got != "12:00:00 PM" {
		t.Fatalf("custom AM/PM display = %q, want %q", got, "12:00:00 PM")
	}
}

func TestXlsxCustomDateDisplayOverridesBuiltInStyleID(t *testing.T) {
	// OOXML permits producers to assign a custom format to an ID in the
	// built-in range. Excel honors formatCode, not the numeric range alone.
	styles := xlsxCellStyles{
		numFmtIDs: []string{"178"},
		formats:   map[string]string{"178": "d"},
	}
	if got := xlsxDisplayNumberForCell("40547", 0, styles); got != "4" {
		t.Fatalf("custom built-in-ID date display = %q, want %q", got, "4")
	}
}

func TestXlsxCustomLongDateKeepsEscapedSpaces(t *testing.T) {
	styles := xlsxCellStyles{
		numFmtIDs: []string{"164"},
		formats:   map[string]string{"164": `[$-1409]d\ mmmm\ yyyy;@`},
	}
	if got := xlsxDisplayNumberForCell("45292", 0, styles); got != "1 January 2024" {
		t.Fatalf("custom long date display = %q, want %q", got, "1 January 2024")
	}
}

func TestStrictXLSXCellTextDecodesEscapedControlCharacters(t *testing.T) {
	// SpreadsheetML escapes C0 controls as _xNNNN_. Excel renders those as no
	// printable cell glyph; retaining the source spelling produces bogus tokens
	// (for example _x001A_ becomes "x001a") in a strict COM comparison.
	if got := xlsxStrictOfficeCellText("_x001A_"); got != "" {
		t.Fatalf("escaped control cell text = %q, want empty", got)
	}
	if got := xlsxStrictOfficeCellText("name_x000A_value"); got != "name\nvalue" {
		t.Fatalf("escaped newline cell text = %q, want newline", got)
	}
}

func TestStrictXLSXCellTextKeepsOrdinaryLiterals(t *testing.T) {
	const literal = "wml.xsd or dml__ROOT"
	if got := xlsxStrictOfficeCellText(literal); got != literal {
		t.Fatalf("strict cell literal = %q, want %q", got, literal)
	}
}

func TestStrictXLSXCellTextKeepsLiteralHash(t *testing.T) {
	// General Word-oriented text cleanup recognizes a leading hash as part of
	// a field instruction.  A SpreadsheetML cell is not a Word field: Excel
	// Range.Text exposes the hash glyph verbatim.
	if got := xlsxStrictOfficeCellText("#"); got != "#" {
		t.Fatalf("strict cell hash = %q, want #", got)
	}
}

func TestPptxStrictTextConcatenatesFormattedRuns(t *testing.T) {
	data := []byte(`<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:sp><p:txBody><a:p><a:r><a:t>my</a:t></a:r><a:r><a:t>children</a:t></a:r></a:p></p:txBody></p:sp></p:sld>`)
	got, err := visiblePptxShapeText(data)
	if err != nil {
		t.Fatal(err)
	}
	if got != "mychildren" {
		t.Fatalf("formatted runs = %q, want %q", got, "mychildren")
	}
}

func TestPptxStrictTextKeepsAuthoredSpaceAcrossFormattedRuns(t *testing.T) {
	data := []byte(`<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:sp><p:txBody><a:p><a:r><a:t>my</a:t></a:r><a:r><a:t> children</a:t></a:r></a:p></p:txBody></p:sp></p:sld>`)
	got, err := visiblePptxShapeText(data)
	if err != nil {
		t.Fatal(err)
	}
	if got != "my children" {
		t.Fatalf("authored-space runs = %q, want %q", got, "my children")
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
