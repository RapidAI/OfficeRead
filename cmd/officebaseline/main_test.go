package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"officeread"
)

func TestOrderedTokenOverlapDetectsReordering(t *testing.T) {
	matched, reference, candidate, available := orderedTokenOverlap("one two three", "three two one")
	if !available || reference != 3 || candidate != 3 || matched != 1 {
		t.Fatalf("orderedTokenOverlap = (%d, %d, %d, %v), want (1, 3, 3, true)", matched, reference, candidate, available)
	}
}

func TestImageQualityComparisonUsesDecodedPixelsRatherThanContainerBytes(t *testing.T) {
	imageData := func(c color.Color) []byte {
		t.Helper()
		var buf bytes.Buffer
		canvas := image.NewRGBA(image.Rect(0, 0, 2, 1))
		canvas.Set(0, 0, c)
		canvas.Set(1, 0, c)
		if err := png.Encode(&buf, canvas); err != nil {
			t.Fatal(err)
		}
		return buf.Bytes()
	}
	path := filepath.Join(t.TempDir(), "office-export.png")
	red := imageData(color.RGBA{R: 255, A: 255})
	if err := os.WriteFile(path, red, 0o644); err != nil {
		t.Fatal(err)
	}
	available, match, note := imageQualityComparison([]string{path}, []officeread.Image{{Data: append([]byte(nil), red...)}})
	if !available || !match || !strings.Contains(note, "match") {
		t.Fatalf("same pixels = (%v, %v, %q), want available match", available, match, note)
	}
	blue := imageData(color.RGBA{B: 255, A: 255})
	available, match, note = imageQualityComparison([]string{path}, []officeread.Image{{Data: blue}})
	if available || match || !strings.Contains(note, "not a valid") {
		t.Fatalf("different pixels = (%v, %v, %q), want unavailable supplemental comparison", available, match, note)
	}
}

func TestImageQualityComparisonLeavesUndecodableFormatsAsCountOnly(t *testing.T) {
	available, match, note := imageQualityComparison([]string{"not-used"}, []officeread.Image{{Data: []byte("not an image")}})
	if available || match || !strings.Contains(note, "could not be read") {
		t.Fatalf("undecodable result = (%v, %v, %q), want unavailable", available, match, note)
	}
}

func TestReportDocumentsImageQualityLimitations(t *testing.T) {
	r := newReport([]string{"sample.pptx"})
	data, err := json.Marshal(r)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(data), "Image count parity is mandatory") {
		t.Fatalf("report omitted image-quality limitation: %s", data)
	}
	if !strings.Contains(string(data), "Excel ranges above excel-max-cells") || !strings.Contains(string(data), "baseline-unavailable result") {
		t.Fatalf("report omitted baseline limitations: %s", data)
	}
}

func TestTokenStreamNormalizesWordFieldBoundaryConcatenation(t *testing.T) {
	matched, reference, candidate, missing, extra := tokenOverlap(
		"Release No. 0043.04Office at http://example.test",
		"Release No. 0043.04 Office athttp://example.test",
	)
	if matched != reference || reference != candidate || len(missing) != 0 || len(extra) != 0 {
		t.Fatalf("Word field-boundary overlap = (%d, %d, %d, %#v, %#v), want exact match", matched, reference, candidate, missing, extra)
	}
}

func TestOrderedTokenOverlapSkipsLargeInputs(t *testing.T) {
	text := "x "
	for i := 0; i < 2500; i++ {
		text += "x "
	}
	_, reference, candidate, available := orderedTokenOverlap(text, text)
	if available || reference != 2501 || candidate != 2501 {
		t.Fatalf("large ordered comparison = (%d, %d, %v), want (2501, 2501, false)", reference, candidate, available)
	}
}

func TestFormulaTokenStreamPreservesSymbolsWithoutLayoutGrouping(t *testing.T) {
	got := formulaTokenStream("𝑓\n𝑥=a₀+n=1∞")
	want := []string{"f", "x", "a", "0", "n", "1"}
	if len(got) != len(want) {
		t.Fatalf("formula tokens = %#v, want %#v", got, want)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("formula tokens = %#v, want %#v", got, want)
		}
	}
}

func TestEquationDocumentUsesFormulaComparison(t *testing.T) {
	file := filepath.Join(t.TempDir(), "arbitrary-name.docx")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	w, err := zw.Create("word/document.xml")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := w.Write([]byte(`<w:document xmlns:w="urn:w" xmlns:m="urn:m"><m:oMath/></w:document>`)); err != nil {
		t.Fatal(err)
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	if got := comparisonModeForPath(file); got != comparisonModeFormula {
		t.Fatalf("equation fixture comparison mode = %q, want %q", got, comparisonModeFormula)
	}
}

func TestExcelPaths(t *testing.T) {
	if !excelPaths([]string{"one.xls", "two.xlsx"}) {
		t.Fatal("Excel paths should request isolation")
	}
	if excelPaths([]string{"one.xlsx", "two.docx"}) {
		t.Fatal("mixed paths must not be treated as an Excel batch")
	}
}

func TestDiagnose(t *testing.T) {
	cases := []struct {
		name string
		in   fileResult
		want string
	}{
		{"aligned", fileResult{F1: 1, OrderedComparisonAvailable: true, OrderedF1: 1, ImageMatch: true}, "aligned"},
		{"layout-order", fileResult{F1: 1, OrderedComparisonAvailable: true, OrderedF1: .9, ImageMatch: true}, "aligned"},
		{"text", fileResult{F1: .9, ImageMatch: true}, "text-mismatch"},
		{"image", fileResult{F1: 1, ImageMatch: false}, "image-mismatch"},
		{"both", fileResult{F1: .9, ImageMatch: false}, "text-and-image-mismatch"},
		{"stored-value", fileResult{F1: .9, ImageMatch: false, ComparisonScope: comparisonScopeOfficeStored}, "baseline-scope-mismatch"},
		{"office", fileResult{BaselineStatus: "baseline-unavailable", Error: "timeout"}, "office-baseline-unavailable"},
	}
	for _, tc := range cases {
		if got := diagnose(tc.in); got != tc.want {
			t.Errorf("%s: diagnose() = %q, want %q", tc.name, got, tc.want)
		}
	}
}

func TestComparisonScopeForOfficeSource(t *testing.T) {
	if got := comparisonScopeForOfficeSource("Excel.visible-worksheets.UsedRange.Value2"); got != comparisonScopeOfficeStored {
		t.Fatalf("Value2 source scope = %q, want %q", got, comparisonScopeOfficeStored)
	}
	if got := comparisonScopeForOfficeSource("Excel.visible-worksheets.UsedRange.Text"); got != comparisonScopeOfficeVisible {
		t.Fatalf("Text source scope = %q, want %q", got, comparisonScopeOfficeVisible)
	}
}

func TestSummaryDiagnosisCounts(t *testing.T) {
	s := summary{}
	s = add(s, fileResult{F1: 1, ImageMatch: true})
	s = add(s, fileResult{F1: .9, ImageMatch: false})
	s = add(s, fileResult{BaselineStatus: "baseline-unavailable", Error: "timeout"})
	for kind, want := range map[string]int{"aligned": 1, "text-and-image-mismatch": 1, "office-baseline-unavailable": 1} {
		if got := s.DiagnosisCounts[kind]; got != want {
			t.Errorf("DiagnosisCounts[%q] = %d, want %d", kind, got, want)
		}
	}
}

func TestQualityGateExcludesStoredValueBaselineAndRequiresExactParity(t *testing.T) {
	g := qualityGate{}
	g = addQualityGate(g, fileResult{Path: "aligned.pptx", BaselineStatus: "compared", ComparisonScope: comparisonScopeOfficeVisible, F1: 1, ImageMatch: true})
	g = addQualityGate(g, fileResult{Path: "text-only.docx", BaselineStatus: "compared", ComparisonScope: comparisonScopeOfficeVisible, F1: .999, ImageMatch: true})
	g = addQualityGate(g, fileResult{Path: "large.xlsx", BaselineStatus: "compared", ComparisonScope: comparisonScopeOfficeStored, F1: .8, ImageMatch: false})
	g = addQualityGate(g, fileResult{Path: "unavailable.ppt", BaselineStatus: "baseline-unavailable", Error: "timeout"})
	g = finalizeQualityGate(g)
	if g.Compared != 3 || g.BaselineUnavailable != 1 || g.ContentCompared != 2 {
		t.Fatalf("quality gate counts = %#v", g)
	}
	if g.ContentTextMatches != 1 || g.ContentImageMatches != 2 || g.ContentFullyAligned != 1 {
		t.Fatalf("quality gate parity = %#v", g)
	}
	if g.ContentFullyAlignedRate != .5 || len(g.ExcludedScopeMismatchFiles) != 1 || g.ExcludedScopeMismatchFiles[0] != "large.xlsx" {
		t.Fatalf("quality gate final = %#v", g)
	}
}

func TestWordBaselineUsesFinalRevisionView(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_baseline.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(data), "RevisionsView = 0") {
		t.Fatal("Word COM baseline must select wdRevisionsViewFinal (0)")
	}
}

func TestOfficeBaselinePreservesControlCharacterWordBoundaries(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_baseline.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(data), "-replace '\\p{Cc}', ' '") {
		t.Fatal("Office baseline must replace non-visible control characters with spaces")
	}
}

func TestExcelBaselineUsesConfigurableVisibleCellLimit(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_baseline.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	if !strings.Contains(script, "[int]$ExcelMaxCells = 200000") || !strings.Contains(script, "$cellCount -gt $ExcelMaxCells") {
		t.Fatal("Excel COM baseline must expose its visible-cell threshold as ExcelMaxCells")
	}
}
