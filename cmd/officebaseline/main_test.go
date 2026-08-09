package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"image"
	"image/color"
	"image/png"
	"os"
	"os/exec"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
	"time"

	"github.com/RapidAI/OfficeRead"
)

func TestAddFileSkipsOfficeOwnerLockFile(t *testing.T) {
	files := []string{}
	counts := map[string]int{}
	addFile(&files, counts, `C:\\samples\\~$report.docx`, 1)
	addFile(&files, counts, `C:\\samples\\report.docx`, 1)
	if got, want := files, []string{`C:\\samples\\report.docx`}; !reflect.DeepEqual(got, want) {
		t.Fatalf("files = %#v, want %#v", got, want)
	}
	if got := counts[".docx"]; got != 1 {
		t.Fatalf(".docx count = %d, want 1", got)
	}
}

func TestOrderedTokenOverlapDetectsReordering(t *testing.T) {
	matched, reference, candidate, available := orderedTokenOverlap("one two three", "three two one")
	if !available || reference != 3 || candidate != 3 || matched != 1 {
		t.Fatalf("orderedTokenOverlap = (%d, %d, %d, %v), want (1, 3, 3, true)", matched, reference, candidate, available)
	}
}

func TestWaitCommandReturnsAfterKillGrace(t *testing.T) {
	// The child ignores neither termination nor Wait: it simply remains alive
	// beyond the timeout. The helper must still return promptly so one wedged
	// COM automation call cannot block a full-corpus checkpoint run.
	cmd := exec.Command("powershell", "-NoProfile", "-NonInteractive", "-Command", "Start-Sleep -Seconds 10")
	if err := cmd.Start(); err != nil {
		t.Skipf("PowerShell unavailable: %v", err)
	}
	started := time.Now()
	err := waitCommand(cmd, 20*time.Millisecond, 20*time.Millisecond, "test child")
	if err == nil || !strings.Contains(err.Error(), "timed out") {
		t.Fatalf("waitCommand error = %v, want timeout", err)
	}
	if elapsed := time.Since(started); elapsed > time.Second {
		t.Fatalf("waitCommand took %s, want bounded return", elapsed)
	}
}

func TestKillProcessTreeStartsWindowsTreeKillerBeforeRootKill(t *testing.T) {
	// The ordering is important for a timed-out COM child: once the direct
	// PowerShell process exits, Windows can re-parent its broker descendants and
	// taskkill /T can no longer reach them.  Keep this source-level regression
	// check platform-neutral; the actual command is only executed on Windows.
	path := filepath.Join("main.go")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	source := string(data)
	start := strings.Index(source, "func killProcessTree")
	if start < 0 {
		t.Fatal("could not locate killProcessTree")
	}
	remainder := source[start:]
	end := strings.Index(remainder, "// waitCommand places")
	if end < 0 {
		t.Fatal("could not locate end of killProcessTree")
	}
	body := remainder[:end]
	startTaskkill := strings.Index(body, "taskkill.Start()")
	killRoot := strings.Index(body, "cmd.Process.Kill()")
	if startTaskkill < 0 || killRoot < 0 || startTaskkill > killRoot {
		t.Fatal("killProcessTree must start taskkill before terminating the root process")
	}
}

func TestKillProcessTreeAllowsTaskkillToTraverseBeforeRootKill(t *testing.T) {
	path := filepath.Join("main.go")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	text := string(data)
	start := strings.Index(text, "if taskkill := exec.Command")
	if start < 0 {
		t.Fatal("taskkill branch missing")
	}
	branch := text[start:]
	wait := strings.Index(branch, "time.After(750 * time.Millisecond)")
	killRoot := strings.Index(branch, "cmd.Process.Kill()")
	if wait < 0 || killRoot < 0 || wait > killRoot {
		t.Fatal("taskkill must receive a bounded traversal window before root kill")
	}
}

func TestWriteJSONReplacesExistingCheckpoint(t *testing.T) {
	path := filepath.Join(t.TempDir(), "checkpoint.json")
	first := newReport([]string{"first.doc"})
	first.Files = []fileResult{{Path: "first.doc", BaselineStatus: "compared", F1: 1, ImageMatch: true}}
	if err := writeJSON(path, first); err != nil {
		t.Fatal(err)
	}
	second := newReport([]string{"second.doc"})
	second.Files = []fileResult{{Path: "second.doc", BaselineStatus: "compared", F1: 1, ImageMatch: true}}
	if err := writeJSON(path, second); err != nil {
		t.Fatalf("replace checkpoint: %v", err)
	}
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(data), "second.doc") || strings.Contains(string(data), "first.doc") {
		t.Fatalf("checkpoint replacement did not persist the latest report: %s", data)
	}
}

func TestMaxFilesStopsAfterPendingWork(t *testing.T) {
	// max-files is intentionally measured after successful resume filtering so
	// a supervisor can make bounded, restartable slices of a 6008-file corpus.
	completed := map[string]fileResult{"already.doc": {Path: "already.doc"}}
	files := []string{"already.doc", "one.doc", "two.doc"}
	pending := 0
	for _, file := range files {
		if _, ok := completed[file]; !ok {
			pending++
		}
	}
	if pending != 2 {
		t.Fatalf("pending count = %d, want 2", pending)
	}
}

func TestResumeCanRetainFailedEntriesForCoveragePass(t *testing.T) {
	prior := []fileResult{{Path: "good.doc"}, {Path: "timeout.doc", Error: "timeout"}}
	for _, keepErrors := range []bool{false, true} {
		completed := map[string]fileResult{}
		for _, result := range prior {
			if result.Error == "" || keepErrors {
				completed[result.Path] = result
			}
		}
		_, hasFailure := completed["timeout.doc"]
		if hasFailure != keepErrors {
			t.Fatalf("keepErrors=%v retained failed result=%v", keepErrors, hasFailure)
		}
	}
}

func TestPathsFileAcceptsOneAbsolutePathPerLine(t *testing.T) {
	dir := t.TempDir()
	paths := filepath.Join(dir, "paths.txt")
	first := filepath.Join(dir, "one.xlsx")
	second := filepath.Join(dir, "two.xlsx")
	if err := os.WriteFile(paths, append([]byte{0xEF, 0xBB, 0xBF}, []byte("\r\n"+first+"\r\n"+second+"\n")...), 0o644); err != nil {
		t.Fatal(err)
	}
	data, err := os.ReadFile(paths)
	if err != nil {
		t.Fatal(err)
	}
	data = bytes.TrimPrefix(data, []byte{0xEF, 0xBB, 0xBF})
	var got []string
	for _, line := range strings.Split(string(data), "\n") {
		line = strings.TrimSpace(strings.TrimSuffix(line, "\r"))
		if line != "" {
			got = append(got, line)
		}
	}
	if len(got) != 2 || got[0] != first || got[1] != second {
		t.Fatalf("paths file = %#v, want %q and %q", got, first, second)
	}
}

func TestContentKeyIncludesExtensionAndFileBytes(t *testing.T) {
	dir := t.TempDir()
	doc := filepath.Join(dir, "one.doc")
	docCopy := filepath.Join(dir, "two.doc")
	docx := filepath.Join(dir, "one.docx")
	if err := os.WriteFile(doc, []byte("same bytes"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(docCopy, []byte("same bytes"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(docx, []byte("same bytes"), 0o644); err != nil {
		t.Fatal(err)
	}
	first, err := contentKey(doc)
	if err != nil {
		t.Fatal(err)
	}
	second, err := contentKey(docCopy)
	if err != nil {
		t.Fatal(err)
	}
	otherExt, err := contentKey(docx)
	if err != nil {
		t.Fatal(err)
	}
	if first != second {
		t.Fatalf("same extension and bytes gave %q and %q", first, second)
	}
	if first == otherExt {
		t.Fatalf("same bytes across extensions shared key %q", first)
	}
}

func TestStrictNoReuseResumeDiscardsBorrowedCheckpoint(t *testing.T) {
	// A strict full-suite audit may be resumed after an earlier throughput run
	// that borrowed outcomes for byte-identical files. Its report must never
	// turn those borrowed records into a false per-path COM execution claim.
	path := filepath.Join("main.go")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	source := string(data)
	for _, required := range []string{
		"if !*reuseIdentical && result.ReusedFrom != \"\"",
		"continue\n\t\t}\n\t\tif result.Error == \"\" || *keepErrors",
	} {
		if !strings.Contains(source, required) {
			t.Errorf("strict no-reuse resume protection missing %q", required)
		}
	}
}

func TestOfficePolicyBlockedClassification(t *testing.T) {
	result := fileResult{BaselineStatus: "baseline-unavailable", Error: "Office baseline: 文件被信任中心的文件阻止设置阻止"}
	if got := diagnose(result); got != "office-policy-blocked" {
		t.Fatalf("diagnose(policy-blocked) = %q", got)
	}
}

func TestDiagnoseClassifiesCOMTimeout(t *testing.T) {
	result := fileResult{BaselineStatus: "baseline-unavailable", Error: "PowerShell COM invocation timed out after 20s"}
	if got := diagnose(result); got != "office-baseline-unavailable" {
		t.Fatalf("diagnose(timeout) = %q", got)
	}
}

func TestDiagnoseClassifiesPasswordProtectedOfficePackage(t *testing.T) {
	result := fileResult{BaselineStatus: "baseline-unavailable", Error: "Office baseline: password-protected Office package; no plaintext Office COM baseline is available without a password (password prompt avoided)"}
	if got := diagnose(result); got != "office-password-protected" {
		t.Fatalf("diagnose(password-protected package) = %q", got)
	}
}

func TestDiagnoseClassifiesMissingOfficeLogonSession(t *testing.T) {
	result := fileResult{BaselineStatus: "baseline-unavailable", Error: "Office baseline: HRESULT:0x80070520 指定的登录会话不存在"}
	if got := diagnose(result); got != "office-session-unavailable" {
		t.Fatalf("diagnose(session unavailable) = %q", got)
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

func TestImageVisualQualityComparisonToleratesExportScaling(t *testing.T) {
	pattern := func(width, height int, checkerboard bool) []byte {
		t.Helper()
		canvas := image.NewRGBA(image.Rect(0, 0, width, height))
		for y := 0; y < height; y++ {
			for x := 0; x < width; x++ {
				light := x*2 < width
				if checkerboard {
					light = (x/2+y/2)%2 == 0
				}
				if light {
					canvas.Set(x, y, color.White)
				} else {
					canvas.Set(x, y, color.Black)
				}
			}
		}
		var buf bytes.Buffer
		if err := png.Encode(&buf, canvas); err != nil {
			t.Fatal(err)
		}
		return buf.Bytes()
	}
	path := filepath.Join(t.TempDir(), "office-export.png")
	if err := os.WriteFile(path, pattern(90, 80, false), 0o644); err != nil {
		t.Fatal(err)
	}
	available, match, note, pairs := imageVisualQualityComparison([]string{path}, []officeread.Image{{Data: pattern(18, 8, false)}})
	if !available || !match || !strings.Contains(note, "Hamming") {
		t.Fatalf("scaled identical visual = (%v, %v, %q), want available match", available, match, note)
	}
	if len(pairs) != 1 || pairs[0].OfficeIndex != 1 || pairs[0].ExtractedIndex != 1 {
		t.Fatalf("visual pairs = %#v, want one reproducible pair", pairs)
	}
	available, match, note, _ = imageVisualQualityComparison([]string{path}, []officeread.Image{{Data: pattern(18, 8, true)}})
	if !available || match || !strings.Contains(note, "Hamming") {
		t.Fatalf("different visual = (%v, %v, %q), want available mismatch", available, match, note)
	}
}

func TestMinimumHammingAssignmentAvoidsGreedyFalseMismatch(t *testing.T) {
	// For Office A, B and extracted X, Y: A->X is locally best, but B->Y is
	// very poor. The globally optimal mapping instead assigns A->Y and B->X.
	// This is the kind of duplicate/reordered picture case a greedy comparator
	// can incorrectly report as a visual mismatch.
	office := []uint64{0b0001, 0b0010}
	extracted := []uint64{0b0000, 0b1101}
	got := minimumHammingAssignment(office, extracted)
	if len(got) != 2 || got[0] != 1 || got[1] != 0 {
		t.Fatalf("assignment = %#v, want []int{1, 0}", got)
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

func TestMissingTokenSegmentsLocateWordParagraph(t *testing.T) {
	segments := []officeTextSegment{
		{Index: 1, Context: "document-paragraph-1", Text: "Visible body text"},
		{Index: 2, Context: "document-paragraph-2", Text: "Date of completion"},
	}
	got := missingTokenSegments(comparisonModeText, segments, "Visible body text completion")
	if len(got) != 1 || got[0].Context != "document-paragraph-2" || !strings.Contains(strings.Join(got[0].MissingTokens, " "), "date") {
		t.Fatalf("Word paragraph diagnosis = %#v, want visible paragraph segment", got)
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

func TestEquationPresentationUsesFormulaComparison(t *testing.T) {
	file := filepath.Join(t.TempDir(), "equation.pptx")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	w, err := zw.Create("ppt/slides/slide1.xml")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := w.Write([]byte(`<p:sld xmlns:p="urn:p" xmlns:m="urn:m"><m:oMath/></p:sld>`)); err != nil {
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

func TestLegacyPPTNormalizationOnlyTargetsPPT(t *testing.T) {
	for path, want := range map[string]bool{
		"deck.ppt":  true,
		"deck.PPT":  true,
		"deck.pptx": false,
		"sheet.xls": false,
	} {
		if got := isLegacyPPT(path); got != want {
			t.Errorf("isLegacyPPT(%q) = %v, want %v", path, got, want)
		}
	}
}

func TestLegacyDOCNormalizationOnlyTargetsDOC(t *testing.T) {
	for path, want := range map[string]bool{
		"document.doc":  true,
		"document.DOC":  true,
		"document.docx": false,
		"deck.ppt":      false,
	} {
		if got := isLegacyDOC(path); got != want {
			t.Errorf("isLegacyDOC(%q) = %v, want %v", path, got, want)
		}
	}
}

func TestNormalizationScriptIsReadOnlyAndTemporarySafe(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_normalize_ppt.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"only legacy .ppt files may be normalized",
		"$true, $false, $false",
		"SaveCopyAs($destination, 24)",
		"normalization output must differ from input",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("normalization script missing safety contract %q", required)
		}
	}
}

func TestDOCNormalizationScriptIsReadOnlyAndTemporarySafe(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_normalize_doc.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"only legacy .doc files may be normalized",
		"$false, $true, $false",
		"SaveAs2($destination, 12)",
		"normalization output must differ from input",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("DOC normalization script missing safety contract %q", required)
		}
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

func TestExcelStrictSparseBridgeReadsOnlyRenderedText(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_baseline.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"function Get-ExcelSparseRenderedTexts",
		"function Get-ExcelRenderedTexts",
		"object value = cell.Text",
		"Microsoft.CSharp.dll",
		"dynamic populated = usedRange.SpecialCells(cellType);",
		"return texts.ToArray()",
		"public static string[] ReadAll",
		"OfficeBaselineExcelTextBridge]::ReadAll($Range)",
		"foreach ($text in (Get-ExcelSparseRenderedTexts $used))",
		"Excel.visible-worksheets.UsedRange.Text",
		"Release(cell)",
		"Release(cells)",
		"ReadWholeRange",
		"if (!foundRenderedCells) ReadWholeRange(usedRange, texts);",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("strict Excel sparse Text bridge missing %q", required)
		}
	}
	for _, forbidden := range []string{
		"OfficeBaselineRenderedCell",
		"cell.Row",
		"cell.Column",
	} {
		if strings.Contains(script, forbidden) {
			t.Errorf("strict Excel sparse bridge must not make extra coordinate COM calls: found %q", forbidden)
		}
	}
}

func TestExcelBaselineBuildsVisibleWorksheetCollection(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_baseline.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	collection := "$visibleSheetIndexes = [System.Collections.Generic.List[int]]::new()"
	add := "$visibleSheetIndexes.Add($sheetIndex)"
	if !strings.Contains(script, collection) || !strings.Contains(script, add) {
		t.Fatal("Excel baseline must retain visible Worksheet indexes for the rendered Text pass")
	}
	measurement := strings.Index(script, "$worksheets = $workbook.Worksheets")
	renderedPass := strings.Index(script, "foreach ($sheetIndex in $visibleSheetIndexes)")
	if measurement < 0 || renderedPass < 0 ||
		strings.Index(script, collection) > measurement ||
		strings.Index(script, add) < measurement ||
		strings.Index(script, add) > renderedPass {
		t.Fatal("visible worksheet index collection must be populated by the UsedRange measurement pass before rendered Text extraction")
	}
}

func TestExcelBaselineDoesNotTreatVariantBoolVisibleSheetsAsHidden(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_baseline.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"$rawSheetVisibility = $sheet.Visible",
		"$rawSheetVisibility -is [bool]",
		"[int]$rawSheetVisibility -eq -1",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("Excel visible-sheet contract missing %q", required)
		}
	}
	if strings.Contains(script, "$sheetVisibility -eq 0 -or $sheetVisibility -eq 2") {
		t.Fatal("Excel baseline must not reinterpret a marshaled VARIANT_BOOL as an XlSheetVisibility enum")
	}
}

func TestExcelBaselineReadsFormulaOnlySheetsThroughText(t *testing.T) {
	body, err := os.ReadFile(filepath.Join("..", "..", "tools", "office_baseline.ps1"))
	if err != nil {
		t.Fatal(err)
	}
	source := string(body)
	for _, want := range []string{
		"foreach (int cellType in new[] { 2, -4123 })",
		"bool foundRenderedCells = false;",
		"foundRenderedCells = true;",
		"if (!foundRenderedCells) ReadWholeRange(usedRange, texts);",
	} {
		if !strings.Contains(source, want) {
			t.Fatalf("Excel strict Text bridge missing %q", want)
		}
	}
	if strings.Contains(source, "bool specialCellsAvailable = false;") {
		t.Fatal("Excel strict Text bridge must not conflate an empty constants range with a formula-only sheet")
	}
}

func TestExcelBaselineAvoidsEnumeratorRCWLifetimeForSheetsAndShapes(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_baseline.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"$worksheets = $workbook.Worksheets",
		"$worksheets.Item($sheetIndex)",
		"Release-ComObject $worksheets",
		"$shapes = $sheet.Shapes",
		"$shapes.Item($shapeIndex)",
		"Release-ComObject $shapes",
		"Release-ComObject $shape",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("Excel COM lifetime contract missing %q", required)
		}
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
	if !strings.Contains(script, "[int]$ExcelMaxCells = 10000") || !strings.Contains(script, "$cellCount -gt $ExcelMaxCells") {
		t.Fatal("Excel COM baseline must expose its visible-cell threshold as ExcelMaxCells")
	}
	if !strings.Contains(script, "foreach ($value in $values)") || strings.Contains(script, "$values.GetValue($rowLower + $row, $columnLower + $column)") {
		t.Fatal("large Excel ranges must enumerate the Value2 SafeArray directly rather than make one managed GetValue call per formatted cell")
	}
}

func TestExcelBaselineNormalOpenDoesNotPassNullPlaceholders(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_baseline.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	if !strings.Contains(script, "$script:excelApp.Workbooks.Open($File, 0, $true)") {
		t.Fatal("Excel baseline must attempt a normal read-only open before recovery mode")
	}
	if strings.Contains(script, "Workbooks.Open($File, 0, $true, 5, '', '', $true, $null") {
		t.Fatal("Excel baseline must not pass null placeholders that break PowerShell COM overload resolution")
	}
}

func TestWordBaselineRetriesWordOpenAndRepairAfterNormalOpen(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_baseline.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"$script:wordApp.Documents.Open($File, $false, $true, $false)",
		"$script:wordApp.Documents.Open($File, $false, $true, $false, '', '', $false, '', '', 0, '', $false, $false, $true, $false)",
		"$script:wordApp.ProtectedViewWindows.Open($File, $false, '', $false)",
		"$source = 'Word.ProtectedView.Content'",
		"if ($null -ne $protectedView)",
		"$protectedView.Close()",
		"$source += '.recovered'",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("Word baseline missing normal/recovery open contract %q", required)
		}
	}
}

func TestRecoveryOptOutRetainsNormalOfficeOpenButSkipsModalFallbacks(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_baseline.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"[switch]$NoRecoveryOpen",
		"if ($NoRecoveryOpen) { throw }",
		"$script:wordApp.Documents.Open($File, $false, $true, $false)",
		"$script:excelApp.Workbooks.Open($File, 0, $true)",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("recovery opt-out baseline contract missing %q", required)
		}
	}
}

func TestWordBaselineAvoidsNonInteractivePasswordDialogs(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "office_baseline.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"function Test-PasswordProtectedOfficePackage",
		"EncryptionInfo",
		"EncryptedPackage",
		"password prompt avoided",
		"Test-PasswordProtectedOfficePackage $onePath",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("password-protected baseline guard missing %q", required)
		}
	}
}

func TestFullSupervisorKeepsUnavailableRetrySeparateFromCoverageCheckpoint(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "run_officebaseline_full.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"RetryUnavailableReportPath",
		"RetryUnavailableTimeoutSeconds",
		"RetryExcelTimeoutSeconds",
		"RetryExcelRetries",
		"[bool]$RetryNoRecoveryOpen = $true",
		"if ($RetryNoRecoveryOpen) { $retryArgs += '-no-recovery-open' }",
		"if ($RetryNoRecoveryOpen) { $excelRetryArgs += '-no-recovery-open' }",
		"--write-errors $retryList --category office-baseline-issue",
		"--extensions .xls,.xlsx",
		"source=Range.Text",
		"focused Excel retry report",
		"function Stop-OfficeBaselineAutomationServersForReport",
		"Stop-OfficeBaselineAutomationServersForReport $report",
		"Stop-OfficeBaselineAutomationServersForReport $retryReport",
		"'-resume', '-keep-errors', '-batch-size', '1', '-checkpoint', '1'",
		"[Math]::Max(90, $TimeoutSeconds * 3)",
		"[Math]::Max(600, $retryTimeout * 2)",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("full supervisor missing unavailable retry contract %q", required)
		}
	}
}

func TestFullSupervisorCanUseNonModalOpenForInitialCoverage(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "run_officebaseline_full.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"[switch]$NoRecoveryOpen",
		"$coverageArgs",
		"if ($NoRecoveryOpen) { $coverageArgs += '-no-recovery-open' }",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("full supervisor non-modal coverage contract missing %q", required)
		}
	}
}

func TestFullSupervisorStopsRepeatedExcelSessionFailuresWithoutCorruptingCheckpoint(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "run_officebaseline_full.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"MaxConsecutiveExcelSessionFailures = 3",
		"$consecutiveExcelSessionFailures",
		"80070520",
		"Excel COM session circuit breaker opened",
		"function Test-ExcelComSession",
		"Excel COM health probe failed before coverage resumed",
		"'-checkpoint', '1'",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("full supervisor missing COM session circuit-breaker contract %q", required)
		}
	}
}

func TestFullSupervisorRequiresUnboundedExcelTextForStrictVisibleAudit(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "run_officebaseline_full.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"[switch]$StrictOfficeVisible",
		"StrictOfficeVisible requires ExcelMaxCells=2147483647",
		"$StrictOfficeVisible -and $ExcelMaxCells -lt 2147483647",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("strict visible audit contract missing %q", required)
		}
	}
}

func TestFullSupervisorDoesNotKillCurrentAutomationServerBetweenSlices(t *testing.T) {
	path := filepath.Join("..", "..", "tools", "run_officebaseline_full.ps1")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	script := string(data)
	for _, required := range []string{
		"function Stop-OfficeBaselineAutomationServers",
		"$oldestActiveStart",
		"-lt $oldestActiveStart",
		"Stop-OfficeBaselineAutomationServers $activeBaselinePids",
	} {
		if !strings.Contains(script, required) {
			t.Errorf("full supervisor missing stale-server ownership guard %q", required)
		}
	}
}
