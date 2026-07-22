// officebaseline compares officeread output with the visible content exposed by
// locally installed Microsoft Office through COM automation. It is intentionally
// a test tool, not a library dependency.
package main

import (
	"bytes"
	"encoding/json"
	"flag"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"sort"
	"strings"
	"time"
	"unicode"

	"officeread"
)

var supportedExts = map[string]bool{
	".doc": true, ".docx": true, ".ppt": true, ".pptx": true, ".xls": true, ".xlsx": true,
}

type officeResult struct {
	Path   string `json:"path"`
	Ext    string `json:"ext"`
	Text   string `json:"text"`
	Images int    `json:"images"`
	Source string `json:"source"`
	Error  string `json:"error"`
}

type fileResult struct {
	Path            string   `json:"path"`
	Ext             string   `json:"ext"`
	OfficeSource    string   `json:"officeSource,omitempty"`
	OfficeTextBytes int      `json:"officeTextBytes"`
	ExtractedBytes  int      `json:"extractedTextBytes"`
	OfficeImages    int      `json:"officeImages"`
	ExtractedImages int      `json:"extractedImages"`
	MatchedTokens   int      `json:"matchedTokens"`
	OfficeTokens    int      `json:"officeTokens"`
	ExtractedTokens int      `json:"extractedTokens"`
	Recall          float64  `json:"recall"`
	Precision       float64  `json:"precision"`
	F1              float64  `json:"f1"`
	ImageDelta      int      `json:"imageDelta"`
	ImageMatch      bool     `json:"imageMatch"`
	MissingTokens   []string `json:"missingTokens,omitempty"`
	ExtraTokens     []string `json:"extraTokens,omitempty"`
	Error           string   `json:"error,omitempty"`
}

type summary struct {
	Total           int     `json:"total"`
	Compared        int     `json:"compared"`
	Errors          int     `json:"errors"`
	OfficeTokens    int     `json:"officeTokens"`
	ExtractedTokens int     `json:"extractedTokens"`
	MatchedTokens   int     `json:"matchedTokens"`
	Recall          float64 `json:"recall"`
	Precision       float64 `json:"precision"`
	F1              float64 `json:"f1"`
	OfficeImages    int     `json:"officeImages"`
	ExtractedImages int     `json:"extractedImages"`
}

type report struct {
	Inputs  []string           `json:"inputs"`
	Summary summary            `json:"summary"`
	ByExt   map[string]summary `json:"byExt"`
	Files   []fileResult       `json:"files"`
}

func main() {
	jsonOut := flag.String("json", "office-baseline-report.json", "JSON report path")
	limit := flag.Int("limit", 0, "maximum files per extension; 0 means unlimited")
	minRecall := flag.Float64("min-recall", 0, "fail if aggregate token recall is below this value")
	minPrecision := flag.Float64("min-precision", 0, "fail if aggregate token precision is below this value")
	timeout := flag.Duration("timeout", 45*time.Second, "maximum COM baseline time per file; 0 disables timeout")
	baselineScript := flag.String("baseline-script", filepath.Join("tools", "office_baseline.ps1"), "PowerShell COM baseline script")
	flag.Parse()
	if script := strings.TrimSpace(os.Getenv("OFFICEBASELINE_SCRIPT")); script != "" && *baselineScript == filepath.Join("tools", "office_baseline.ps1") {
		*baselineScript = script
	}
	if output := strings.TrimSpace(os.Getenv("OFFICEBASELINE_JSON")); output != "" && *jsonOut == "office-baseline-report.json" {
		*jsonOut = output
	}
	if flag.NArg() == 0 {
		if input := strings.TrimSpace(os.Getenv("OFFICEBASELINE_INPUT")); input != "" {
			flag.CommandLine.Parse([]string{input})
		}
	}
	if flag.NArg() == 0 {
		fmt.Fprintln(os.Stderr, "usage: officebaseline [-json report.json] [-limit n] [-timeout 45s] [-min-recall 0.95] [-min-precision 0.95] file-or-directory [...]")
		os.Exit(2)
	}
	if *limit < 0 || *minRecall < 0 || *minRecall > 1 || *minPrecision < 0 || *minPrecision > 1 {
		fmt.Fprintln(os.Stderr, "limit must be >= 0 and thresholds must be between 0 and 1")
		os.Exit(2)
	}
	files, err := collectFiles(flag.Args(), *limit)
	if err != nil {
		fatal(err)
	}
	if _, err := os.Stat(*baselineScript); err != nil {
		fatal(fmt.Errorf("baseline script: %w", err))
	}
	report := report{Inputs: flag.Args(), ByExt: map[string]summary{}}
	for _, file := range files {
		result := compareOne(file, *baselineScript, *timeout)
		report.Files = append(report.Files, result)
		report.Summary = add(report.Summary, result)
		report.ByExt[result.Ext] = add(report.ByExt[result.Ext], result)
		fmt.Printf("%s recall=%.4f precision=%.4f f1=%.4f officeImages=%d extractedImages=%d err=%q\n", result.Path, result.Recall, result.Precision, result.F1, result.OfficeImages, result.ExtractedImages, result.Error)
	}
	report.Summary = finalize(report.Summary)
	for ext, value := range report.ByExt {
		report.ByExt[ext] = finalize(value)
	}
	if err := writeJSON(*jsonOut, report); err != nil {
		fatal(err)
	}
	if report.Summary.Errors > 0 || (*minRecall > 0 && report.Summary.Recall < *minRecall) || (*minPrecision > 0 && report.Summary.Precision < *minPrecision) {
		os.Exit(1)
	}
}

func collectFiles(inputs []string, limit int) ([]string, error) {
	counts, files := map[string]int{}, []string{}
	for _, input := range inputs {
		info, err := os.Stat(input)
		if err != nil {
			return nil, err
		}
		if !info.IsDir() {
			addFile(&files, counts, input, limit)
			continue
		}
		err = filepath.WalkDir(input, func(path string, d os.DirEntry, err error) error {
			if err != nil {
				return err
			}
			if !d.IsDir() {
				addFile(&files, counts, path, limit)
			}
			return nil
		})
		if err != nil {
			return nil, err
		}
	}
	sort.Strings(files)
	return files, nil
}

func addFile(files *[]string, counts map[string]int, path string, limit int) {
	ext := strings.ToLower(filepath.Ext(path))
	if !supportedExts[ext] || (limit > 0 && counts[ext] >= limit) {
		return
	}
	counts[ext]++
	*files = append(*files, path)
}

func compareOne(path, script string, timeout time.Duration) fileResult {
	result := fileResult{Path: path, Ext: strings.ToLower(filepath.Ext(path))}
	absPath, err := filepath.Abs(path)
	if err != nil {
		result.Error = "absolute path: " + err.Error()
		return result
	}
	office, err := runOfficeBaseline(script, absPath, timeout)
	if err != nil {
		result.Error = err.Error()
		return result
	}
	if office.Error != "" {
		result.Error = "Office baseline: " + office.Error
		return result
	}
	extracted, err := officeread.Extract(path, officeread.Options{StrictOfficeImages: true, StrictOfficeContent: true})
	if err != nil {
		result.Error = "officeread: " + err.Error()
		return result
	}
	matched, officeTokens, extractedTokens, missing, extra := tokenOverlap(office.Text, extracted.Text)
	result.OfficeSource, result.OfficeTextBytes, result.ExtractedBytes = office.Source, len(office.Text), len(extracted.Text)
	result.OfficeImages, result.ExtractedImages = office.Images, len(extracted.Images)
	result.MatchedTokens, result.OfficeTokens, result.ExtractedTokens = matched, officeTokens, extractedTokens
	result.MissingTokens, result.ExtraTokens = missing, extra
	result.Recall, result.Precision, result.F1 = scores(matched, officeTokens, extractedTokens)
	result.ImageDelta = result.ExtractedImages - result.OfficeImages
	result.ImageMatch = result.ImageDelta == 0
	return result
}

func runOfficeBaseline(script, path string, timeout time.Duration) (officeResult, error) {
	cmd := exec.Command("powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script, "-Path", path)
	var stderr bytes.Buffer
	cmd.Stderr = &stderr
	var output bytes.Buffer
	cmd.Stdout = &output
	if err := cmd.Start(); err != nil {
		return officeResult{}, fmt.Errorf("start PowerShell COM invocation: %w", err)
	}
	done := make(chan error, 1)
	go func() { done <- cmd.Wait() }()
	var err error
	if timeout > 0 {
		select {
		case err = <-done:
		case <-time.After(timeout):
			_ = cmd.Process.Kill()
			<-done
			return officeResult{}, fmt.Errorf("PowerShell COM invocation timed out after %s", timeout)
		}
	} else {
		err = <-done
	}
	if err != nil {
		return officeResult{}, fmt.Errorf("PowerShell COM invocation: %w: %s", err, strings.TrimSpace(stderr.String()))
	}
	var result officeResult
	if err := json.Unmarshal(output.Bytes(), &result); err != nil {
		return officeResult{}, fmt.Errorf("decode Office baseline: %w", err)
	}
	return result, nil
}

func tokenOverlap(reference, candidate string) (matched, referenceCount, candidateCount int, missing, extra []string) {
	ref, got := tokenCounts(reference), tokenCounts(candidate)
	for token, count := range ref {
		referenceCount += count
		if got[token] < count {
			matched += got[token]
			missing = appendDiagnosticTokens(missing, token, count-got[token])
		} else {
			matched += count
		}
	}
	for token, count := range got {
		candidateCount += count
		if ref[token] < count {
			extra = appendDiagnosticTokens(extra, token, count-ref[token])
		}
	}
	sort.Strings(missing)
	sort.Strings(extra)
	return matched, referenceCount, candidateCount, missing, extra
}

const maxDiagnosticTokens = 40

func appendDiagnosticTokens(out []string, token string, count int) []string {
	for count > 0 && len(out) < maxDiagnosticTokens {
		out = append(out, token)
		count--
	}
	return out
}

func tokenCounts(text string) map[string]int {
	counts := map[string]int{}
	var token strings.Builder
	flush := func() {
		if token.Len() > 0 {
			counts[strings.ToLower(token.String())]++
			token.Reset()
		}
	}
	for _, r := range strings.ToValidUTF8(text, "") {
		if unicode.IsLetter(r) || unicode.IsNumber(r) || unicode.IsMark(r) {
			token.WriteRune(r)
		} else {
			flush()
		}
	}
	flush()
	return counts
}

func scores(matched, reference, candidate int) (recall, precision, f1 float64) {
	if reference == 0 {
		recall = 1
	} else {
		recall = float64(matched) / float64(reference)
	}
	if candidate == 0 {
		precision = 1
	} else {
		precision = float64(matched) / float64(candidate)
	}
	if recall+precision > 0 {
		f1 = 2 * recall * precision / (recall + precision)
	}
	return
}

func add(s summary, result fileResult) summary {
	s.Total++
	if result.Error != "" {
		s.Errors++
		return s
	}
	s.Compared++
	s.OfficeTokens += result.OfficeTokens
	s.ExtractedTokens += result.ExtractedTokens
	s.MatchedTokens += result.MatchedTokens
	s.OfficeImages += result.OfficeImages
	s.ExtractedImages += result.ExtractedImages
	return s
}

func finalize(s summary) summary {
	s.Recall, s.Precision, s.F1 = scores(s.MatchedTokens, s.OfficeTokens, s.ExtractedTokens)
	return s
}

func writeJSON(path string, value report) error {
	if dir := filepath.Dir(path); dir != "." {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return err
		}
	}
	data, err := json.MarshalIndent(value, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(path, data, 0o644)
}
func fatal(err error) { fmt.Fprintln(os.Stderr, err); os.Exit(1) }
