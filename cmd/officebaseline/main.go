// officebaseline compares officeread output with the visible content exposed by
// locally installed Microsoft Office through COM automation. It is intentionally
// a test tool, not a library dependency.
package main

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/base64"
	"encoding/json"
	"flag"
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"time"
	"unicode"

	"golang.org/x/text/unicode/norm"

	"officeread"
)

var supportedExts = map[string]bool{
	".doc": true, ".docx": true, ".ppt": true, ".pptx": true, ".xls": true, ".xlsx": true,
}

type officeResult struct {
	Path        string   `json:"path"`
	Ext         string   `json:"ext"`
	Text        string   `json:"text"`
	Images      int      `json:"images"`
	GroupImages int      `json:"groupImages"`
	ImageFiles  []string `json:"imageFiles"`
	Source      string   `json:"source"`
	Error       string   `json:"error"`
}

type comparisonMode string

const (
	comparisonModeText    comparisonMode = "text"
	comparisonModeFormula comparisonMode = "formula"
)

type comparisonScope string

const (
	comparisonScopeOfficeVisible comparisonScope = "office-visible"
	comparisonScopeOfficeStored  comparisonScope = "office-stored-value"
)

type fileResult struct {
	Path string `json:"path"`
	Ext  string `json:"ext"`
	// BaselineStatus separates an Office/COM execution problem from a content
	// mismatch.  "compared" means that Office returned a baseline; "baseline-
	// unavailable" means the file could not be measured on this machine.
	BaselineStatus             string          `json:"baselineStatus,omitempty"`
	OfficeSource               string          `json:"officeSource,omitempty"`
	OfficeTextBytes            int             `json:"officeTextBytes"`
	ExtractedBytes             int             `json:"extractedTextBytes"`
	OfficeImages               int             `json:"officeImages"`
	OfficeGroupImages          int             `json:"officeGroupImages"`
	ExtractedImages            int             `json:"extractedImages"`
	MatchedTokens              int             `json:"matchedTokens"`
	OfficeTokens               int             `json:"officeTokens"`
	ExtractedTokens            int             `json:"extractedTokens"`
	Recall                     float64         `json:"recall"`
	Precision                  float64         `json:"precision"`
	F1                         float64         `json:"f1"`
	OrderedMatchedTokens       int             `json:"orderedMatchedTokens"`
	OrderedComparisonAvailable bool            `json:"orderedComparisonAvailable"`
	OrderedComparisonNote      string          `json:"orderedComparisonNote,omitempty"`
	OrderedRecall              float64         `json:"orderedRecall"`
	OrderedPrecision           float64         `json:"orderedPrecision"`
	OrderedF1                  float64         `json:"orderedF1"`
	ComparisonMode             comparisonMode  `json:"comparisonMode,omitempty"`
	ComparisonScope            comparisonScope `json:"comparisonScope,omitempty"`
	ImageDelta                 int             `json:"imageDelta"`
	ImageCountMatch            bool            `json:"imageCountMatch"`
	ImageQualityAvailable      bool            `json:"imageQualityAvailable"`
	ImageQualityMatch          bool            `json:"imageQualityMatch"`
	ImageQualityNote           string          `json:"imageQualityNote,omitempty"`
	ImageMatch                 bool            `json:"imageMatch"`
	// Diagnosis makes a report actionable without interpreting raw F1/image
	// counts. It is intentionally a classification, not a second score.
	Diagnosis     string   `json:"diagnosis,omitempty"`
	MissingTokens []string `json:"missingTokens,omitempty"`
	ExtraTokens   []string `json:"extraTokens,omitempty"`
	Error         string   `json:"error,omitempty"`
}

type summary struct {
	Total                  int            `json:"total"`
	Compared               int            `json:"compared"`
	Errors                 int            `json:"errors"`
	BaselineUnavailable    int            `json:"baselineUnavailable"`
	OfficeTokens           int            `json:"officeTokens"`
	ExtractedTokens        int            `json:"extractedTokens"`
	MatchedTokens          int            `json:"matchedTokens"`
	Recall                 float64        `json:"recall"`
	Precision              float64        `json:"precision"`
	F1                     float64        `json:"f1"`
	OrderedMatchedTokens   int            `json:"orderedMatchedTokens"`
	OrderedCompared        int            `json:"orderedCompared"`
	OrderedOfficeTokens    int            `json:"orderedOfficeTokens"`
	OrderedExtractedTokens int            `json:"orderedExtractedTokens"`
	OrderedRecall          float64        `json:"orderedRecall"`
	OrderedPrecision       float64        `json:"orderedPrecision"`
	OrderedF1              float64        `json:"orderedF1"`
	OfficeImages           int            `json:"officeImages"`
	ExtractedImages        int            `json:"extractedImages"`
	DiagnosisCounts        map[string]int `json:"diagnosisCounts,omitempty"`
}

type qualityGate struct {
	Compared                   int      `json:"compared"`
	BaselineUnavailable        int      `json:"baselineUnavailable"`
	ContentCompared            int      `json:"contentCompared"`
	ContentTextMatches         int      `json:"contentTextMatches"`
	ContentImageMatches        int      `json:"contentImageMatches"`
	ContentFullyAligned        int      `json:"contentFullyAligned"`
	ContentTextMatchRate       float64  `json:"contentTextMatchRate"`
	ContentImageMatchRate      float64  `json:"contentImageMatchRate"`
	ContentFullyAlignedRate    float64  `json:"contentFullyAlignedRate"`
	ExcludedScopeMismatchFiles []string `json:"excludedScopeMismatchFiles,omitempty"`
}

type report struct {
	Inputs                  []string           `json:"inputs"`
	Summary                 summary            `json:"summary"`
	QualityGate             qualityGate        `json:"qualityGate"`
	ByExt                   map[string]summary `json:"byExt"`
	Files                   []fileResult       `json:"files"`
	ImageQualityLimitations []string           `json:"imageQualityLimitations,omitempty"`
	BaselineLimitations     []string           `json:"baselineLimitations,omitempty"`
}

func main() {
	jsonOut := flag.String("json", "office-baseline-report.json", "JSON report path")
	limit := flag.Int("limit", 0, "maximum files per extension; 0 means unlimited")
	resume := flag.Bool("resume", false, "reuse successful entries from an existing JSON report and retry failed entries")
	checkpoint := flag.Int("checkpoint", 25, "write a durable JSON checkpoint after this many comparisons; 0 writes only at the end")
	batchSize := flag.Int("batch-size", 1, "files per PowerShell/Office COM session; 1 preserves per-file isolation")
	minRecall := flag.Float64("min-recall", 0, "fail if aggregate token recall is below this value")
	minPrecision := flag.Float64("min-precision", 0, "fail if aggregate token precision is below this value")
	minContentAlignment := flag.Float64("min-content-alignment", 0, "fail if the share of Office-visible files with exact text and image parity is below this value")
	timeout := flag.Duration("timeout", 45*time.Second, "maximum COM baseline time per file; 0 disables timeout")
	baselineRetries := flag.Int("baseline-retries", 1, "extra isolated COM attempts after a baseline transport or Office error")
	baselineScript := flag.String("baseline-script", filepath.Join("tools", "office_baseline.ps1"), "PowerShell COM baseline script")
	excelMaxCells := flag.Int("excel-max-cells", 200000, "maximum Excel UsedRange cells read through rendered Text before Value2 fallback")
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
	if *limit < 0 || *checkpoint < 0 || *batchSize < 1 || *baselineRetries < 0 || *excelMaxCells < 0 || *minRecall < 0 || *minRecall > 1 || *minPrecision < 0 || *minPrecision > 1 || *minContentAlignment < 0 || *minContentAlignment > 1 {
		fmt.Fprintln(os.Stderr, "limit, checkpoint, baseline-retries, and excel-max-cells must be >= 0, batch-size must be >= 1, and thresholds must be between 0 and 1")
		os.Exit(2)
	}
	files, err := collectFiles(flag.Args(), *limit)
	if err != nil {
		fatal(err)
	}
	if _, err := os.Stat(*baselineScript); err != nil {
		fatal(fmt.Errorf("baseline script: %w", err))
	}
	prior, err := loadReport(*jsonOut, *resume)
	if err != nil {
		fatal(err)
	}
	completed := make(map[string]fileResult, len(prior.Files))
	for _, result := range prior.Files {
		// Errors are deliberately not resumed: a transient Office/COM failure
		// must be retried before the complete-corpus audit can pass.
		if result.Error == "" {
			completed[result.Path] = result
		}
	}
	report := newReport(flag.Args())
	for _, file := range files {
		if result, ok := completed[file]; ok {
			report.Files = append(report.Files, result)
		}
	}
	rebuildSummary(&report)
	checked := 0
	for start := 0; start < len(files); {
		pending := make([]string, 0, *batchSize)
		for start < len(files) && len(pending) < *batchSize {
			file := files[start]
			start++
			if _, ok := completed[file]; !ok {
				pending = append(pending, file)
			}
		}
		if len(pending) == 0 {
			continue
		}
		results := compareBatch(pending, *baselineScript, *timeout, *baselineRetries, *excelMaxCells)
		for _, result := range results {
			report.Files = append(report.Files, result)
			fmt.Printf("%s recall=%.4f precision=%.4f f1=%.4f officeImages=%d extractedImages=%d err=%q\n", result.Path, result.Recall, result.Precision, result.F1, result.OfficeImages, result.ExtractedImages, result.Error)
			checked++
			if *checkpoint > 0 && checked%*checkpoint == 0 {
				orderReportFiles(&report, files)
				rebuildSummary(&report)
				if err := writeJSON(*jsonOut, report); err != nil {
					fatal(err)
				}
			}
		}
	}
	orderReportFiles(&report, files)
	rebuildSummary(&report)
	if err := writeJSON(*jsonOut, report); err != nil {
		fatal(err)
	}
	if report.Summary.Errors > 0 || (*minRecall > 0 && report.Summary.Recall < *minRecall) || (*minPrecision > 0 && report.Summary.Precision < *minPrecision) || (*minContentAlignment > 0 && report.QualityGate.ContentFullyAlignedRate < *minContentAlignment) {
		os.Exit(1)
	}
}

func loadReport(path string, resume bool) (report, error) {
	if !resume {
		return report{}, nil
	}
	data, err := os.ReadFile(path)
	if os.IsNotExist(err) {
		return report{}, nil
	}
	if err != nil {
		return report{}, fmt.Errorf("read resume report: %w", err)
	}
	var prior report
	if err := json.Unmarshal(data, &prior); err != nil {
		return report{}, fmt.Errorf("decode resume report: %w", err)
	}
	return prior, nil
}

func newReport(inputs []string) report {
	return report{
		Inputs: inputs,
		ByExt:  map[string]summary{},
		ImageQualityLimitations: []string{
			"Image count parity is mandatory. Source image identity/pixel parity is supplemental because Microsoft Office Shape.Export can rasterize, scale, or re-encode picture shapes.",
		},
		BaselineLimitations: []string{
			"Excel ranges above excel-max-cells are compared through Value2 rather than rendered Text and are marked office-stored-value, so they are excluded from the Office-visible quality gate.",
			"A baseline-unavailable result is a COM transport or Office automation failure, not an extractor content mismatch.",
		},
	}
}

func orderReportFiles(r *report, files []string) {
	byPath := make(map[string]fileResult, len(r.Files))
	for _, result := range r.Files {
		byPath[result.Path] = result
	}
	r.Files = r.Files[:0]
	for _, file := range files {
		if result, ok := byPath[file]; ok {
			r.Files = append(r.Files, result)
		}
	}
}

func rebuildSummary(r *report) {
	r.Summary = summary{}
	r.QualityGate = qualityGate{}
	r.ByExt = map[string]summary{}
	for _, result := range r.Files {
		r.Summary = add(r.Summary, result)
		r.ByExt[result.Ext] = add(r.ByExt[result.Ext], result)
		r.QualityGate = addQualityGate(r.QualityGate, result)
	}
	r.Summary = finalize(r.Summary)
	r.QualityGate = finalizeQualityGate(r.QualityGate)
	for ext, value := range r.ByExt {
		r.ByExt[ext] = finalize(value)
	}
}

// addQualityGate reports exact parity only for files whose COM source is the
// Office-visible contract. Stored-value fallbacks (large Excel Value2 ranges)
// remain in the diagnostic summary but are excluded rather than silently
// depressing a rendered-content quality claim.
func addQualityGate(g qualityGate, result fileResult) qualityGate {
	if result.BaselineStatus == "baseline-unavailable" {
		g.BaselineUnavailable++
		return g
	}
	if result.BaselineStatus != "compared" || result.Error != "" {
		return g
	}
	g.Compared++
	if result.ComparisonScope != comparisonScopeOfficeVisible {
		g.ExcludedScopeMismatchFiles = append(g.ExcludedScopeMismatchFiles, result.Path)
		return g
	}
	g.ContentCompared++
	if result.F1 == 1 {
		g.ContentTextMatches++
	}
	if result.ImageMatch {
		g.ContentImageMatches++
	}
	if result.F1 == 1 && result.ImageMatch {
		g.ContentFullyAligned++
	}
	return g
}

func finalizeQualityGate(g qualityGate) qualityGate {
	if g.ContentCompared == 0 {
		return g
	}
	denominator := float64(g.ContentCompared)
	g.ContentTextMatchRate = float64(g.ContentTextMatches) / denominator
	g.ContentImageMatchRate = float64(g.ContentImageMatches) / denominator
	g.ContentFullyAlignedRate = float64(g.ContentFullyAligned) / denominator
	return g
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

func compareOne(path, script string, timeout time.Duration, excelMaxCells int) fileResult {
	result := fileResult{Path: path, Ext: strings.ToLower(filepath.Ext(path))}
	absPath, err := filepath.Abs(path)
	if err != nil {
		result.Error = "absolute path: " + err.Error()
		return result
	}
	office, err := runOfficeBaseline(script, absPath, timeout, excelMaxCells)
	if err != nil {
		result.BaselineStatus = "baseline-unavailable"
		result.Error = err.Error()
		return result
	}
	if office.Error != "" {
		result.BaselineStatus = "baseline-unavailable"
		result.Error = "Office baseline: " + office.Error
		return result
	}
	extracted, err := officeread.Extract(path, officeread.Options{StrictOfficeImages: true, StrictOfficeContent: true})
	if err != nil {
		result.BaselineStatus = "extractor-error"
		result.Error = "officeread: " + err.Error()
		return result
	}
	result.BaselineStatus = "compared"
	mode := comparisonModeForPath(path)
	matched, officeTokens, extractedTokens, missing, extra := overlapForMode(mode, office.Text, extracted.Text)
	orderedMatched, orderedOfficeTokens, orderedExtractedTokens, orderedAvailable := orderedOverlapForMode(mode, office.Text, extracted.Text)
	result.OfficeSource, result.OfficeTextBytes, result.ExtractedBytes = office.Source, len(office.Text), len(extracted.Text)
	result.OfficeImages, result.OfficeGroupImages, result.ExtractedImages = office.Images, office.GroupImages, len(extracted.Images)
	result.MatchedTokens, result.OfficeTokens, result.ExtractedTokens = matched, officeTokens, extractedTokens
	result.ComparisonMode = mode
	result.ComparisonScope = comparisonScopeForOfficeSource(office.Source)
	result.MissingTokens, result.ExtraTokens = missing, extra
	result.Recall, result.Precision, result.F1 = scores(matched, officeTokens, extractedTokens)
	result.OrderedMatchedTokens = orderedMatched
	result.OrderedComparisonAvailable = orderedAvailable
	if !orderedAvailable {
		result.OrderedComparisonNote = "token product exceeds ordered-comparison limit"
	}
	result.OrderedRecall, result.OrderedPrecision, result.OrderedF1 = scores(orderedMatched, orderedOfficeTokens, orderedExtractedTokens)
	applyImageComparison(&result, office, extracted.Images)
	result.Diagnosis = diagnose(result)
	return result
}

func compareBatch(paths []string, script string, timeout time.Duration, retries int, excelMaxCells int) []fileResult {
	if len(paths) == 1 {
		return []fileResult{compareOneWithRetries(paths[0], script, timeout, retries, excelMaxCells)}
	}
	// Excel COM is not reliable after opening an arbitrary workbook in the same
	// automation process: a modal repair/link prompt can wedge every later file
	// in that session.  Preserve per-file evidence by always isolating Excel
	// files, even if the caller requested a larger generic batch size.
	if excelPaths(paths) {
		results := make([]fileResult, len(paths))
		for i, path := range paths {
			results[i] = compareOneWithRetries(path, script, timeout, retries, excelMaxCells)
		}
		return results
	}
	results := make([]fileResult, len(paths))
	for i, path := range paths {
		results[i] = fileResult{Path: path, Ext: strings.ToLower(filepath.Ext(path))}
	}
	absolute := make([]string, len(paths))
	for i, path := range paths {
		value, err := filepath.Abs(path)
		if err != nil {
			results[i].BaselineStatus = "baseline-unavailable"
			results[i].Error = "absolute path: " + err.Error()
			continue
		}
		absolute[i] = value
	}
	office, err := runOfficeBaselineBatch(script, absolute, timeout, excelMaxCells)
	if err != nil {
		for i := range results {
			if results[i].Error == "" {
				results[i].BaselineStatus = "baseline-unavailable"
				results[i].Error = err.Error()
			}
		}
		if retries > 0 {
			for i := range results {
				if results[i].BaselineStatus == "baseline-unavailable" {
					results[i] = compareOneWithRetries(paths[i], script, timeout, retries-1, excelMaxCells)
				}
			}
		}
		return results
	}
	for i, result := range office {
		if i >= len(results) {
			break
		}
		if result.Error != "" {
			results[i].BaselineStatus = "baseline-unavailable"
			results[i].Error = "Office baseline: " + result.Error
			continue
		}
		extracted, err := officeread.Extract(paths[i], officeread.Options{StrictOfficeImages: true, StrictOfficeContent: true})
		if err != nil {
			results[i].BaselineStatus = "extractor-error"
			results[i].Error = "officeread: " + err.Error()
			continue
		}
		results[i].BaselineStatus = "compared"
		mode := comparisonModeForPath(paths[i])
		matched, officeTokens, extractedTokens, missing, extra := overlapForMode(mode, result.Text, extracted.Text)
		orderedMatched, orderedOfficeTokens, orderedExtractedTokens, orderedAvailable := orderedOverlapForMode(mode, result.Text, extracted.Text)
		results[i].OfficeSource, results[i].OfficeTextBytes, results[i].ExtractedBytes = result.Source, len(result.Text), len(extracted.Text)
		results[i].OfficeImages, results[i].OfficeGroupImages, results[i].ExtractedImages = result.Images, result.GroupImages, len(extracted.Images)
		results[i].MatchedTokens, results[i].OfficeTokens, results[i].ExtractedTokens = matched, officeTokens, extractedTokens
		results[i].ComparisonMode = mode
		results[i].ComparisonScope = comparisonScopeForOfficeSource(result.Source)
		results[i].MissingTokens, results[i].ExtraTokens = missing, extra
		results[i].Recall, results[i].Precision, results[i].F1 = scores(matched, officeTokens, extractedTokens)
		results[i].OrderedMatchedTokens = orderedMatched
		results[i].OrderedComparisonAvailable = orderedAvailable
		if !orderedAvailable {
			results[i].OrderedComparisonNote = "token product exceeds ordered-comparison limit"
		}
		results[i].OrderedRecall, results[i].OrderedPrecision, results[i].OrderedF1 = scores(orderedMatched, orderedOfficeTokens, orderedExtractedTokens)
		applyImageComparison(&results[i], result, extracted.Images)
		results[i].Diagnosis = diagnose(results[i])
	}
	if len(office) != len(results) {
		for i := len(office); i < len(results); i++ {
			results[i].BaselineStatus = "baseline-unavailable"
			results[i].Error = "Office baseline returned incomplete batch"
		}
	}
	// A shared COM session is faster but Excel/Office can leave it wedged after
	// one hostile document. Retry only failed records in a fresh process; this
	// distinguishes a transient automation failure from a reproducible one.
	if retries > 0 {
		for i := range results {
			if results[i].BaselineStatus != "baseline-unavailable" {
				continue
			}
			for attempt := 0; attempt < retries; attempt++ {
				retried := compareOneWithRetries(paths[i], script, timeout, 0, excelMaxCells)
				if retried.BaselineStatus == "compared" || retried.BaselineStatus == "extractor-error" {
					results[i] = retried
					break
				}
				results[i] = retried
			}
		}
	}
	return results
}

// applyImageComparison treats the COM picture count as the mandatory image
// occurrence baseline. Office's Shape.Export is supplemental evidence only:
// it commonly scales/re-rasterizes a picture, so a raw pixel mismatch is not
// evidence that extracted source content is wrong.
func applyImageComparison(result *fileResult, office officeResult, extracted []officeread.Image) {
	result.ImageDelta = len(extracted) - office.Images
	result.ImageCountMatch = result.ImageDelta == 0
	result.ImageMatch = result.ImageCountMatch
	available, match, note := imageQualityComparison(office.ImageFiles, extracted)
	result.ImageQualityAvailable = available
	result.ImageQualityMatch = match
	result.ImageQualityNote = note
}

func imageQualityComparison(officeFiles []string, extracted []officeread.Image) (available, match bool, note string) {
	if len(officeFiles) == 0 {
		return false, false, "Office COM image export unavailable; compared visible picture occurrences by count"
	}
	if len(officeFiles) != len(extracted) {
		return false, false, "Office export/extractor image occurrence count differs"
	}
	officeDigests := make([]string, 0, len(officeFiles))
	for _, path := range officeFiles {
		data, err := os.ReadFile(path)
		if err != nil {
			return false, false, "Office COM image export could not be read"
		}
		digest, ok := decodedImageDigest(data)
		if !ok {
			return false, false, "Office COM export format is not decodable by the baseline tool"
		}
		officeDigests = append(officeDigests, digest)
	}
	extractedDigests := make([]string, 0, len(extracted))
	for _, value := range extracted {
		digest, ok := decodedImageDigest(value.Data)
		if !ok {
			return false, false, "extractor image format is not decodable by the baseline tool"
		}
		extractedDigests = append(extractedDigests, digest)
	}
	sort.Strings(officeDigests)
	sort.Strings(extractedDigests)
	for i := range officeDigests {
		if officeDigests[i] != extractedDigests[i] {
			return false, false, "Office Shape.Export rasterizes/scales pictures; pixel equality is not a valid direct quality baseline"
		}
	}
	return true, true, "decoded image pixels match Microsoft Office export"
}

func decodedImageDigest(data []byte) (string, bool) {
	decoded, _, err := image.Decode(bytes.NewReader(data))
	if err != nil {
		return "", false
	}
	bounds := decoded.Bounds()
	if bounds.Dx() <= 0 || bounds.Dy() <= 0 {
		return "", false
	}
	hash := sha256.New()
	_, _ = fmt.Fprintf(hash, "%dx%d\x00", bounds.Dx(), bounds.Dy())
	for y := bounds.Min.Y; y < bounds.Max.Y; y++ {
		for x := bounds.Min.X; x < bounds.Max.X; x++ {
			r, g, b, a := decoded.At(x, y).RGBA()
			hash.Write([]byte{byte(r >> 8), byte(g >> 8), byte(b >> 8), byte(a >> 8)})
		}
	}
	return fmt.Sprintf("%x", hash.Sum(nil)), true
}

func excelPaths(paths []string) bool {
	if len(paths) == 0 {
		return false
	}
	for _, path := range paths {
		ext := strings.ToLower(filepath.Ext(path))
		if ext != ".xls" && ext != ".xlsx" {
			return false
		}
	}
	return true
}

func compareOneWithRetries(path, script string, timeout time.Duration, retries int, excelMaxCells int) fileResult {
	result := compareOne(path, script, timeout, excelMaxCells)
	for attempt := 0; attempt < retries && result.BaselineStatus == "baseline-unavailable"; attempt++ {
		result = compareOne(path, script, timeout, excelMaxCells)
	}
	return result
}

func runOfficeBaseline(script, path string, timeout time.Duration, excelMaxCells int) (officeResult, error) {
	encoded, marshalErr := json.Marshal([]string{path})
	if marshalErr != nil {
		return officeResult{}, fmt.Errorf("encode Office baseline path: %w", marshalErr)
	}
	cmd := exec.Command("powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script, "-PathsBase64", base64.StdEncoding.EncodeToString(encoded), "-ExcelMaxCells", fmt.Sprint(excelMaxCells))
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
			killProcessTree(cmd)
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
	if err := json.Unmarshal(bytes.TrimSpace(output.Bytes()), &result); err != nil {
		return officeResult{}, fmt.Errorf("decode Office baseline: %w", err)
	}
	return result, nil
}

func runOfficeBaselineBatch(script string, paths []string, timeout time.Duration, excelMaxCells int) ([]officeResult, error) {
	encoded, marshalErr := json.Marshal(paths)
	if marshalErr != nil {
		return nil, fmt.Errorf("encode Office baseline paths: %w", marshalErr)
	}
	args := []string{"-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script, "-PathsBase64", base64.StdEncoding.EncodeToString(encoded), "-ExcelMaxCells", fmt.Sprint(excelMaxCells)}
	cmd := exec.Command("powershell", args...)
	var stderr, output bytes.Buffer
	cmd.Stderr, cmd.Stdout = &stderr, &output
	if err := cmd.Start(); err != nil {
		return nil, fmt.Errorf("start PowerShell COM batch: %w", err)
	}
	done := make(chan error, 1)
	go func() { done <- cmd.Wait() }()
	batchTimeout := timeout*time.Duration(len(paths)) + 15*time.Second
	var err error
	if timeout > 0 {
		select {
		case err = <-done:
		case <-time.After(batchTimeout):
			killProcessTree(cmd)
			<-done
			return nil, fmt.Errorf("PowerShell COM batch invocation timed out after %s", batchTimeout)
		}
	} else {
		err = <-done
	}
	if err != nil {
		return nil, fmt.Errorf("PowerShell COM batch invocation: %w: %s", err, strings.TrimSpace(stderr.String()))
	}
	var results []officeResult
	for _, line := range bytes.Split(output.Bytes(), []byte{'\n'}) {
		line = bytes.TrimSpace(line)
		if len(line) == 0 {
			continue
		}
		var result officeResult
		if err := json.Unmarshal(line, &result); err != nil {
			return nil, fmt.Errorf("decode Office baseline batch: %w", err)
		}
		results = append(results, result)
	}
	return results, nil
}

func killProcessTree(cmd *exec.Cmd) {
	if cmd == nil || cmd.Process == nil {
		return
	}
	if runtime.GOOS == "windows" {
		// COM servers launched by PowerShell can outlive their direct parent.
		// taskkill /T terminates the complete process tree on a timed-out file,
		// preventing a later comparison from attaching to a wedged instance.
		_ = exec.Command("taskkill", "/PID", fmt.Sprint(cmd.Process.Pid), "/T", "/F").Run()
		return
	}
	_ = cmd.Process.Kill()
}

func tokenOverlap(reference, candidate string) (matched, referenceCount, candidateCount int, missing, extra []string) {
	return tokenOverlapStreams(tokenStream(reference), tokenStream(candidate))
}

func tokenOverlapStreams(reference, candidate []string) (matched, referenceCount, candidateCount int, missing, extra []string) {
	ref, got := tokenCountsFromStream(reference), tokenCountsFromStream(candidate)
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

func overlapForMode(mode comparisonMode, reference, candidate string) (matched, referenceCount, candidateCount int, missing, extra []string) {
	if mode == comparisonModeFormula {
		return tokenOverlapStreams(formulaTokenStream(reference), formulaTokenStream(candidate))
	}
	return tokenOverlap(reference, candidate)
}

func comparisonModeForPath(path string) comparisonMode {
	if strings.EqualFold(filepath.Ext(path), ".docx") && docxContainsOfficeMath(path) {
		return comparisonModeFormula
	}
	return comparisonModeText
}

func docxContainsOfficeMath(path string) bool {
	zr, err := zip.OpenReader(path)
	if err != nil {
		return false
	}
	defer zr.Close()
	for _, f := range zr.File {
		if !strings.EqualFold(f.Name, "word/document.xml") {
			continue
		}
		r, err := f.Open()
		if err != nil {
			return false
		}
		data, err := io.ReadAll(r)
		_ = r.Close()
		if err != nil {
			return false
		}
		return bytes.Contains(data, []byte(":oMath"))
	}
	return false
}

// orderedTokenOverlap computes an LCS over normalized token streams. The
// existing bag-of-tokens score is intentionally tolerant of layout changes,
// but it cannot reveal reordered content. Reporting both makes quality claims
// auditable without turning harmless document reflow into a failure.  Very
// large workbooks are reported as unavailable for this expensive metric rather
// than allocating an unbounded LCS work buffer.
const maxOrderedComparisonCells = 4_000_000

func orderedTokenOverlap(reference, candidate string) (matched, referenceCount, candidateCount int, available bool) {
	return orderedTokenOverlapStreams(tokenStream(reference), tokenStream(candidate))
}

func orderedOverlapForMode(mode comparisonMode, reference, candidate string) (matched, referenceCount, candidateCount int, available bool) {
	if mode == comparisonModeFormula {
		return orderedTokenOverlapStreams(formulaTokenStream(reference), formulaTokenStream(candidate))
	}
	return orderedTokenOverlap(reference, candidate)
}

func orderedTokenOverlapStreams(ref, got []string) (matched, referenceCount, candidateCount int, available bool) {
	referenceCount, candidateCount = len(ref), len(got)
	if len(ref) == 0 || len(got) == 0 {
		return 0, referenceCount, candidateCount, true
	}
	if len(ref) > maxOrderedComparisonCells/len(got) {
		return 0, referenceCount, candidateCount, false
	}
	if len(ref) < len(got) {
		// Keep the DP row bounded by the shorter stream.
		ref, got = got, ref
	}
	row := make([]int, len(got)+1)
	for _, left := range ref {
		prev := 0
		for j, right := range got {
			at := row[j+1]
			if left == right {
				row[j+1] = prev + 1
			} else if row[j] > row[j+1] {
				row[j+1] = row[j]
			}
			prev = at
		}
	}
	return row[len(got)], referenceCount, candidateCount, true
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
	return tokenCountsFromStream(tokenStream(text))
}

func tokenCountsFromStream(values []string) map[string]int {
	counts := map[string]int{}
	for _, value := range values {
		counts[value]++
	}
	return counts
}

// formulaTokenStream compares Office Math as a bag of visible symbols rather
// than as prose words. Word's TextRange flattens fractions, scripts and matrix
// cells by visual rows, while OOXML stores their semantic tree order. Keeping
// letters and numbers as individual symbols captures content equivalence
// without claiming a false line-by-line layout match.
func formulaTokenStream(text string) []string {
	var out []string
	for _, r := range norm.NFKD.String(strings.ToValidUTF8(text, "")) {
		if unicode.IsLetter(r) || unicode.IsNumber(r) || unicode.IsMark(r) {
			out = append(out, strings.ToLower(string(r)))
		}
	}
	return out
}

func tokenStream(text string) []string {
	var out []string
	var token strings.Builder
	flush := func() {
		if token.Len() > 0 {
			out = append(out, strings.ToLower(token.String()))
			token.Reset()
		}
	}
	// Word's Content.Text can omit a separating character around a field or an
	// inline object (for example "0043.04Office" and "athttp://..."). The
	// library preserves the logical text boundary as whitespace. Split these
	// only for comparison, so the quality metric evaluates visible content and
	// does not penalize a harmless Word COM serialization quirk.
	text = splitOfficeCompoundTokens(strings.ToValidUTF8(text, ""))
	for _, r := range text {
		if unicode.IsLetter(r) || unicode.IsNumber(r) || unicode.IsMark(r) {
			token.WriteRune(r)
		} else {
			flush()
		}
	}
	flush()
	return out
}

func splitOfficeCompoundTokens(text string) string {
	var out strings.Builder
	runes := []rune(text)
	for i, r := range runes {
		if i > 0 && i+1 < len(runes) && officeCompoundTokenBoundary(runes[i-1], r, runes[i+1]) {
			out.WriteByte(' ')
		}
		out.WriteRune(r)
	}
	return out.String()
}

func officeCompoundTokenBoundary(prev, cur, next rune) bool {
	// Numeric release identifiers such as 0043.04 immediately followed by a
	// title are a common Word field boundary. Retain ordinary alphanumeric
	// words (e.g. "Office365") by splitting only when the next token begins
	// with an uppercase letter after a digit.
	if unicode.IsUpper(cur) && unicode.IsDigit(prev) && !unicode.IsDigit(next) {
		return true
	}
	// Word can concatenate the preceding prose word with a URL field result.
	return cur == 'h' && next == 't' && prev != ' '
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

func diagnose(result fileResult) string {
	if result.BaselineStatus == "baseline-unavailable" {
		return "office-baseline-unavailable"
	}
	if result.BaselineStatus == "extractor-error" {
		return "extractor-error"
	}
	if result.Error != "" {
		return "comparison-error"
	}
	// Large Excel ranges are intentionally read through one COM Value2 SafeArray
	// to avoid a per-cell automation timeout. That is a stored-value baseline,
	// not Excel's rendered Text property: number formatting and formula display
	// can legitimately differ from the extractor's Office-visible contract.
	// Keep the comparison in the report for diagnostics, but never present it as
	// a content-quality failure.
	if result.ComparisonScope == comparisonScopeOfficeStored {
		return "baseline-scope-mismatch"
	}
	// Token equality is the primary text-content gate. The ordered LCS is an
	// audit metric: Word/Excel/PowerPoint can serialize equivalent visible
	// content in a different traversal order (for example footer text or
	// floating-shape anchors), so a fractional ordered score must not turn an
	// otherwise exact extraction into a misleading text-mismatch diagnosis.
	textMatch := result.F1 == 1
	imageMatch := result.ImageMatch
	switch {
	case textMatch && imageMatch:
		return "aligned"
	case !textMatch && !imageMatch:
		return "text-and-image-mismatch"
	case !textMatch:
		return "text-mismatch"
	default:
		return "image-mismatch"
	}
}

func comparisonScopeForOfficeSource(source string) comparisonScope {
	if strings.Contains(strings.ToLower(source), ".value2") {
		return comparisonScopeOfficeStored
	}
	return comparisonScopeOfficeVisible
}

func add(s summary, result fileResult) summary {
	s.Total++
	if s.DiagnosisCounts == nil {
		s.DiagnosisCounts = map[string]int{}
	}
	s.DiagnosisCounts[diagnose(result)]++
	if result.Error != "" {
		s.Errors++
		if result.BaselineStatus == "baseline-unavailable" {
			s.BaselineUnavailable++
		}
		return s
	}
	s.Compared++
	s.OfficeTokens += result.OfficeTokens
	s.ExtractedTokens += result.ExtractedTokens
	s.MatchedTokens += result.MatchedTokens
	if result.OrderedComparisonAvailable {
		s.OrderedCompared++
		s.OrderedMatchedTokens += result.OrderedMatchedTokens
		s.OrderedOfficeTokens += result.OfficeTokens
		s.OrderedExtractedTokens += result.ExtractedTokens
	}
	s.OfficeImages += result.OfficeImages
	s.ExtractedImages += result.ExtractedImages
	return s
}

func finalize(s summary) summary {
	s.Recall, s.Precision, s.F1 = scores(s.MatchedTokens, s.OfficeTokens, s.ExtractedTokens)
	s.OrderedRecall, s.OrderedPrecision, s.OrderedF1 = scores(s.OrderedMatchedTokens, s.OrderedOfficeTokens, s.OrderedExtractedTokens)
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
	// Checkpoints are written while an external COM process is active.  Use a
	// same-directory temporary file and replace it atomically so a cancelled
	// batch never destroys the last valid resume point.
	dir := filepath.Dir(path)
	base := filepath.Base(path)
	tmp, err := os.CreateTemp(dir, base+".tmp-*")
	if err != nil {
		return err
	}
	tmpName := tmp.Name()
	defer os.Remove(tmpName)
	if _, err := tmp.Write(data); err != nil {
		tmp.Close()
		return err
	}
	if err := tmp.Close(); err != nil {
		return err
	}
	return os.Rename(tmpName, path)
}
func fatal(err error) { fmt.Fprintln(os.Stderr, err); os.Exit(1) }
