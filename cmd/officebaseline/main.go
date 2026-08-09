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
	"math/bits"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"time"
	"unicode"

	"golang.org/x/text/unicode/norm"

	"github.com/RapidAI/OfficeRead"
)

var supportedExts = map[string]bool{
	".doc": true, ".docx": true, ".ppt": true, ".pptx": true, ".xls": true, ".xlsx": true,
}

type officeResult struct {
	Path           string              `json:"path"`
	Ext            string              `json:"ext"`
	Text           string              `json:"text"`
	TextSegments   []officeTextSegment `json:"textSegments"`
	Images         int                 `json:"images"`
	GroupImages    int                 `json:"groupImages"`
	InlineImages   int                 `json:"inlineImages"`
	FloatingImages int                 `json:"floatingImages"`
	InlineAnchors  []int               `json:"inlineAnchors,omitempty"`
	ShapeAnchors   []int               `json:"shapeAnchors,omitempty"`
	ImageFiles     []string            `json:"imageFiles"`
	Source         string              `json:"source"`
	Error          string              `json:"error"`
	// FieldTime is the clock Word used when returning dynamic field text. It is
	// passed to the strict extractor so DATE pictures containing seconds do not
	// race the COM invocation.
	FieldTime time.Time `json:"fieldTime,omitempty"`
}

type officeTextSegment struct {
	Index   int    `json:"index"`
	Context string `json:"context"`
	Text    string `json:"text"`
}

type missingTokenSegment struct {
	Index         int      `json:"index"`
	Context       string   `json:"context,omitempty"`
	MissingTokens []string `json:"missingTokens"`
}

// imageVisualPair records the optimal one-to-one pairing used for the
// supplemental visual comparison.  Image order is not a stable contract
// across Office and the file formats, so retaining the pairing makes an image
// quality mismatch reproducible instead of merely reporting an aggregate.
type imageVisualPair struct {
	OfficeIndex    int `json:"officeIndex"`
	ExtractedIndex int `json:"extractedIndex"`
	Hamming        int `json:"hamming"`
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
	// ContentSHA256 makes a reused comparison auditable.  The full-corpus
	// sample set contains byte-identical copies under different paths; opening
	// each copy through Office only repeats the same COM work.  Reuse is limited
	// to an identical extension and byte hash, and ReusedFrom names the original
	// independently compared path.
	ContentSHA256 string `json:"contentSha256,omitempty"`
	ReusedFrom    string `json:"reusedFrom,omitempty"`
	// BaselineStatus separates an Office/COM execution problem from a content
	// mismatch.  "compared" means that Office returned a baseline; "baseline-
	// unavailable" means the file could not be measured on this machine.
	BaselineStatus string `json:"baselineStatus,omitempty"`
	// NormalizedFromOffice is set only when -normalize-legacy-ppt asked the
	// test tool to resave a legacy PPT into a temporary PPTX before extraction.
	// It never changes the source file or the library's normal pure-Go path.
	NormalizedFromOffice       bool            `json:"normalizedFromOffice,omitempty"`
	OfficeSource               string          `json:"officeSource,omitempty"`
	OfficeTextBytes            int             `json:"officeTextBytes"`
	ExtractedBytes             int             `json:"extractedTextBytes"`
	OfficeImages               int             `json:"officeImages"`
	OfficeGroupImages          int             `json:"officeGroupImages"`
	OfficeInlineImages         int             `json:"officeInlineImages"`
	OfficeFloatingImages       int             `json:"officeFloatingImages"`
	OfficeInlineAnchors        []int           `json:"officeInlineAnchors,omitempty"`
	OfficeShapeAnchors         []int           `json:"officeShapeAnchors,omitempty"`
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
	// ImageVisualQuality is a scaling-tolerant, supplemental comparison of
	// decoded Office Shape.Export pictures and extracted source images. Unlike
	// ImageQualityMatch it can retain useful visual evidence when Office has
	// resized or re-rasterized a picture. Count parity remains the hard gate.
	ImageVisualQualityAvailable bool              `json:"imageVisualQualityAvailable"`
	ImageVisualQualityMatch     bool              `json:"imageVisualQualityMatch"`
	ImageVisualQualityNote      string            `json:"imageVisualQualityNote,omitempty"`
	ImageVisualPairs            []imageVisualPair `json:"imageVisualPairs,omitempty"`
	ImageMatch                  bool              `json:"imageMatch"`
	// Diagnosis makes a report actionable without interpreting raw F1/image
	// counts. It is intentionally a classification, not a second score.
	Diagnosis     string   `json:"diagnosis,omitempty"`
	MissingTokens []string `json:"missingTokens,omitempty"`
	ExtraTokens   []string `json:"extraTokens,omitempty"`
	// MissingTokenSegments identifies the Office text shape/range segments that
	// contribute missing visible tokens. It makes a large deck actionable
	// without storing the full (potentially sensitive) baseline text in a report.
	MissingTokenSegments []missingTokenSegment `json:"missingTokenSegments,omitempty"`
	Error                string                `json:"error,omitempty"`
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
	keepErrors := flag.Bool("keep-errors", false, "when resuming, retain prior failed entries instead of retrying them; use for bounded full-corpus coverage before a dedicated retry pass")
	reuseIdentical := flag.Bool("reuse-identical", true, "reuse a same-extension byte-identical result while retaining a path-level audit record")
	checkpoint := flag.Int("checkpoint", 25, "write a durable JSON checkpoint after this many comparisons; 0 writes only at the end")
	batchSize := flag.Int("batch-size", 1, "files per PowerShell/Office COM session; 1 preserves per-file isolation")
	minRecall := flag.Float64("min-recall", 0, "fail if aggregate token recall is below this value")
	minPrecision := flag.Float64("min-precision", 0, "fail if aggregate token precision is below this value")
	minContentAlignment := flag.Float64("min-content-alignment", 0, "fail if the share of Office-visible files with exact text and image parity is below this value")
	timeout := flag.Duration("timeout", 45*time.Second, "maximum COM baseline time per file; 0 disables timeout")
	killGrace := flag.Duration("kill-grace", 5*time.Second, "maximum wait after killing a timed-out PowerShell/Office process tree")
	baselineRetries := flag.Int("baseline-retries", 1, "extra isolated COM attempts after a baseline transport or Office error")
	noRecoveryOpen := flag.Bool("no-recovery-open", false, "do not invoke Office repair/protected-view recovery opens after normal open failure")
	maxFiles := flag.Int("max-files", 0, "maximum eligible files to process after resume; 0 means unlimited")
	baselineScript := flag.String("baseline-script", filepath.Join("tools", "office_baseline.ps1"), "PowerShell COM baseline script")
	normalizeLegacyPPT := flag.Bool("normalize-legacy-ppt", false, "resave legacy .ppt through PowerPoint COM to a temporary .pptx before extracting; test-tool only")
	normalizeScript := flag.String("normalize-script", filepath.Join("tools", "office_normalize_ppt.ps1"), "PowerShell PowerPoint COM normalization script")
	normalizeLegacyDOC := flag.Bool("normalize-legacy-doc", false, "resave legacy .doc through Word COM to a temporary .docx before extracting; test-tool only")
	normalizeDOCScript := flag.String("normalize-doc-script", filepath.Join("tools", "office_normalize_doc.ps1"), "PowerShell Word COM normalization script")
	excelMaxCells := flag.Int("excel-max-cells", 10000, "maximum Excel UsedRange cells read through rendered Text before Value2 fallback")
	pathsFile := flag.String("paths-file", "", "UTF-8 text file containing one input path per line")
	flag.Parse()
	if script := strings.TrimSpace(os.Getenv("OFFICEBASELINE_SCRIPT")); script != "" && *baselineScript == filepath.Join("tools", "office_baseline.ps1") {
		*baselineScript = script
	}
	if output := strings.TrimSpace(os.Getenv("OFFICEBASELINE_JSON")); output != "" && *jsonOut == "office-baseline-report.json" {
		*jsonOut = output
	}
	inputs := append([]string(nil), flag.Args()...)
	if strings.TrimSpace(*pathsFile) != "" {
		data, err := os.ReadFile(*pathsFile)
		if err != nil {
			fatal(fmt.Errorf("read paths file: %w", err))
		}
		// Text editors on Windows commonly write a UTF-8 BOM. Treat it as an
		// encoding marker, not as the first byte of an input path; otherwise
		// GetFileAttributesEx receives a literal U+FEFF prefix and the focused
		// recovery queue fails before its first checkpoint.
		data = bytes.TrimPrefix(data, []byte{0xEF, 0xBB, 0xBF})
		for _, line := range strings.Split(string(data), "\n") {
			line = strings.TrimSpace(strings.TrimSuffix(line, "\r"))
			if line != "" {
				inputs = append(inputs, line)
			}
		}
	}
	if len(inputs) == 0 {
		if input := strings.TrimSpace(os.Getenv("OFFICEBASELINE_INPUT")); input != "" {
			inputs = append(inputs, input)
		}
	}
	if len(inputs) == 0 {
		fmt.Fprintln(os.Stderr, "usage: officebaseline [-json report.json] [-limit n] [-timeout 45s] [-min-recall 0.95] [-min-precision 0.95] file-or-directory [...]")
		os.Exit(2)
	}
	if *limit < 0 || *maxFiles < 0 || *checkpoint < 0 || *batchSize < 1 || *baselineRetries < 0 || *excelMaxCells < 0 || *killGrace < 0 || *minRecall < 0 || *minRecall > 1 || *minPrecision < 0 || *minPrecision > 1 || *minContentAlignment < 0 || *minContentAlignment > 1 {
		fmt.Fprintln(os.Stderr, "limit, checkpoint, baseline-retries, and excel-max-cells must be >= 0, batch-size must be >= 1, and thresholds must be between 0 and 1")
		os.Exit(2)
	}
	// A strict no-reuse audit makes the path-level COM execution requirement
	// durable across resume.  A report created with reuse enabled may already
	// contain a successful borrowed result.  Do not silently regard it as a
	// completed baseline after the caller switches to -reuse-identical=false.
	// Re-run it and clear the audit annotation, while retaining genuine direct
	// successful checkpoints.
	files, err := collectFiles(inputs, *limit)
	if err != nil {
		fatal(err)
	}
	if _, err := os.Stat(*baselineScript); err != nil {
		fatal(fmt.Errorf("baseline script: %w", err))
	}
	if *normalizeLegacyPPT {
		if _, err := os.Stat(*normalizeScript); err != nil {
			fatal(fmt.Errorf("normalization script: %w", err))
		}
		// A normalization creates a separate PowerPoint COM process for each
		// legacy file.  Preserve isolation and avoid mixing source baselines with
		// a converted extraction in a shared batch.
		*batchSize = 1
	}
	if *normalizeLegacyDOC {
		if _, err := os.Stat(*normalizeDOCScript); err != nil {
			fatal(fmt.Errorf("DOC normalization script: %w", err))
		}
		*batchSize = 1
	}
	prior, err := loadReport(*jsonOut, *resume)
	if err != nil {
		fatal(err)
	}
	completed := make(map[string]fileResult, len(prior.Files))
	for _, result := range prior.Files {
		// By default errors are retried because COM failures can be transient.
		// A large-corpus supervisor first needs a complete, stable coverage map,
		// however, so -keep-errors explicitly retains them for a later focused
		// retry pass rather than repeatedly consuming every slice.
		if !*reuseIdentical && result.ReusedFrom != "" {
			continue
		}
		if result.Error == "" || *keepErrors {
			completed[result.Path] = result
		}
	}
	contentKeys := make(map[string]string, len(files))
	duplicates := make(map[string][]string)
	if *reuseIdentical {
		// Select one representative for each byte-identical same-extension
		// group. Once its outcome is known, every matching path receives an
		// explicit result referencing that representative. This retains all 6008
		// path records but avoids repeatedly triggering the same Office parser or
		// modal repair path for a duplicated file.
		known := make(map[string]string)
		for path, result := range completed {
			if result.Error != "" || result.BaselineStatus != "compared" {
				continue
			}
			key, hashErr := contentKey(path)
			if hashErr != nil {
				continue
			}
			contentKeys[path] = key
			result.ContentSHA256 = contentDigest(key)
			completed[path] = result
			known[key] = path
		}
		for _, file := range files {
			if _, alreadyCompleted := completed[file]; alreadyCompleted {
				continue
			}
			key, hashErr := contentKey(file)
			if hashErr != nil {
				continue
			}
			contentKeys[file] = key
			if source, ok := known[key]; ok {
				if sourceResult, alreadyCompleted := completed[source]; alreadyCompleted {
					reused := sourceResult
					reused.Path = file
					reused.Ext = strings.ToLower(filepath.Ext(file))
					reused.ContentSHA256 = contentDigest(key)
					reused.ReusedFrom = source
					completed[file] = reused
					continue
				}
				duplicates[source] = append(duplicates[source], file)
				continue
			}
			known[key] = file
		}
	}
	report := newReport(inputs)
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
		if *maxFiles > 0 && checked >= *maxFiles {
			break
		}
		if *maxFiles > 0 && len(pending) > *maxFiles-checked {
			pending = pending[:*maxFiles-checked]
		}
		results := compareBatch(pending, *baselineScript, *normalizeScript, *normalizeLegacyPPT, *normalizeDOCScript, *normalizeLegacyDOC, *timeout, *killGrace, *baselineRetries, *excelMaxCells, *noRecoveryOpen)
		for _, result := range results {
			if key := contentKeys[result.Path]; key != "" {
				result.ContentSHA256 = contentDigest(key)
			}
			// Errors return early from the per-file comparison path, before its
			// successful-content diagnosis is calculated. Persist the same
			// classification used by summary aggregation so a retry/audit tool can
			// reliably select COM timeouts and local-policy blocks per path.
			result.Diagnosis = diagnose(result)
			report.Files = append(report.Files, result)
			completed[result.Path] = result
			fmt.Printf("%s recall=%.4f precision=%.4f f1=%.4f officeImages=%d extractedImages=%d err=%q\n", result.Path, result.Recall, result.Precision, result.F1, result.OfficeImages, result.ExtractedImages, result.Error)
			for _, duplicate := range duplicates[result.Path] {
				reused := result
				reused.Path = duplicate
				reused.Ext = strings.ToLower(filepath.Ext(duplicate))
				reused.ContentSHA256 = contentDigest(contentKeys[duplicate])
				reused.ReusedFrom = result.Path
				report.Files = append(report.Files, reused)
				completed[duplicate] = reused
			}
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
	if checked > 0 && *checkpoint > 0 && checked%*checkpoint != 0 {
		orderReportFiles(&report, files)
		rebuildSummary(&report)
		if err := writeJSON(*jsonOut, report); err != nil {
			fatal(err)
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
			"Image count parity is mandatory. Exact source-pixel parity and scaling-tolerant visual fingerprints are supplemental because Microsoft Office Shape.Export can rasterize, scale, or re-encode picture shapes.",
		},
		BaselineLimitations: []string{
			"Excel ranges above excel-max-cells are compared through Value2 rather than rendered Text and are marked office-stored-value, so they are excluded from the Office-visible quality gate.",
			"A baseline-unavailable result is a COM transport or Office automation failure, not an extractor content mismatch.",
			"When -normalize-legacy-ppt or -normalize-legacy-doc is set, only the selected legacy file type is extracted from a temporary Office-authored OOXML copy. The source is opened read-only and never modified; this is an optional diagnostic/deployment mode, not the library default.",
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

func contentKey(path string) (string, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return "", err
	}
	ext := strings.ToLower(filepath.Ext(path))
	sum := sha256.Sum256(data)
	return ext + ":" + fmt.Sprintf("%x", sum[:]), nil
}

func contentDigest(key string) string {
	if colon := strings.IndexByte(key, ':'); colon >= 0 {
		return key[colon+1:]
	}
	return key
}

func addFile(files *[]string, counts map[string]int, path string, limit int) {
	ext := strings.ToLower(filepath.Ext(path))
	// Office owner/lock files ("~$<name>.docx", etc.) are transient metadata,
	// not documents.  They can appear while a COM baseline is in flight and
	// must never displace a real corpus sample under a per-extension limit.
	if strings.HasPrefix(filepath.Base(path), "~$") || !supportedExts[ext] || (limit > 0 && counts[ext] >= limit) {
		return
	}
	counts[ext]++
	*files = append(*files, path)
}

func extractForOfficeComparison(path string, fieldTime time.Time) (*officeread.Result, error) {
	return officeread.Extract(path, officeread.Options{
		StrictOfficeImages:  true,
		StrictOfficeContent: true,
		OfficeFieldTime:     fieldTime,
	})
}

func compareOne(path, script, normalizePPTScript string, normalizeLegacyPPT bool, normalizeDOCScript string, normalizeLegacyDOC bool, timeout, killGrace time.Duration, excelMaxCells int, noRecoveryOpen bool) fileResult {
	result := fileResult{Path: path, Ext: strings.ToLower(filepath.Ext(path))}
	absPath, err := filepath.Abs(path)
	if err != nil {
		result.Error = "absolute path: " + err.Error()
		return result
	}
	office, err := runOfficeBaseline(script, absPath, timeout, killGrace, excelMaxCells, noRecoveryOpen)
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
	extractionPath := path
	if normalizeLegacyPPT && isLegacyPPT(path) {
		// Office itself can open a small class of protected/RC4 legacy decks
		// whose binary records are unavailable to the pure Go parser. The
		// explicitly requested test-only normalization creates a temporary PPTX
		// so this comparison measures the extraction contract rather than a
		// known unsupported ciphertext representation. The source remains
		// read-only and is never saved in place.
		normalized, cleanup, normalizeErr := normalizeLegacyOfficeFile(normalizePPTScript, absPath, ".pptx", timeout, killGrace)
		if normalizeErr != nil {
			result.BaselineStatus = "baseline-unavailable"
			result.Error = "PowerPoint normalization: " + normalizeErr.Error()
			return result
		}
		defer cleanup()
		extractionPath = normalized
		result.NormalizedFromOffice = true
	}
	if normalizeLegacyDOC && isLegacyDOC(path) {
		normalized, cleanup, normalizeErr := normalizeLegacyOfficeFile(normalizeDOCScript, absPath, ".docx", timeout, killGrace)
		if normalizeErr != nil {
			result.BaselineStatus = "baseline-unavailable"
			result.Error = "Word normalization: " + normalizeErr.Error()
			return result
		}
		defer cleanup()
		extractionPath = normalized
		result.NormalizedFromOffice = true
	}
	extracted, err := extractForOfficeComparison(extractionPath, office.FieldTime)
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
	if len(missing) > 0 {
		result.MissingTokenSegments = missingTokenSegments(mode, office.TextSegments, extracted.Text)
	}
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

func compareBatch(paths []string, script, normalizePPTScript string, normalizeLegacyPPT bool, normalizeDOCScript string, normalizeLegacyDOC bool, timeout, killGrace time.Duration, retries int, excelMaxCells int, noRecoveryOpen bool) []fileResult {
	if len(paths) == 1 {
		return []fileResult{compareOneWithRetries(paths[0], script, normalizePPTScript, normalizeLegacyPPT, normalizeDOCScript, normalizeLegacyDOC, timeout, killGrace, retries, excelMaxCells, noRecoveryOpen)}
	}
	// Excel COM is not reliable after opening an arbitrary workbook in the same
	// automation process: a modal repair/link prompt can wedge every later file
	// in that session.  Preserve per-file evidence by always isolating Excel
	// files, even if the caller requested a larger generic batch size.
	if excelPaths(paths) {
		results := make([]fileResult, len(paths))
		for i, path := range paths {
			results[i] = compareOneWithRetries(path, script, normalizePPTScript, normalizeLegacyPPT, normalizeDOCScript, normalizeLegacyDOC, timeout, killGrace, retries, excelMaxCells, noRecoveryOpen)
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
	office, err := runOfficeBaselineBatch(script, absolute, timeout, killGrace, excelMaxCells, noRecoveryOpen)
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
					results[i] = compareOneWithRetries(paths[i], script, normalizePPTScript, normalizeLegacyPPT, normalizeDOCScript, normalizeLegacyDOC, timeout, killGrace, retries-1, excelMaxCells, noRecoveryOpen)
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
		extractionPath := paths[i]
		var cleanup func()
		if normalizeLegacyPPT && isLegacyPPT(paths[i]) {
			normalized, closeNormalized, normalizeErr := normalizeLegacyOfficeFile(normalizePPTScript, absolute[i], ".pptx", timeout, killGrace)
			if normalizeErr != nil {
				results[i].BaselineStatus = "baseline-unavailable"
				results[i].Error = "PowerPoint normalization: " + normalizeErr.Error()
				continue
			}
			extractionPath, cleanup = normalized, closeNormalized
			results[i].NormalizedFromOffice = true
		}
		if normalizeLegacyDOC && isLegacyDOC(paths[i]) {
			normalized, closeNormalized, normalizeErr := normalizeLegacyOfficeFile(normalizeDOCScript, absolute[i], ".docx", timeout, killGrace)
			if normalizeErr != nil {
				results[i].BaselineStatus = "baseline-unavailable"
				results[i].Error = "Word normalization: " + normalizeErr.Error()
				continue
			}
			extractionPath, cleanup = normalized, closeNormalized
			results[i].NormalizedFromOffice = true
		}
		extracted, err := extractForOfficeComparison(extractionPath, result.FieldTime)
		if cleanup != nil {
			cleanup()
		}
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
		results[i].OfficeInlineImages, results[i].OfficeFloatingImages = result.InlineImages, result.FloatingImages
		results[i].OfficeInlineAnchors = append([]int(nil), result.InlineAnchors...)
		results[i].OfficeShapeAnchors = append([]int(nil), result.ShapeAnchors...)
		results[i].MatchedTokens, results[i].OfficeTokens, results[i].ExtractedTokens = matched, officeTokens, extractedTokens
		results[i].ComparisonMode = mode
		results[i].ComparisonScope = comparisonScopeForOfficeSource(result.Source)
		results[i].MissingTokens, results[i].ExtraTokens = missing, extra
		if len(missing) > 0 {
			results[i].MissingTokenSegments = missingTokenSegments(mode, result.TextSegments, extracted.Text)
		}
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
				retried := compareOneWithRetries(paths[i], script, normalizePPTScript, normalizeLegacyPPT, normalizeDOCScript, normalizeLegacyDOC, timeout, killGrace, 0, excelMaxCells, noRecoveryOpen)
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
	visualAvailable, visualMatch, visualNote, visualPairs := imageVisualQualityComparison(office.ImageFiles, extracted)
	result.ImageVisualQualityAvailable = visualAvailable
	result.ImageVisualQualityMatch = visualMatch
	result.ImageVisualQualityNote = visualNote
	result.ImageVisualPairs = visualPairs
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

// imageVisualQualityComparison compares the occurrence multiset through a
// small perceptual fingerprint. Shape.Export may resize/re-encode a source
// picture, so exact pixels are often unavailable even when the visual content
// agrees. This remains supplemental evidence: only occurrence count affects
// ImageMatch and the hard quality gate.
func imageVisualQualityComparison(officeFiles []string, extracted []officeread.Image) (available, match bool, note string, pairs []imageVisualPair) {
	if len(officeFiles) == 0 {
		return false, false, "Office COM image export unavailable; visual fingerprint comparison skipped", nil
	}
	if len(officeFiles) != len(extracted) {
		return false, false, "Office export/extractor image occurrence count differs", nil
	}
	officePrints := make([]uint64, 0, len(officeFiles))
	for _, path := range officeFiles {
		data, err := os.ReadFile(path)
		if err != nil {
			return false, false, "Office COM image export could not be read", nil
		}
		fingerprint, ok := decodedImageVisualFingerprint(data)
		if !ok {
			return false, false, "Office COM export format is not decodable by the baseline tool", nil
		}
		officePrints = append(officePrints, fingerprint)
	}
	extractedPrints := make([]uint64, 0, len(extracted))
	for _, value := range extracted {
		fingerprint, ok := decodedImageVisualFingerprint(value.Data)
		if !ok {
			return false, false, "extractor image format is not decodable by the baseline tool", nil
		}
		extractedPrints = append(extractedPrints, fingerprint)
	}
	// Pair the multisets with a global minimum-cost assignment.  Greedy nearest
	// neighbour matching can consume an image that is the only close match for
	// a later Office export, producing a false image-quality mismatch.
	assignment := minimumHammingAssignment(officePrints, extractedPrints)
	totalDistance := 0
	maxDistance := 0
	for officeIndex, extractedIndex := range assignment {
		if extractedIndex < 0 || extractedIndex >= len(extractedPrints) {
			return false, false, "visual fingerprint matching could not pair all picture occurrences", nil
		}
		distance := bits.OnesCount64(officePrints[officeIndex] ^ extractedPrints[extractedIndex])
		pairs = append(pairs, imageVisualPair{OfficeIndex: officeIndex + 1, ExtractedIndex: extractedIndex + 1, Hamming: distance})
		totalDistance += distance
		if distance > maxDistance {
			maxDistance = distance
		}
	}
	meanDistance := float64(totalDistance) / float64(len(officePrints))
	// A 64-bit dHash tolerates modest scaling/codec changes. Empirically, a
	// distance <= 12 per image and <= 8 on average distinguishes the same
	// rendering from unrelated content without becoming a correctness gate.
	match = maxDistance <= 12 && meanDistance <= 8
	return true, match, fmt.Sprintf("scaling-tolerant visual fingerprint (optimal one-to-one pairing): mean Hamming distance %.1f/64, maximum %d/64", meanDistance, maxDistance), pairs
}

// minimumHammingAssignment returns a minimum-total-Hamming-distance matching
// from each Office picture to exactly one extracted picture.  It is the
// square Hungarian algorithm (O(n^3)); the largest observed deck has only a
// few hundred image occurrences, so it is cheap compared with COM export.
func minimumHammingAssignment(office, extracted []uint64) []int {
	n := len(office)
	if n == 0 || n != len(extracted) {
		return nil
	}
	// The implementation is 1-indexed following the standard potential-based
	// formulation.  Add the extracted index as a tiny tie-breaker so reports
	// are deterministic when duplicate pictures have the same fingerprint.
	u := make([]int, n+1)
	v := make([]int, n+1)
	p := make([]int, n+1)
	way := make([]int, n+1)
	for i := 1; i <= n; i++ {
		p[0] = i
		j0 := 0
		minv := make([]int, n+1)
		used := make([]bool, n+1)
		for j := 1; j <= n; j++ {
			minv[j] = 1 << 30
		}
		for {
			used[j0] = true
			i0 := p[j0]
			delta, j1 := 1<<30, 0
			for j := 1; j <= n; j++ {
				if used[j] {
					continue
				}
				cost := bits.OnesCount64(office[i0-1]^extracted[j-1])*1024 + j
				cur := cost - u[i0] - v[j]
				if cur < minv[j] {
					minv[j], way[j] = cur, j0
				}
				if minv[j] < delta {
					delta, j1 = minv[j], j
				}
			}
			for j := 0; j <= n; j++ {
				if used[j] {
					u[p[j]] += delta
					v[j] -= delta
				} else {
					minv[j] -= delta
				}
			}
			j0 = j1
			if p[j0] == 0 {
				break
			}
		}
		for {
			j1 := way[j0]
			p[j0] = p[j1]
			j0 = j1
			if j0 == 0 {
				break
			}
		}
	}
	assignment := make([]int, n)
	for i := range assignment {
		assignment[i] = -1
	}
	for j := 1; j <= n; j++ {
		assignment[p[j]-1] = j - 1
	}
	return assignment
}

func decodedImageVisualFingerprint(data []byte) (uint64, bool) {
	decoded, _, err := image.Decode(bytes.NewReader(data))
	if err != nil {
		return 0, false
	}
	bounds := decoded.Bounds()
	if bounds.Dx() <= 0 || bounds.Dy() <= 0 {
		return 0, false
	}
	// Difference hash: sample a 9x8 luminance grid, then compare horizontal
	// neighbours. It is insensitive to source dimensions and PNG/JPEG wrapper
	// differences while retaining broad visual structure.
	var hash uint64
	for y := 0; y < 8; y++ {
		for x := 0; x < 8; x++ {
			left := sampledImageLuma(decoded, bounds, x, y, 9, 8)
			right := sampledImageLuma(decoded, bounds, x+1, y, 9, 8)
			if left >= right {
				hash |= uint64(1) << uint(y*8+x)
			}
		}
	}
	return hash, true
}

func sampledImageLuma(value image.Image, bounds image.Rectangle, x, y, width, height int) uint32 {
	px := bounds.Min.X + (x*(bounds.Dx()-1))/(width-1)
	py := bounds.Min.Y + (y*(bounds.Dy()-1))/(height-1)
	r, g, b, _ := value.At(px, py).RGBA()
	return (299*r + 587*g + 114*b) / 1000
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

func compareOneWithRetries(path, script, normalizePPTScript string, normalizeLegacyPPT bool, normalizeDOCScript string, normalizeLegacyDOC bool, timeout, killGrace time.Duration, retries int, excelMaxCells int, noRecoveryOpen bool) fileResult {
	result := compareOne(path, script, normalizePPTScript, normalizeLegacyPPT, normalizeDOCScript, normalizeLegacyDOC, timeout, killGrace, excelMaxCells, noRecoveryOpen)
	for attempt := 0; attempt < retries && result.BaselineStatus == "baseline-unavailable"; attempt++ {
		// A timed-out Office automation process may have left an /Automation
		// server alive outside the PowerShell child tree.  Starting the retry
		// while that server is still shutting down merely consumes another full
		// timeout.  Drain only automation-tagged servers; interactive Office
		// windows are deliberately never selected by this cleanup.
		terminateOfficeAutomationServers(killGrace)
		if killGrace > 0 {
			time.Sleep(killGrace)
		}
		result = compareOne(path, script, normalizePPTScript, normalizeLegacyPPT, normalizeDOCScript, normalizeLegacyDOC, timeout, killGrace, excelMaxCells, noRecoveryOpen)
	}
	return result
}

func isLegacyPPT(path string) bool {
	return strings.EqualFold(filepath.Ext(path), ".ppt")
}

func isLegacyDOC(path string) bool {
	return strings.EqualFold(filepath.Ext(path), ".doc")
}

// normalizeLegacyPPTX invokes the companion PowerShell script to convert a
// legacy PowerPoint deck into an Office-authored temporary PPTX.  The original
// input is opened read-only and remains untouched.  This intentionally lives
// in the COM baseline command, not the officeread library.
func normalizeLegacyOfficeFile(script, path, outputExt string, timeout, killGrace time.Duration) (string, func(), error) {
	if runtime.GOOS != "windows" {
		return "", nil, fmt.Errorf("PowerPoint COM normalization requires Windows")
	}
	work, err := os.MkdirTemp("", "officebaseline-normalized-")
	if err != nil {
		return "", nil, fmt.Errorf("create temporary normalization directory: %w", err)
	}
	cleanup := func() { _ = os.RemoveAll(work) }
	output := filepath.Join(work, strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))+outputExt)
	args := []string{"-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script, "-InputPath", path, "-OutputPath", output}
	command := exec.Command("powershell.exe", args...)
	var outputBytes bytes.Buffer
	command.Stdout, command.Stderr = &outputBytes, &outputBytes
	if err := command.Start(); err != nil {
		cleanup()
		return "", nil, fmt.Errorf("start Office normalization: %w", err)
	}
	runErr := waitCommand(command, timeout, killGrace, "Office normalization")
	if runErr != nil {
		cleanup()
		return "", nil, fmt.Errorf("%w: %s", runErr, strings.TrimSpace(outputBytes.String()))
	}
	info, err := os.Stat(output)
	if err != nil || info.Size() == 0 {
		cleanup()
		return "", nil, fmt.Errorf("PowerPoint produced no PPTX output")
	}
	return output, cleanup, nil
}

func runOfficeBaseline(script, path string, timeout, killGrace time.Duration, excelMaxCells int, noRecoveryOpen bool) (officeResult, error) {
	encoded, marshalErr := json.Marshal([]string{path})
	if marshalErr != nil {
		return officeResult{}, fmt.Errorf("encode Office baseline path: %w", marshalErr)
	}
	runID := newOfficeBaselineRunID()
	args := []string{"-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script, "-PathsBase64", base64.StdEncoding.EncodeToString(encoded), "-ExcelMaxCells", fmt.Sprint(excelMaxCells), "-RunId", runID}
	if noRecoveryOpen {
		args = append(args, "-NoRecoveryOpen")
	}
	cmd := exec.Command("powershell", args...)
	var stderr bytes.Buffer
	cmd.Stderr = &stderr
	var output bytes.Buffer
	cmd.Stdout = &output
	if err := cmd.Start(); err != nil {
		return officeResult{}, fmt.Errorf("start PowerShell COM invocation: %w", err)
	}
	err := waitCommandWithCleanup(cmd, timeout, killGrace, "PowerShell COM invocation", func() { terminateOfficeBaselineRun(runID, killGrace) })
	if err != nil {
		return officeResult{}, fmt.Errorf("PowerShell COM invocation: %w: %s", err, strings.TrimSpace(stderr.String()))
	}
	var result officeResult
	if err := json.Unmarshal(bytes.TrimSpace(output.Bytes()), &result); err != nil {
		return officeResult{}, fmt.Errorf("decode Office baseline: %w", err)
	}
	return result, nil
}

func runOfficeBaselineBatch(script string, paths []string, timeout, killGrace time.Duration, excelMaxCells int, noRecoveryOpen bool) ([]officeResult, error) {
	encoded, marshalErr := json.Marshal(paths)
	if marshalErr != nil {
		return nil, fmt.Errorf("encode Office baseline paths: %w", marshalErr)
	}
	runID := newOfficeBaselineRunID()
	args := []string{"-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script, "-PathsBase64", base64.StdEncoding.EncodeToString(encoded), "-ExcelMaxCells", fmt.Sprint(excelMaxCells), "-RunId", runID}
	if noRecoveryOpen {
		args = append(args, "-NoRecoveryOpen")
	}
	cmd := exec.Command("powershell", args...)
	var stderr, output bytes.Buffer
	cmd.Stderr, cmd.Stdout = &stderr, &output
	if err := cmd.Start(); err != nil {
		return nil, fmt.Errorf("start PowerShell COM batch: %w", err)
	}
	batchTimeout := timeout*time.Duration(len(paths)) + 15*time.Second
	err := waitCommandWithCleanup(cmd, batchTimeout, killGrace, "PowerShell COM batch invocation", func() { terminateOfficeBaselineRun(runID, killGrace) })
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
		pid := fmt.Sprint(cmd.Process.Pid)
		// Launch taskkill before terminating the root.  Killing the root first can
		// re-parent a blocked PowerShell child, making /T unable to find it.  This
		// used to be a goroutine around Run(), which still allowed the scheduler to
		// kill the root before taskkill had even started.  Start establishes the
		// helper process synchronously without waiting for its potentially slow
		// traversal, so the timeout path stays bounded and the tree relationship is
		// retained for taskkill.
		if taskkill := exec.Command("taskkill", "/PID", pid, "/T", "/F"); taskkill.Start() == nil {
			// Give taskkill a short head start to enumerate and signal the tree
			// while its parent relationship is still intact. Merely starting the
			// helper then immediately killing the root can re-parent the COM host
			// before taskkill traverses it, which is exactly how a timed-out Excel
			// child survives past the watchdog. This remains strictly bounded: a
			// blocked taskkill never holds the corpus runner for more than 750ms.
			done := make(chan struct{})
			go func() { _ = taskkill.Wait(); close(done) }()
			select {
			case <-done:
			case <-time.After(750 * time.Millisecond):
			}
		}
		_ = cmd.Process.Kill()
		return
	}
	_ = cmd.Process.Kill()
}

// waitCommand places a strict upper bound on both the normal execution time
// and the post-kill wait.  Previous code waited indefinitely after taskkill;
// a wedged Word COM call could therefore stall the complete 6008-file audit.
// Returning after killGrace leaves a best-effort cleanup goroutine behind but
// guarantees that the next isolated file can be checkpointed and attempted.
func waitCommand(cmd *exec.Cmd, timeout, killGrace time.Duration, label string) error {
	return waitCommandWithCleanup(cmd, timeout, killGrace, label, nil)
}

func waitCommandWithCleanup(cmd *exec.Cmd, timeout, killGrace time.Duration, label string, cleanup func()) error {
	done := make(chan error, 1)
	go func() { done <- cmd.Wait() }()
	if timeout <= 0 {
		return <-done
	}
	select {
	case err := <-done:
		return err
	case <-time.After(timeout):
		killProcessTree(cmd)
		if cleanup != nil {
			cleanup()
		}
		// Office COM servers are often launched through an out-of-process broker,
		// not as children of PowerShell.  Killing just PowerShell can leave a
		// modal WINWORD/POWERPNT/EXCEL automation instance behind; every later
		// COM call then blocks behind that instance.  Remove only processes whose
		// command line identifies them as Office automation servers, never a
		// normal interactive Office window.
		terminateOfficeAutomationServers(killGrace)
		if killGrace <= 0 {
			return fmt.Errorf("%s timed out after %s (process cleanup continues asynchronously)", label, timeout)
		}
		select {
		case <-done:
			return fmt.Errorf("%s timed out after %s", label, timeout)
		case <-time.After(killGrace):
			return fmt.Errorf("%s timed out after %s; process cleanup exceeded %s and continues asynchronously", label, timeout, killGrace)
		}
	}
}

// newOfficeBaselineRunID is passed to the PowerShell baseline process so a
// timed-out runner can remove an orphan even after Windows re-parents it.
// The value is internal and contains only command-line-safe characters.
func newOfficeBaselineRunID() string {
	return fmt.Sprintf("officebaseline-%d-%d", os.Getpid(), time.Now().UnixNano())
}

func terminateOfficeBaselineRun(runID string, grace time.Duration) {
	if runtime.GOOS != "windows" || runID == "" {
		return
	}
	// RunId is generated locally (not a filename or user input). Exclude this
	// cleanup PowerShell process so it cannot terminate itself while scanning.
	const script = "param([string]$RunId); $ErrorActionPreference='SilentlyContinue'; Get-CimInstance Win32_Process -Filter \"Name='powershell.exe'\" | Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -like ('*' + $RunId + '*') } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
	cmd := exec.Command("powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script, "-RunId", runID)
	if err := cmd.Start(); err != nil {
		return
	}
	if grace <= 0 {
		return
	}
	done := make(chan struct{})
	go func() { _ = cmd.Wait(); close(done) }()
	select {
	case <-done:
	case <-time.After(grace):
		_ = cmd.Process.Kill()
	}
}

func terminateOfficeAutomationServers(grace time.Duration) {
	if runtime.GOOS != "windows" {
		return
	}
	// Constant script: it does not interpolate a filename or user input. The
	// command-line filter protects interactive Office use while cleaning the
	// orphaned /Automation -Embedding servers created by this test command.
	const script = "$ErrorActionPreference='SilentlyContinue'; Get-CimInstance Win32_Process -Filter \"Name='WINWORD.EXE' OR Name='POWERPNT.EXE' OR Name='EXCEL.EXE'\" | Where-Object { $_.CommandLine -match '(?i)(/automation|-embedding)' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
	cmd := exec.Command("powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script)
	if err := cmd.Start(); err != nil {
		return
	}
	if grace <= 0 {
		return
	}
	done := make(chan struct{})
	go func() { _ = cmd.Wait(); close(done) }()
	select {
	case <-done:
	case <-time.After(grace):
		_ = cmd.Process.Kill()
	}
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

func missingTokenSegments(mode comparisonMode, segments []officeTextSegment, candidate string) []missingTokenSegment {
	if len(segments) == 0 {
		return nil
	}
	candidateCounts := tokenCountsForMode(mode, candidate)
	// A sheet's UsedRange can expose tens of thousands of rows.  Segment-level
	// diagnostics are intended to make a mismatch actionable, not to duplicate
	// the whole workbook in the checkpoint JSON.  Keep a bounded representative
	// sample; aggregate token scores above remain the complete evidence.
	const maxDiagnosticSegments = 40
	var out []missingTokenSegment
	for _, segment := range segments {
		values := tokenStreamForMode(mode, segment.Text)
		var missing []string
		for _, token := range values {
			// Segment text can be repeated across a deck (footers, template
			// labels, and recurring bullets). Allocating a global token count to
			// segments in traversal order makes later, otherwise matching shapes
			// look missing. Report only tokens absent from the candidate entirely;
			// this gives the context field an evidence-backed meaning.
			if candidateCounts[token] == 0 {
				missing = appendDiagnosticTokens(missing, token, 1)
			}
		}
		if len(missing) > 0 {
			sort.Strings(missing)
			out = append(out, missingTokenSegment{Index: segment.Index, Context: segment.Context, MissingTokens: missing})
			if len(out) >= maxDiagnosticSegments {
				break
			}
		}
	}
	return out
}

func tokenCountsForMode(mode comparisonMode, text string) map[string]int {
	return tokenCountsFromStream(tokenStreamForMode(mode, text))
}

func tokenStreamForMode(mode comparisonMode, text string) []string {
	if mode == comparisonModeFormula {
		return formulaTokenStream(text)
	}
	return tokenStream(text)
}

func comparisonModeForPath(path string) comparisonMode {
	ext := strings.ToLower(filepath.Ext(path))
	if ext == ".docx" && docxContainsOfficeMath(path) {
		return comparisonModeFormula
	}
	// PowerPoint's TextRange flattens DrawingML Office Math in rendered visual
	// order (scripts, fractions and matrices), whereas the package stores its
	// semantic OOXML tree order.  Treat this precisely like Word Office Math:
	// compare the visible mathematical symbols, not an accidental tree walk
	// order.  Ordinary PPTX text remains in the strict prose mode above.
	if ext == ".pptx" && pptxContainsOfficeMath(path) {
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

func pptxContainsOfficeMath(path string) bool {
	zr, err := zip.OpenReader(path)
	if err != nil {
		return false
	}
	defer zr.Close()
	for _, f := range zr.File {
		name := strings.ToLower(f.Name)
		if !strings.HasPrefix(name, "ppt/slides/") || !strings.HasSuffix(name, ".xml") {
			continue
		}
		r, err := f.Open()
		if err != nil {
			continue
		}
		data, readErr := io.ReadAll(r)
		_ = r.Close()
		if readErr == nil && bytes.Contains(data, []byte(":oMath")) {
			return true
		}
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
		if isOfficePasswordProtected(result.Error) {
			return "office-password-protected"
		}
		if isOfficePolicyBlocked(result.Error) {
			return "office-policy-blocked"
		}
		if isOfficeSessionUnavailable(result.Error) {
			return "office-session-unavailable"
		}
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

func isOfficePasswordProtected(message string) bool {
	return strings.Contains(strings.ToLower(message), "password-protected office package")
}

// isOfficeSessionUnavailable recognizes Windows COM activation failures that
// identify a vanished interactive logon session (HRESULT 0x80070520).  This
// is environment state, not a document-level extraction or Office-content
// failure. Keeping it distinct lets a later recovered desktop session retry
// exactly these paths without hiding genuine per-file automation failures.
func isOfficeSessionUnavailable(message string) bool {
	lower := strings.ToLower(message)
	return strings.Contains(lower, "80070520") ||
		strings.Contains(lower, "specified logon session does not exist") ||
		strings.Contains(message, "指定的登录会话不存在")
}

func isOfficePolicyBlocked(message string) bool {
	value := strings.ToLower(message)
	return strings.Contains(value, "file block") ||
		strings.Contains(value, "trust center") ||
		strings.Contains(message, "文件阻止") ||
		strings.Contains(message, "信任中心")
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
	// Close alone is not a durability boundary.  Flush the complete JSON before
	// replacing a long-running audit checkpoint so an abrupt machine/process
	// interruption cannot publish a truncated report.
	if err := tmp.Sync(); err != nil {
		tmp.Close()
		return err
	}
	if err := tmp.Close(); err != nil {
		return err
	}
	// Antivirus/indexing processes can briefly hold a report open on Windows.
	// os.Rename replaces an existing file on Windows; do not delete the old
	// checkpoint first.  Deleting it creates an observable gap in which a
	// concurrent audit reader sees a missing or half-replaced report, defeating
	// the reason for writing the new checkpoint through a temporary file.
	// Retrying the replace prevents a healthy supervisor from aborting merely
	// because the report was transiently held open.
	var lastErr error
	for attempt := 0; attempt < 20; attempt++ {
		if err := replaceFileAtomically(tmpName, path); err == nil {
			return nil
		} else {
			lastErr = err
		}
		time.Sleep(150 * time.Millisecond)
	}
	return lastErr
}

// replaceFileAtomically exists to make the Windows replacement semantics
// explicit. os.Rename is atomic on supported local filesystems, but it can
// fail when a report reader or indexer holds the destination open; callers
// retry the complete replacement rather than deleting the durable checkpoint.
func replaceFileAtomically(source, destination string) error {
	return os.Rename(source, destination)
}
func fatal(err error) { fmt.Fprintln(os.Stderr, err); os.Exit(1) }
