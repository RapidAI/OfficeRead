package main

import (
	"encoding/csv"
	"encoding/json"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"runtime/debug"
	"sort"
	"strings"
	"sync"
	"time"

	"officeread"
)

var supportedExts = map[string]bool{
	".doc":  true,
	".docx": true,
	".ppt":  true,
	".pptx": true,
	".xls":  true,
	".xlsx": true,
}

type FileResult struct {
	Path           string  `json:"path"`
	Ext            string  `json:"ext"`
	OK             bool    `json:"ok"`
	Empty          bool    `json:"empty"`
	TextBytes      int     `json:"textBytes"`
	MarkdownBytes  int     `json:"markdownBytes"`
	Images         int     `json:"images"`
	MarkdownImages int     `json:"markdownImages"`
	Error          string  `json:"error,omitempty"`
	Panic          string  `json:"panic,omitempty"`
	Millis         int64   `json:"millis"`
	MinMillis      int64   `json:"minMillis,omitempty"`
	MaxMillis      int64   `json:"maxMillis,omitempty"`
	Runs           []int64 `json:"runs,omitempty"`
}

type ExtSummary struct {
	Total          int   `json:"total"`
	OK             int   `json:"ok"`
	Errors         int   `json:"errors"`
	Panics         int   `json:"panics"`
	Empty          int   `json:"empty"`
	TextBytes      int64 `json:"textBytes"`
	Images         int64 `json:"images"`
	MarkdownBytes  int64 `json:"markdownBytes"`
	MarkdownImages int64 `json:"markdownImages"`
	Millis         int64 `json:"millis"`
	MaxMillis      int64 `json:"maxMillis"`
	Over10Sec      int   `json:"over10Sec"`
}

type Report struct {
	StartedAt  string                `json:"startedAt"`
	FinishedAt string                `json:"finishedAt"`
	Inputs     []string              `json:"inputs"`
	Summary    map[string]ExtSummary `json:"summary"`
	Files      []FileResult          `json:"files"`
}

func main() {
	jsonOut := flag.String("json", "compat-report.json", "JSON report path")
	csvOut := flag.String("csv", "compat-report.csv", "CSV report path")
	limitPerExt := flag.Int("limit", 0, "maximum files per extension; 0 means unlimited")
	repeat := flag.Int("repeat", 1, "number of times to run each file; millis reports the median when repeat > 1")
	markdown := flag.Bool("markdown", false, "also render Markdown and verify its image references")
	jobs := flag.Int("jobs", 1, "number of files to check concurrently")
	flag.Parse()
	if flag.NArg() == 0 {
		fmt.Fprintln(os.Stderr, "usage: compatcheck [-json report.json] [-csv report.csv] [-limit n] [-repeat n] [-jobs n] [-markdown] file-or-directory [...]")
		os.Exit(2)
	}
	if *repeat <= 0 {
		fmt.Fprintln(os.Stderr, "-repeat must be >= 1")
		os.Exit(2)
	}
	if *jobs <= 0 {
		fmt.Fprintln(os.Stderr, "-jobs must be >= 1")
		os.Exit(2)
	}

	files, err := collectFiles(flag.Args(), *limitPerExt)
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	report := Report{
		StartedAt: time.Now().Format(time.RFC3339),
		Inputs:    flag.Args(),
		Summary:   map[string]ExtSummary{},
	}
	for _, res := range checkFiles(files, *repeat, *markdown, *jobs) {
		report.Files = append(report.Files, res)
		s := report.Summary[res.Ext]
		s.Total++
		s.Millis += res.Millis
		if res.Millis > s.MaxMillis {
			s.MaxMillis = res.Millis
		}
		if res.Millis > 10000 {
			s.Over10Sec++
		}
		if res.OK {
			s.OK++
		}
		if res.Error != "" {
			s.Errors++
		}
		if res.Panic != "" {
			s.Panics++
		}
		if res.Empty {
			s.Empty++
		}
		s.TextBytes += int64(res.TextBytes)
		s.Images += int64(res.Images)
		s.MarkdownBytes += int64(res.MarkdownBytes)
		s.MarkdownImages += int64(res.MarkdownImages)
		report.Summary[res.Ext] = s
		if len(res.Runs) > 1 {
			fmt.Printf("%s ok=%v empty=%v text=%d images=%d err=%q panic=%q ms=%d min=%d max=%d runs=%v\n", res.Path, res.OK, res.Empty, res.TextBytes, res.Images, res.Error, res.Panic, res.Millis, res.MinMillis, res.MaxMillis, res.Runs)
		} else {
			fmt.Printf("%s ok=%v empty=%v text=%d images=%d err=%q panic=%q ms=%d\n", res.Path, res.OK, res.Empty, res.TextBytes, res.Images, res.Error, res.Panic, res.Millis)
		}
	}
	report.FinishedAt = time.Now().Format(time.RFC3339)

	if err := writeJSON(*jsonOut, report); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	if err := writeCSV(*csvOut, report.Files); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	if hasFailure(report) {
		os.Exit(1)
	}
}

func checkFiles(files []string, repeat int, markdown bool, jobs int) []FileResult {
	results := make([]FileResult, len(files))
	if len(files) == 0 {
		return results
	}
	if jobs > len(files) {
		jobs = len(files)
	}
	indexes := make(chan int)
	var workers sync.WaitGroup
	for i := 0; i < jobs; i++ {
		workers.Add(1)
		go func() {
			defer workers.Done()
			for index := range indexes {
				results[index] = checkOne(files[index], repeat, markdown)
			}
		}()
	}
	for i := range files {
		indexes <- i
	}
	close(indexes)
	workers.Wait()
	return results
}

func collectFiles(inputs []string, limitPerExt int) ([]string, error) {
	var files []string
	counts := map[string]int{}
	for _, input := range inputs {
		info, err := os.Stat(input)
		if err != nil {
			return nil, err
		}
		if !info.IsDir() {
			addFile(&files, counts, input, limitPerExt)
			continue
		}
		err = filepath.WalkDir(input, func(path string, d os.DirEntry, err error) error {
			if err != nil {
				return err
			}
			if !d.IsDir() {
				addFile(&files, counts, path, limitPerExt)
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

func addFile(files *[]string, counts map[string]int, path string, limitPerExt int) {
	ext := strings.ToLower(filepath.Ext(path))
	if !supportedExts[ext] {
		return
	}
	if limitPerExt > 0 && counts[ext] >= limitPerExt {
		return
	}
	counts[ext]++
	*files = append(*files, path)
}

func checkOne(path string, repeat int, markdown bool) (result FileResult) {
	if repeat <= 1 {
		return checkOneRun(path, markdown)
	}
	result = checkOneRun(path, markdown)
	result.Runs = []int64{result.Millis}
	result.MinMillis = result.Millis
	result.MaxMillis = result.Millis
	for i := 1; i < repeat; i++ {
		run := checkOneRun(path, markdown)
		result.Runs = append(result.Runs, run.Millis)
		if run.Millis < result.MinMillis {
			result.MinMillis = run.Millis
		}
		if run.Millis > result.MaxMillis {
			result.MaxMillis = run.Millis
		}
		result.OK = result.OK && run.OK
		result.Empty = result.Empty && run.Empty
		if result.Error == "" && run.Error != "" {
			result.Error = run.Error
		}
		if result.Panic == "" && run.Panic != "" {
			result.Panic = run.Panic
		}
		if run.TextBytes != result.TextBytes || run.Images != result.Images || run.MarkdownBytes != result.MarkdownBytes || run.MarkdownImages != result.MarkdownImages {
			if result.Error == "" {
				result.Error = fmt.Sprintf("repeat output mismatch: first text=%d markdown=%d images=%d markdownImages=%d, run %d text=%d markdown=%d images=%d markdownImages=%d", result.TextBytes, result.MarkdownBytes, result.Images, result.MarkdownImages, i+1, run.TextBytes, run.MarkdownBytes, run.Images, run.MarkdownImages)
			}
			result.OK = false
		}
	}
	result.Millis = medianMillis(result.Runs)
	return result
}

func checkOneRun(path string, markdown bool) (result FileResult) {
	result.Path = path
	result.Ext = strings.ToLower(filepath.Ext(path))
	start := time.Now()
	defer func() {
		result.Millis = time.Since(start).Milliseconds()
		if r := recover(); r != nil {
			result.Panic = fmt.Sprint(r)
			result.Error = string(debug.Stack())
		}
	}()
	res, err := officeread.Extract(path, officeread.Options{})
	if err != nil {
		result.Error = err.Error()
		return result
	}
	result.OK = true
	result.TextBytes = len(res.Text)
	result.Images = len(res.Images)
	result.Empty = strings.TrimSpace(res.Text) == "" && len(res.Images) == 0
	if markdown {
		md := res.Markdown("images")
		result.MarkdownBytes = len(md)
		result.MarkdownImages = markdownImageReferenceCount(md)
		if result.MarkdownImages != result.Images {
			result.Error = fmt.Sprintf("markdown image reference mismatch: images=%d references=%d", result.Images, result.MarkdownImages)
			result.OK = false
		}
	}
	return result
}

func markdownImageReferenceCount(markdown string) int {
	count := 0
	for len(markdown) > 0 {
		i := strings.Index(markdown, "![")
		if i < 0 {
			break
		}
		markdown = markdown[i+2:]
		closeAlt := strings.IndexByte(markdown, ']')
		if closeAlt < 0 || closeAlt+1 >= len(markdown) || markdown[closeAlt+1] != '(' {
			continue
		}
		closeURL := strings.IndexByte(markdown[closeAlt+2:], ')')
		if closeURL < 0 {
			continue
		}
		count++
		markdown = markdown[closeAlt+3+closeURL:]
	}
	return count
}

func writeJSON(path string, report Report) error {
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil && filepath.Dir(path) != "." {
		return err
	}
	data, err := json.MarshalIndent(report, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(path, data, 0o644)
}

func writeCSV(path string, files []FileResult) error {
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil && filepath.Dir(path) != "." {
		return err
	}
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	w := csv.NewWriter(f)
	defer w.Flush()
	if err := w.Write([]string{"path", "ext", "ok", "empty", "textBytes", "markdownBytes", "images", "markdownImages", "error", "panic", "millis", "minMillis", "maxMillis", "runs"}); err != nil {
		return err
	}
	for _, r := range files {
		runs := ""
		if len(r.Runs) > 0 {
			parts := make([]string, 0, len(r.Runs))
			for _, ms := range r.Runs {
				parts = append(parts, fmt.Sprint(ms))
			}
			runs = strings.Join(parts, "|")
		}
		if err := w.Write([]string{
			r.Path,
			r.Ext,
			fmt.Sprint(r.OK),
			fmt.Sprint(r.Empty),
			fmt.Sprint(r.TextBytes),
			fmt.Sprint(r.MarkdownBytes),
			fmt.Sprint(r.Images),
			fmt.Sprint(r.MarkdownImages),
			r.Error,
			r.Panic,
			fmt.Sprint(r.Millis),
			fmt.Sprint(r.MinMillis),
			fmt.Sprint(r.MaxMillis),
			runs,
		}); err != nil {
			return err
		}
	}
	return nil
}

func hasFailure(report Report) bool {
	for _, r := range report.Files {
		if !r.OK {
			return true
		}
	}
	return false
}

func medianMillis(values []int64) int64 {
	if len(values) == 0 {
		return 0
	}
	sortedValues := append([]int64(nil), values...)
	sort.Slice(sortedValues, func(i, j int) bool { return sortedValues[i] < sortedValues[j] })
	mid := len(sortedValues) / 2
	if len(sortedValues)%2 == 1 {
		return sortedValues[mid]
	}
	return (sortedValues[mid-1] + sortedValues[mid]) / 2
}
