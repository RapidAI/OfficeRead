package main

import (
	"os"
	"path/filepath"
	"testing"
)

func TestMarkdownImageReferenceCount(t *testing.T) {
	markdown := "plain ![first](images/one.png) text ![broken] and ![asset\\[1\\]](images/one\\)copy.png) and ![second](images/two.png)"
	if got, want := markdownImageReferenceCount(markdown), 3; got != want {
		t.Fatalf("markdownImageReferenceCount() = %d, want %d", got, want)
	}
}

func TestCheckFilesPreservesInputOrder(t *testing.T) {
	dir := t.TempDir()
	paths := []string{
		filepath.Join(dir, "first.txt"),
		filepath.Join(dir, "second.txt"),
	}
	for _, path := range paths {
		if err := os.WriteFile(path, []byte("not an Office file"), 0o600); err != nil {
			t.Fatal(err)
		}
	}
	results := checkFiles(paths, 1, false, 2)
	if len(results) != len(paths) {
		t.Fatalf("got %d results, want %d", len(results), len(paths))
	}
	for i, result := range results {
		if result.Path != paths[i] {
			t.Errorf("result %d path = %q, want %q", i, result.Path, paths[i])
		}
	}
}

func TestProgressWriterRecordsFinalPartialBatch(t *testing.T) {
	dir := t.TempDir()
	p := &progressWriter{
		jsonPath: filepath.Join(dir, "report.json"),
		csvPath:  filepath.Join(dir, "report.csv"),
		report:   &Report{Summary: map[string]ExtSummary{}},
		results:  make([]FileResult, 3),
		done:     make([]bool, 3),
	}
	for i := range p.results {
		p.record(i, FileResult{Path: filepath.Join(dir, "sample"), Ext: ".pptx", OK: true})
	}
	if _, err := os.Stat(p.jsonPath); err != nil {
		t.Fatalf("final partial checkpoint was not written: %v", err)
	}
}
