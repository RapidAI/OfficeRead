package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"unicode"

	"github.com/RapidAI/OfficeRead"
)

var officeExts = map[string]bool{
	".doc":  true,
	".docx": true,
	".ppt":  true,
	".pptx": true,
	".xls":  true,
	".xlsx": true,
}

func main() {
	outDir := flag.String("out", "extract-output", "directory for extracted text and images")
	flag.Parse()

	if flag.NArg() == 0 {
		fmt.Fprintln(os.Stderr, "usage: extracttest [-out dir] file-or-directory [...]")
		os.Exit(2)
	}

	files, err := collectOfficeFiles(flag.Args())
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	if len(files) == 0 {
		fmt.Fprintln(os.Stderr, "no supported Office files found")
		os.Exit(1)
	}

	if err := os.MkdirAll(*outDir, 0o755); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}

	seenNames := map[string]int{}
	var failed int
	for _, file := range files {
		if err := extractOne(file, *outDir, seenNames); err != nil {
			failed++
			fmt.Fprintf(os.Stderr, "%s: %v\n", file, err)
		}
	}
	if failed > 0 {
		os.Exit(1)
	}
}

func collectOfficeFiles(inputs []string) ([]string, error) {
	var files []string
	for _, input := range inputs {
		info, err := os.Stat(input)
		if err != nil {
			return nil, err
		}
		if !info.IsDir() {
			if isOfficeFile(input) {
				files = append(files, input)
			}
			continue
		}
		err = filepath.WalkDir(input, func(path string, d os.DirEntry, err error) error {
			if err != nil {
				return err
			}
			if !d.IsDir() && isOfficeFile(path) {
				files = append(files, path)
			}
			return nil
		})
		if err != nil {
			return nil, err
		}
	}
	return files, nil
}

func extractOne(file, outDir string, seenNames map[string]int) error {
	dirName := uniqueDirName(safeName(filepath.Base(file)), seenNames)
	targetDir := filepath.Join(outDir, dirName)
	imagesDir := filepath.Join(targetDir, "images")

	res, err := officeread.Extract(file, officeread.Options{ImageDir: imagesDir})
	if err != nil {
		return err
	}
	if err := os.MkdirAll(targetDir, 0o755); err != nil {
		return err
	}
	if err := os.WriteFile(filepath.Join(targetDir, "text.txt"), []byte(res.Text), 0o644); err != nil {
		return err
	}
	if err := os.WriteFile(filepath.Join(targetDir, "markdown.md"), []byte(res.Markdown("images")), 0o644); err != nil {
		return err
	}

	fmt.Printf("%s\n", file)
	fmt.Printf("  text:   %s (%d bytes)\n", filepath.Join(targetDir, "text.txt"), len(res.Text))
	fmt.Printf("  md:     %s\n", filepath.Join(targetDir, "markdown.md"))
	fmt.Printf("  images: %s (%d files)\n", imagesDir, len(res.Images))
	for _, img := range res.Images {
		name := img.Name
		if name == "" {
			name = "(generated name)"
		}
		fmt.Printf("    - %s (%d bytes)\n", name, len(img.Data))
	}
	return nil
}

func isOfficeFile(name string) bool {
	return officeExts[strings.ToLower(filepath.Ext(name))]
}

func uniqueDirName(name string, seen map[string]int) string {
	key := strings.ToLower(name)
	seen[key]++
	if seen[key] == 1 {
		return name
	}
	return fmt.Sprintf("%s-%d", name, seen[key])
}

func safeName(name string) string {
	var b strings.Builder
	for _, r := range name {
		if unicode.IsLetter(r) || unicode.IsDigit(r) || r == '-' || r == '_' || r == '.' {
			b.WriteRune(r)
			continue
		}
		b.WriteByte('_')
	}
	s := strings.Trim(b.String(), "._")
	if s == "" {
		return "document"
	}
	return s
}
