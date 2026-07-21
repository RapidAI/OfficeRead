package main

import (
	"flag"
	"fmt"
	"os"

	"officeread"
)

func main() {
	imageDir := flag.String("images", "", "directory for extracted images")
	includeMetadata := flag.Bool("metadata", false, "include document properties, relationships, and custom XML")
	markdown := flag.Bool("markdown", false, "print markdown with image references")
	textOnly := flag.Bool("text-only", false, "print text only")
	flag.Parse()
	if flag.NArg() != 1 {
		fmt.Fprintln(os.Stderr, "usage: officeread [-images dir] [-metadata] [-markdown] [-text-only] file")
		os.Exit(2)
	}
	if *markdown && *imageDir == "" {
		*imageDir = "images"
	}
	res, err := officeread.Extract(flag.Arg(0), officeread.Options{ImageDir: *imageDir, IncludeMetadata: *includeMetadata})
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	if *markdown {
		fmt.Print(res.Markdown(*imageDir))
		if !*textOnly {
			if len(res.Images) > 0 {
				fmt.Println()
			}
			fmt.Fprintf(os.Stderr, "images: %d\n", len(res.Images))
		}
		return
	}
	fmt.Print(res.Text)
	if !*textOnly {
		if res.Text != "" {
			fmt.Println()
		}
		fmt.Fprintf(os.Stderr, "images: %d\n", len(res.Images))
	}
}
