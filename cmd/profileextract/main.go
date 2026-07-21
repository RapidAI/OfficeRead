package main

import (
	"flag"
	"fmt"
	"os"
	"runtime/pprof"
	"time"

	"officeread"
)

func main() {
	out := flag.String("out", "cpu.pprof", "cpu profile output path")
	flag.Parse()
	if flag.NArg() != 1 {
		fmt.Fprintln(os.Stderr, "usage: profileextract -out cpu.pprof file")
		os.Exit(2)
	}

	f, err := os.Create(*out)
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	defer f.Close()

	if err := pprof.StartCPUProfile(f); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	start := time.Now()
	res, err := officeread.Extract(flag.Arg(0), officeread.Options{})
	pprof.StopCPUProfile()
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}

	fmt.Printf("path=%s\n", flag.Arg(0))
	fmt.Printf("text=%d images=%d ms=%d\n", len(res.Text), len(res.Images), time.Since(start).Milliseconds())
}
