package officeread

import (
	"archive/zip"
	"strings"
	"testing"
)

func TestStrictPPTXSymbolPrivateUseRunsMatchPowerPointText(t *testing.T) {
	mapped, ok := fontEncodedSymbolRune("Symbol", 0xae)
	if !isFontEncodedSymbolFont("Symbol") || !ok || mapped != '→' || pptxSymbolFontText("\uf0aeY", "Symbol") != "→Y" {
		t.Fatalf("Symbol private-use mapping failed: %v %q %v %q", isFontEncodedSymbolFont("Symbol"), mapped, ok, pptxSymbolFontText("\uf0aeY", "Symbol"))
	}
	z, err := zip.OpenReader("testdata/web-samples/samples/pptx/00025755.pptx")
	if err != nil {
		t.Fatal(err)
	}
	defer z.Close()
	for _, f := range z.File {
		if f.Name != "ppt/slides/slide2.xml" {
			continue
		}
		b, err := readZipFile(f)
		if err != nil {
			t.Fatal(err)
		}
		s, err := visiblePptxShapeText(b)
		if err != nil {
			t.Fatal(err)
		}
		if !strings.Contains(s, "X→Y") || !strings.Contains(s, "U→W") {
			t.Fatalf("PowerPoint symbol runs were not rendered: %q", s)
		}
	}
}

func TestStrictPPTXOfficeMathRunBoundariesMatchPowerPointAtoms(t *testing.T) {
	// PowerPoint TextRange surfaces adjacent Office Math atoms separately. This
	// real presentation has m:r runs for K and m; joining them changes two
	// visible mathematical identifiers into the unrelated token "Km".
	z, err := zip.OpenReader("testdata/web-samples/samples/pptx/00020158.pptx")
	if err != nil {
		t.Fatal(err)
	}
	defer z.Close()
	var found bool
	for _, f := range z.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		b, err := readZipFile(f)
		if err != nil {
			t.Fatal(err)
		}
		s, err := visiblePptxShapeText(b)
		if err != nil {
			t.Fatal(err)
		}
		if strings.Contains(s, "𝑲 𝒎") {
			found = true
		}
		if strings.Contains(s, "𝑲𝒎") {
			t.Fatalf("Office Math atoms were joined in %s: %q", f.Name, s)
		}
	}
	if !found {
		t.Fatal("fixture no longer contains separated Office Math atoms")
	}
}

func TestStrictPPTXOfficeMathKeepsMultiLetterRunsWhole(t *testing.T) {
	if !pptxMathRunIsAtom("𝑲") || pptxMathRunIsAtom("𝑚𝑎𝑥") || pptxMathRunIsAtom("properties") {
		t.Fatal("Office Math atom classifier must split one glyph but preserve multi-letter runs")
	}
	z, err := zip.OpenReader("testdata/web-samples/samples/pptx/00020433.pptx")
	if err != nil {
		t.Fatal(err)
	}
	defer z.Close()
	for _, f := range z.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		b, err := readZipFile(f)
		if err != nil {
			t.Fatal(err)
		}
		s, err := visiblePptxShapeText(b)
		if err != nil {
			t.Fatal(err)
		}
		if strings.Contains(s, "𝑎 𝑚𝑎𝑥") || strings.Contains(s, "𝑝𝑟𝑜𝑝𝑒𝑟𝑡𝑖𝑒𝑠 𝑥") {
			t.Fatalf("Office Math multi-letter run was incorrectly split in %s: %q", f.Name, s)
		}
	}
}
