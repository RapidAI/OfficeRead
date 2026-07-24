package officeread

import (
	"path/filepath"
	"strings"
	"testing"
)

func TestStrictPPTXContentKeepsTemplateLabels(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "pptx", "00020964.pptx")
	result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"单击此处添加标题文字", "小标题", "报告人", "日期"} {
		if !strings.Contains(result.Text, want) {
			t.Fatalf("strict PPTX text missing %q: %q", want, result.Text)
		}
	}
}

func TestStrictPPTXContentIncludesGroupedShapeText(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "pptx", "00024860.pptx")
	result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(result.Text, "Total Consolidation") {
		t.Fatalf("strict PPTX text must include visible grouped text: %q", result.Text)
	}
}

func TestStrictPPTXContentIncludesTextBearingGraphicFrames(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "pptx", "00024860.pptx")
	result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(result.Text, "State of California") {
		t.Fatalf("strict PPTX text must include graphic-frame text: %q", result.Text)
	}
}

func TestStrictOfficeContentExcludesDOCXChartCache(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "LibreOffice__core__chart2_qa_extras_data_docx_data_point_inherited_color.docx")
	compatible, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(compatible.Text, "Column 1") {
		t.Fatalf("compatibility text lost chart cache: %q", compatible.Text)
	}
	strict, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if strict.Text != "" {
		t.Fatalf("strict document content = %q, want empty", strict.Text)
	}
}

func TestStrictOfficeContentExcludesEmbeddedChartWorkbook(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "LibreOffice__core__chart2_qa_extras_data_docx_3d-bar-label.docx")
	compatible, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(compatible.Text, "Series 3") {
		t.Fatalf("compatibility text lost embedded chart workbook: %q", compatible.Text)
	}
	strict, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if strict.Text != "" {
		t.Fatalf("strict document content = %q, want empty", strict.Text)
	}
}
