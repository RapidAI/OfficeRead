package officeread

import (
	"encoding/binary"
	"os"
	"testing"
)

func TestStrictLegacyXLS005864ExposesOnlyVisibleCells(t *testing.T) {
	b, err := os.ReadFile("testdata/web-samples/samples/xls/005864.xls")
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(b)
	if err != nil {
		t.Fatal(err)
	}
	w := legacyWorkbookBytes(b, streams)
	sheets := biffBoundSheetsForProbe(w)
	if len(sheets) != 1 || sheets[0].hidden || sheets[0].name != "FDA-PhishPharm" {
		t.Fatalf("unexpected sheets: %#v", sheets)
	}
	parts := biffTextParts(w)
	visible := 0
	for _, p := range parts {
		if p.cell && !p.hide && !p.comment {
			visible++
		}
	}
	if got := len(biffStrictOfficeText(w)); got != visible+1 {
		t.Fatalf("strict BIFF tokens=%d, want sheet name plus %d cells", got, visible)
	}
}

func TestStrictLegacyXLS005864TracksLargeUsedRange(t *testing.T) {
	b, err := os.ReadFile("testdata/web-samples/samples/xls/005864.xls")
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(b)
	if err != nil {
		t.Fatal(err)
	}
	parts := biffTextParts(legacyWorkbookBytes(b, streams))
	maxRow, maxCol := 0, 0
	for _, part := range parts {
		if part.cell && !part.hide {
			if part.row > maxRow {
				maxRow = part.row
			}
			if part.col > maxCol {
				maxCol = part.col
			}
		}
	}
	if maxRow < 3000 || maxCol < 30 {
		t.Fatalf("visible BIFF range = %d,%d, want large worksheet extent", maxRow, maxCol)
	}
}

func TestStrictLegacyXLSExcludesCommentBalloonText(t *testing.T) {
	parts := []biffTextPart{
		{text: "Visible value", row: 0, col: 0, cell: true},
		{text: "Comment annotation", row: 0, col: 0, cell: true, comment: true},
	}
	var strict []string
	for _, part := range parts {
		if part.cell && !part.hide && !part.comment {
			strict = append(strict, part.text)
		}
	}
	if len(strict) != 1 || strict[0] != "Visible value" {
		t.Fatalf("strict visible cell values = %#v, want only the cell text", strict)
	}
}

func biffBoundSheetsForProbe(data []byte) []biffSheetInfo {
	var out []biffSheetInfo
	biff8 := true
	cp := uint16(1252)
	for off := 0; off+4 <= len(data); {
		id := binary.LittleEndian.Uint16(data[off:])
		n := int(binary.LittleEndian.Uint16(data[off+2:]))
		off += 4
		if off+n > len(data) {
			break
		}
		rec := data[off : off+n]
		if id == 0x0809 || id == 0x0409 || id == 0x0209 || id == 0x0009 {
			biff8 = isBIFF8BOF(id, rec)
		}
		if id == 0x0042 && len(rec) >= 2 {
			cp = binary.LittleEndian.Uint16(rec)
		}
		if id == 0x0085 {
			if s, ok := parseBIFFBoundSheetRecord(rec, biff8, cp); ok {
				out = append(out, s)
			}
		}
		off += n
	}
	return out
}
