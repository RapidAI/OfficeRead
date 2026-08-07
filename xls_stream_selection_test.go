package officeread

import (
	"os"
	"strings"
	"testing"
)

func TestLegacyXLSWorkbookSelectionHandlesBookAndWorkbookStreams(t *testing.T) {
	for _, tc := range []struct {
		file string
		want string
	}{
		{"001898.xls", "Table 668"},
		{"005560.xls", "AQUACULTURE"},
		{"001990.xls", "Illinois Emergency Management Agency"},
		{"002219.xls", "ALASKA SEAFOODS"},
	} {
		data, err := os.ReadFile("testdata/web-samples/samples/xls/" + tc.file)
		if err != nil {
			t.Fatal(err)
		}
		streams, err := readOLEStreams(data)
		if err != nil {
			t.Fatal(err)
		}
		workbook := legacyWorkbookBytes(data, streams)
		if tc.file == "009941.xls" {
			parts:=biffTextParts(workbook); cells,hidden:=0,0; for _,p:=range parts {if p.cell {cells++;if p.hide{hidden++}}}; t.Logf("%s parsed parts=%d cells=%d hidden=%d",tc.file,len(parts),cells,hidden)
		}
		if got := joinStrictBIFFOfficeText(biffStrictOfficeText(workbook)); !containsFold(got, tc.want) {
			t.Fatalf("%s selected workbook misses %q", tc.file, tc.want)
		}
	}
}

func containsFold(s, needle string) bool {
	return strings.Contains(strings.ToUpper(s), strings.ToUpper(needle))
}
