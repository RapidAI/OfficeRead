package officeread

import (
	"encoding/binary"
	"os"
	"testing"
)

func TestStrictLegacyDOCFallsBackFromMateriallyTruncatedCLX(t *testing.T) {
	for _, name := range []string{"005839.doc", "005664.doc", "004192.doc"} {
		b, err := os.ReadFile("testdata/web-samples/samples/doc/" + name)
		if err != nil { t.Fatal(err) }
		streams, err := readOLEStreams(b); if err != nil { t.Fatal(err) }
		word, ok := findLegacyStream(streams,"WordDocument"); if !ok { t.Fatal("no word") }
		table,_ := findLegacyStream(streams,wordTableStreamName(word.Data))
		ccp:=uint32(0); if len(word.Data)>0x50 {ccp=binary.LittleEndian.Uint32(word.Data[0x4c:])}
		main := wordMainStoryText(word.Data,table.Data)
		fcClx:=int(binary.LittleEndian.Uint32(word.Data[0x1a2:])); lcbClx:=int(binary.LittleEndian.Uint32(word.Data[0x1a6:])); raw:=[]string(nil); if fcClx>=0&&lcbClx>0&&fcClx+lcbClx<=len(table.Data){raw=parseWordCLXTextUntilCP(word.Data,table.Data[fcClx:fcClx+lcbClx],ccp,true)}
		if !wordMainStoryCoverageMateriallyIncomplete(raw, ccp) {
			t.Fatalf("%s fixture must have materially truncated CLX: ccp=%d raw=%d", name, ccp, wordTextByteCount(raw))
		}
		if wordTextByteCount(main) <= wordTextByteCount(raw)*4 {
			t.Fatalf("%s main-story fallback did not restore substantially more text: raw=%d main=%d", name, wordTextByteCount(raw), wordTextByteCount(main))
		}
		result, err := Extract("testdata/web-samples/samples/doc/"+name, Options{StrictOfficeContent:true})
		if err != nil { t.Fatal(err) }
		if len(result.Text) <= wordTextByteCount(raw)*4 {
			t.Fatalf("%s strict result did not preserve recovered main story", name)
		}
	}
}
