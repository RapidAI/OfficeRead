package officeread

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestPPTActivePersistOffsetsUsesCurrentUserDirectory(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "000008.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	active := pptActivePersistOffsets(streams)
	if len(active) != 32 {
		t.Fatalf("active persist entries = %d, want 32", len(active))
	}
	if active[1] != 0 || active[40] != 24152 || active[81] != 108406 {
		t.Fatalf("unexpected active persist mapping: %#v", active)
	}
}

func TestPPTActivePersistRecordTypesDiagnostic(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	doc, ok := findLegacyStream(streams, "PowerPoint Document")
	if !ok {
		t.Fatal("missing PowerPoint Document")
	}
	types := map[uint16]int{}
	for _, offset := range pptActivePersistOffsets(streams) {
		off := int(offset)
		if off+8 <= len(doc.Data) {
			types[binary.LittleEndian.Uint16(doc.Data[off+2:])]++
		}
	}
	t.Logf("active persist record types: %#v", types)
}

func TestPPT133PictureFrameDiagnostic(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	for slideIndex, slide := range pptActiveSlideContainers(streams) {
		var walk func([]byte, int)
		walk = func(data []byte, depth int) {
			pptWalkRecords(data, func(options, recordType uint16, payload []byte) {
				if recordType == 0xf004 {
					var fspType uint16
					var flags, pib uint32
					pptWalkRecords(payload, func(opts, typ uint16, value []byte) {
						switch typ {
						case 0xf00a:
							fspType = opts >> 4
							if len(value) >= 8 {
								flags = binary.LittleEndian.Uint32(value[4:])
							}
						case 0xf00b:
							for at := 0; at+6 <= len(value); at += 6 {
								if binary.LittleEndian.Uint16(value[at:])&0x3fff == 0x0104 {
									pib = binary.LittleEndian.Uint32(value[at+2:])
								}
							}
						}
					})
					if pib != 0 || fspType == 75 {
						_, accepted := pptPictureFrameBlip(payload)
						t.Logf("slide=%d depth=%d type=%d flags=%#x pib=%d accepted=%v", slideIndex+1, depth, fspType, flags, pib, accepted)
					}
					return
				}
				if options&0x000f == 0x000f {
					walk(payload, depth+1)
				}
			})
		}
		walk(slide, 0)
	}
}

func TestPPT724MissingGroupTextDiagnostic(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "000724.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	needle := "Within group differences relate to factors of:"
	for slideIndex, slide := range pptActiveSlideContainers(streams) {
		var walk func([]byte, int, []uint16, []uint16)
		walk = func(data []byte, depth int, ancestors, shapes []uint16) {
			pptWalkRecords(data, func(options, typ uint16, payload []byte) {
				if typ == 0x0fa0 || typ == 0x0fa8 {
					text := utf16BytesToStringAll(payload)
					if typ == 0x0fa8 {
						text = compressedUnicodeBytesToString(payload)
					}
					if strings.Contains(text, needle) {
						t.Logf("slide=%d depth=%d type=%#x ancestors=%#v shapes=%#v text=%q", slideIndex+1, depth, typ, ancestors, shapes, text)
					}
				}
				if options&0x000f == 0x000f && depth < 20 {
					nextShapes := append([]uint16(nil), shapes...)
					if typ == 0xf004 {
						nextShapes = append(nextShapes, pptDiagnosticFSPType(payload))
					}
					walk(payload, depth+1, append(append([]uint16(nil), ancestors...), typ), nextShapes)
				}
			})
		}
		walk(slide, 0, nil, nil)
	}
}

func TestPPT133ChartInternalTextDiagnostic(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	for slideIndex, slide := range pptActiveSlideContainers(streams) {
		var walk func([]byte, int, []uint16, []uint16)
		walk = func(data []byte, depth int, ancestors, shapes []uint16) {
			pptWalkRecords(data, func(options, typ uint16, payload []byte) {
				if typ == 0x0fa0 || typ == 0x0fa8 {
					text := utf16BytesToStringAll(payload)
					if typ == 0x0fa8 {
						text = compressedUnicodeBytesToString(payload)
					}
					if strings.Contains(text, "Santa Monica (I)") || strings.Contains(text, "Log mean ENT") || strings.Contains(text, "Frequency of data") {
						t.Logf("slide=%d depth=%d ancestors=%#v shapes=%#v text=%q", slideIndex+1, depth, ancestors, shapes, text)
					}
				}
				if options&0x000f == 0x000f && depth < 20 {
					nextShapes := append([]uint16(nil), shapes...)
					if typ == 0xf004 {
						nextShapes = append(nextShapes, pptDiagnosticFSPType(payload))
					}
					walk(payload, depth+1, append(append([]uint16(nil), ancestors...), typ), nextShapes)
				}
			})
		}
		walk(slide, 0, nil, nil)
	}
}

func TestPPT725VisibleGroupTextDiagnostic(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "000725.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	for slideIndex, slide := range pptActiveSlideContainers(streams) {
		var walk func([]byte, int, []uint32)
		walk = func(data []byte, depth int, shapes []uint32) {
			pptWalkRecords(data, func(options, typ uint16, payload []byte) {
				if typ == 0x0fa0 || typ == 0x0fa8 {
					text := utf16BytesToStringAll(payload)
					if typ == 0x0fa8 {
						text = compressedUnicodeBytesToString(payload)
					}
					if strings.Contains(text, "Extrinsic") || strings.Contains(text, "Illness") || strings.Contains(text, "incubation") || strings.TrimSpace(text) == "1" {
						t.Logf("slide=%d depth=%d shapes=%#v text=%q", slideIndex+1, depth, shapes, text)
					}
				}
				if options&0x000f == 0x000f && depth < 20 {
					next := append([]uint32(nil), shapes...)
					if typ == 0xf004 {
						next = append(next, pptDiagnosticFSP(payload))
					}
					walk(payload, depth+1, next)
				}
			})
		}
		walk(slide, 0, nil)
	}
}

func TestPPT133ChartFSPDiagnostic(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	for slideIndex, slide := range pptActiveSlideContainers(streams) {
		var walk func([]byte, int, []uint32)
		walk = func(data []byte, depth int, shapes []uint32) {
			pptWalkRecords(data, func(options, typ uint16, payload []byte) {
				if typ == 0x0fa0 || typ == 0x0fa8 {
					text := utf16BytesToStringAll(payload)
					if typ == 0x0fa8 {
						text = compressedUnicodeBytesToString(payload)
					}
					if strings.Contains(text, "Santa Monica (I)") {
						t.Logf("slide=%d depth=%d shapes=%#v text=%q", slideIndex+1, depth, shapes, text)
					}
				}
				if options&0x000f == 0x000f && depth < 20 {
					next := append([]uint32(nil), shapes...)
					if typ == 0xf004 {
						next = append(next, pptDiagnosticFSP(payload))
					}
					walk(payload, depth+1, next)
				}
			})
		}
		walk(slide, 0, nil)
	}
}

func TestPPTGroupVersusChartPathDiagnostic(t *testing.T) {
	for _, sample := range []struct{ file, needle string }{
		{"000725.ppt", "Extrinsic"},
		{"000133.ppt", "Santa Monica (I)"},
	} {
		data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", sample.file))
		if err != nil {
			t.Fatal(err)
		}
		streams, err := readOLEStreams(data)
		if err != nil {
			t.Fatal(err)
		}
		for slideIndex, slide := range pptActiveSlideContainers(streams) {
			var walk func([]byte, int, []string)
			walk = func(data []byte, depth int, path []string) {
				pptWalkRecords(data, func(options, typ uint16, payload []byte) {
					descriptor := fmt.Sprintf("%#x/%#x", typ, options)
					if typ == 0xf004 {
						descriptor += fmt.Sprintf("[fsp=%#x]", pptDiagnosticFSP(payload))
					}
					if typ == 0xf003 {
						descriptor += fmt.Sprintf("[children=%d types=%s]", pptRecordChildCount(payload), pptRecordChildTypes(payload))
					}
					if typ == 0x0fa0 || typ == 0x0fa8 {
						text := utf16BytesToStringAll(payload)
						if typ == 0x0fa8 {
							text = compressedUnicodeBytesToString(payload)
						}
						if strings.Contains(text, sample.needle) {
							t.Logf("file=%s slide=%d depth=%d path=%v text=%q", sample.file, slideIndex+1, depth, append(path, descriptor), text)
						}
					}
					if options&0x000f == 0x000f && depth < 20 {
						walk(payload, depth+1, append(append([]string(nil), path...), descriptor))
					}
				})
			}
			walk(slide, 0, nil)
		}
	}
}

func TestPPT167CapturePathDiagnostic(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "000167.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	for slideIndex, slide := range pptActiveSlideContainers(streams) {
		var walk func([]byte, int, []string)
		walk = func(d []byte, depth int, path []string) {
			pptWalkRecords(d, func(options, typ uint16, p []byte) {
				desc := fmt.Sprintf("%#x/%#x", typ, options)
				if typ == 0xf004 {
					desc += fmt.Sprintf("[fsp=%#x]", pptDiagnosticFSP(p))
				}
				if typ == 0xf003 {
					desc += fmt.Sprintf("[children=%d]", pptRecordChildCount(p))
				}
				if typ == 0x0fa0 || typ == 0x0fa8 {
					text := utf16BytesToStringAll(p)
					if typ == 0x0fa8 {
						text = compressedUnicodeBytesToString(p)
					}
					if strings.Contains(text, "Capture") || strings.TrimSpace(text) == "1" {
						t.Logf("slide=%d depth=%d path=%v text=%q", slideIndex+1, depth, append(path, desc), text)
					}
				}
				if options&0xf == 0xf && depth < 20 {
					walk(p, depth+1, append(append([]string(nil), path...), desc))
				}
			})
		}
		walk(slide, 0, nil)
	}
}

func TestPPTMasterFooterDiagnostic(t *testing.T) {
	for _, name := range []string{"000165.ppt", "013082.ppt"} {
		data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", name))
		if err != nil {
			t.Fatal(err)
		}
		streams, err := readOLEStreams(data)
		if err != nil {
			t.Fatal(err)
		}
		result := pptVisibleMasterFooterTexts(streams, 1)
		t.Logf("%s master footer=%q", name, result[0])
	}
}

func TestPPTMasterFooterHeadersDiagnostic(t *testing.T) {
	for _, name := range []string{"000165.ppt", "013082.ppt"} {
		data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", name))
		if err != nil {
			t.Fatal(err)
		}
		streams, err := readOLEStreams(data)
		if err != nil {
			t.Fatal(err)
		}
		doc, ok := findLegacyStream(streams, "PowerPoint Document")
		if !ok {
			t.Fatal("document stream missing")
		}
		for id, offset := range pptActivePersistOffsets(streams) {
			at := int(offset)
			if at+8 > len(doc.Data) || binary.LittleEndian.Uint16(doc.Data[at+2:]) != 0x03f8 {
				continue
			}
			size := int(binary.LittleEndian.Uint32(doc.Data[at+4:]))
			if size < 0 || size > len(doc.Data)-at-8 {
				continue
			}
			var headers []uint32
			var walk func([]byte, int)
			walk = func(d []byte, depth int) {
				pptWalkRecords(d, func(opts uint16, typ uint16, p []byte) {
					if typ == 0x0f9f && len(p) >= 4 {
						headers = append(headers, binary.LittleEndian.Uint32(p))
					}
					if opts&0x000f == 0x000f && depth < 20 {
						walk(p, depth+1)
					}
				})
			}
			walk(doc.Data[at+8:at+8+size], 0)
			t.Logf("%s masterPersist=%d headers=%#v", name, id, headers)
		}
	}
}

func TestPPTSlideFooterHeadersDiagnostic(t *testing.T) {
	for _, name := range []string{"000165.ppt", "013082.ppt"} {
		data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", name))
		if err != nil {
			t.Fatal(err)
		}
		streams, err := readOLEStreams(data)
		if err != nil {
			t.Fatal(err)
		}
		slides := pptActiveSlideContainers(streams)
		if len(slides) == 0 {
			t.Fatal("no slides")
		}
		var headers []uint32
		var walk func([]byte, int)
		walk = func(d []byte, depth int) {
			pptWalkRecords(d, func(opts uint16, typ uint16, p []byte) {
				if typ == 0x0f9f && len(p) >= 4 {
					headers = append(headers, binary.LittleEndian.Uint32(p))
				}
				if opts&0xf == 0xf && depth < 20 {
					walk(p, depth+1)
				}
			})
		}
		walk(slides[0], 0)
		t.Logf("%s slide1 headers=%#v", name, headers)
	}
}

func TestPPTSlidePersistAtomDiagnostic(t *testing.T) {
	for _, name := range []string{"000165.ppt", "013082.ppt"} {
		data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", name))
		if err != nil {
			t.Fatal(err)
		}
		streams, err := readOLEStreams(data)
		if err != nil {
			t.Fatal(err)
		}
		doc, _ := findLegacyStream(streams, "PowerPoint Document")
		active := pptActivePersistOffsets(streams)
		for id, off := range active {
			at := int(off)
			if at+8 > len(doc.Data) || binary.LittleEndian.Uint16(doc.Data[at+2:]) != 0x03ee {
				continue
			}
			size := int(binary.LittleEndian.Uint32(doc.Data[at+4:]))
			if size < 0 || size > len(doc.Data)-at-8 {
				continue
			}
			var atoms [][]byte
			var walk func([]byte, int)
			walk = func(d []byte, depth int) {
				pptWalkRecords(d, func(opts uint16, typ uint16, p []byte) {
					if typ == 0x03f3 {
						atoms = append(atoms, append([]byte(nil), p...))
					}
					if opts&0xf == 0xf && depth < 16 {
						walk(p, depth+1)
					}
				})
			}
			walk(doc.Data[at+8:at+8+size], 0)
			if len(atoms) > 0 {
				t.Logf("%s slidePersist=%d atom=%x", name, id, atoms[0])
			}
		}
	}
}

func pptRecordChildCount(data []byte) int {
	n := 0
	pptWalkRecords(data, func(_ uint16, _ uint16, _ []byte) { n++ })
	return n
}

func pptRecordChildTypes(data []byte) string {
	var kinds []string
	pptWalkRecords(data, func(_ uint16, typ uint16, _ []byte) { kinds = append(kinds, fmt.Sprintf("%#x", typ)) })
	return strings.Join(kinds, ",")
}

func pptDiagnosticFSPType(data []byte) uint16 {
	var shape uint16
	pptWalkRecords(data, func(options, typ uint16, payload []byte) {
		if typ == 0xf00a {
			shape = options >> 4
		}
	})
	return shape
}

func pptDiagnosticFSP(data []byte) uint32 {
	var value uint32
	pptWalkRecords(data, func(options, typ uint16, payload []byte) {
		if typ == 0xf00a && len(payload) >= 8 {
			value = uint32(options>>4)<<16 | binary.LittleEndian.Uint32(payload[4:])
		}
	})
	return value
}

func TestPPTDocumentSlidePersistDiagnostic(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	doc, ok := findLegacyStream(streams, "PowerPoint Document")
	if !ok {
		t.Fatal("missing PowerPoint Document")
	}
	active := pptActivePersistOffsets(streams)
	off := int(active[1])
	if off+8 > len(doc.Data) {
		t.Fatal("invalid document persist offset")
	}
	var walk func([]byte, int)
	walk = func(data []byte, depth int) {
		for pos := 0; pos+8 <= len(data); {
			options := binary.LittleEndian.Uint16(data[pos:])
			typ := binary.LittleEndian.Uint16(data[pos+2:])
			size := int(binary.LittleEndian.Uint32(data[pos+4:]))
			pos += 8
			if size < 0 || size > len(data)-pos {
				return
			}
			payload := data[pos : pos+size]
			if typ == 0x03f3 && len(payload) >= 4 {
				id := binary.LittleEndian.Uint32(payload)
				if target, ok := active[id]; ok {
					targetOff := int(target)
					if targetOff+8 <= len(doc.Data) {
						t.Logf("SlidePersistAtom depth=%d persistIDRef=%d -> type=%#x offset=%d", depth, id, binary.LittleEndian.Uint16(doc.Data[targetOff+2:]), target)
					}
				}
			}
			if options&0x000f == 0x000f && depth < 8 {
				walk(payload, depth+1)
			}
			pos += size
		}
	}
	size := int(binary.LittleEndian.Uint32(doc.Data[off+4:]))
	walk(doc.Data[off+8:off+8+size], 0)
}

func TestPPTCodeShapeTextDiagnostic(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "009429.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	doc, _ := findLegacyStream(streams, "PowerPoint Document")
	offset := int(pptActivePersistOffsets(streams)[10])
	size := int(binary.LittleEndian.Uint32(doc.Data[offset+4:]))
	var out []string
	pptVisibleShapeTextInto(doc.Data[offset+8:offset+8+size], 0, false, false, &out)
	t.Logf("visible=%q", out)
	for _, part := range out {
		if bytes.Contains([]byte(part), []byte("DO-NOT-DELETE")) {
			return
		}
	}
	t.Fatal("missing visible code comment")
}
