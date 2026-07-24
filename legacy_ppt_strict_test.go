package officeread

import (
	"bytes"
	"encoding/binary"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
)

func TestStrictLegacyPPTWebSampleExcludesCarvedInternalImages(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000008.ppt")
	compatible, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(compatible.Images) != 2 {
		t.Fatalf("compatibility PPT images = %d, want 2", len(compatible.Images))
	}
	strict, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(strict.Images) != 0 {
		t.Fatalf("strict PPT images = %#v, want none", strict.Images)
	}
}

func TestStrictLegacyPPTReadsVisibleWordArtTextEffectProperty(t *testing.T) {
	// The WordArt glyph string lives in the complex gtextUNICODE FOPT property
	// (0x0c0), not in a TextCharsAtom.
	property := make([]byte, 6+len(pptUTF16LEBytes("Visible WordArt")))
	binary.LittleEndian.PutUint16(property[0:], 0x80c0)
	binary.LittleEndian.PutUint32(property[2:], uint32(len(property)-6))
	copy(property[6:], pptUTF16LEBytes("Visible WordArt"))
	fopt := pptRecordWithOptionsForTest(0x0013, 0xf00b, property)
	shape := pptContainerRecord(0xf004, fopt)
	var out []string
	pptVisibleShapePropertyTextInto(shape, 0, false, &out)
	if got := strings.Join(out, "\n"); got != "Visible WordArt" {
		t.Fatalf("WordArt property text = %q, want visible glyph string", got)
	}
}

func TestStrictLegacyPPTRestoresFusedFormattingRunConnector(t *testing.T) {
	for input, want := range map[string]string{
		"Replication and Transmissionof Dengue Virus": "Replication and Transmission of Dengue Virus",
		"Dengue Virusby Aedes aegypti":                "Dengue Virus by Aedes aegypti",
		"office":                                      "office",
		"daylight":                                    "daylight",
	} {
		if got := splitPPTLegacyFusedWords(input); got != want {
			t.Fatalf("split fused connector %q = %q, want %q", input, got, want)
		}
	}
}

func TestStrictLegacyPPTRestoresSpoofFormattingRunWord(t *testing.T) {
	if got := repairPPTLegacySpoofWord("Spo of – Counterfeit GPS Signal"); got != "Spoof – Counterfeit GPS Signal" {
		t.Fatalf("spoof formatting boundary = %q", got)
	}
	if got := repairPPTLegacySpoofWord("Spo of unrelated words"); got != "Spoof unrelated words" {
		t.Fatalf("exact visible spelling was not restored: %q", got)
	}
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000167.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(result.Text, "Spo of") || !strings.Contains(result.Text, "Spoof – Counterfeit GPS Signal") {
		t.Fatalf("visible spoof text did not match PowerPoint: %q", result.Text)
	}
}

func TestStrictLegacyPPTKeepsKnownAcronymCompoundsTogether(t *testing.T) {
	for input, want := range map[string]string{
		"SS Tdrifter":   "SSTdrifter",
		"SS Tship":      "SSTship",
		"VOS Clim":      "VOSClim",
		"FUTURE Skylab": "FUTURE Skylab",
	} {
		if got := joinPPTKnownAcronymCompounds(input); got != want {
			t.Fatalf("acronym compound %q = %q, want %q", input, got, want)
		}
	}
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000165.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"SSTdrifter-SSTship", "VOSClim"} {
		if !strings.Contains(result.Text, want) {
			t.Fatalf("missing visible acronym compound %q", want)
		}
	}
}

func TestStrictLegacyPPTDropsNonVisibleDraftTimestamp(t *testing.T) {
	for input, want := range map[string]bool{
		"Draft 2-4-03 9:30AM": true,
		"Draft 2-6 8:45AM":    true,
		"Draft RFP Release":   false,
		"Final Draft":         false,
	} {
		if got := pptNonVisibleDraftTimestamp(input); got != want {
			t.Fatalf("draft timestamp %q = %v, want %v", input, got, want)
		}
	}
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000727.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, hidden := range []string{"Draft 2-4-03 9:30AM", "Draft 2-6 8:45AM"} {
		if strings.Contains(result.Text, hidden) {
			t.Fatalf("non-visible draft timestamp leaked: %q", hidden)
		}
	}
	if !strings.Contains(result.Text, "Draft RFP Release") {
		t.Fatal("visible draft prose was removed")
	}
}

func TestStrictLegacyPPTKeepsLevelAbbreviationTogether(t *testing.T) {
	if got := splitPPTAdjacentWordRuns("Final LvL 2"); got != "Final LvL 2" {
		t.Fatalf("level abbreviation split = %q", got)
	}
}

func TestStrictLegacyPPTKeepsPluralAcronymTogether(t *testing.T) {
	if got := splitPPTAdjacentWordRuns("Two RFPs"); got != "Two RFPs" {
		t.Fatalf("plural acronym split = %q", got)
	}
}

func TestStrictLegacyPPTKeepsVisibleZeroShapeLabel(t *testing.T) {
	if got := cleanPPTStrictRecordTextParts([]string{"Heading", "0", "5"}); strings.Join(got, "\n") != "Heading\n0\n5" {
		t.Fatalf("visible zero shape label was removed: %#v", got)
	}
	if got := cleanPPTStrictRecordTextParts([]string{"0", "Heading"}); strings.Join(got, "\n") != "Heading" {
		t.Fatalf("leading record-control zero was retained: %#v", got)
	}
}

func pptRecordWithOptionsForTest(options, recType uint16, payload []byte) []byte {
	record := make([]byte, 8+len(payload))
	binary.LittleEndian.PutUint16(record, options)
	binary.LittleEndian.PutUint16(record[2:], recType)
	binary.LittleEndian.PutUint32(record[4:], uint32(len(payload)))
	copy(record[8:], payload)
	return record
}

func TestStrictLegacyPPTSeparatesAdjacentFormattedWords(t *testing.T) {
	if got := splitPPTAdjacentWordRuns("RecoveryUmpqua"); got != "Recovery Umpqua" {
		t.Fatalf("formatted words = %q, want %q", got, "Recovery Umpqua")
	}
	if got := splitPPTAdjacentWordRuns("ECWR"); got != "ECWR" {
		t.Fatalf("acronym = %q, want ECWR", got)
	}
	if got := splitPPTAdjacentWordRuns("FUTURESkylab"); got != "FUTURE Skylab" {
		t.Fatalf("uppercase title boundary = %q, want %q", got, "FUTURE Skylab")
	}
	for _, name := range []string{"McMullin", "MacDonald"} {
		if got := splitPPTAdjacentWordRuns(name); got != name {
			t.Fatalf("name boundary %q = %q, want unchanged", name, got)
		}
	}
}

func TestStrictLegacyPPTRestoresMPNUnitAcrossFormattingRuns(t *testing.T) {
	if got := collapsePPTLegacyMPNUnitSpacing("MPN/100 m L"); got != "MPN/100 mL" {
		t.Fatalf("MPN unit spacing = %q, want %q", got, "MPN/100 mL")
	}
	if got := collapsePPTLegacyMPNUnitSpacing("The variables m L remain separate"); got != "The variables m L remain separate" {
		t.Fatalf("unrelated spacing changed: %q", got)
	}
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Count(result.Text, "MPN/100 mL") != 4 || strings.Contains(result.Text, "MPN/100 m L") {
		t.Fatalf("visible MPN labels did not match PowerPoint: %q", result.Text)
	}
}

func TestLegacyPPTEncryptedStreamsDoNotProduceCiphertextMojibake(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "002163.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if result.Text != "" || len(result.Images) != 0 {
		t.Fatalf("encrypted PPT produced unsupported plaintext/images: text=%q images=%d", result.Text, len(result.Images))
	}
}

func TestStrictLegacyDOCPreservesVisibleInlineImageOccurrences(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "doc", "000100.doc")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// Word's InlineShapes collection exposes sixteen inserted PNG occurrences.
	if len(result.Images) != 16 {
		t.Fatalf("strict DOC inline image occurrences = %d, want 16", len(result.Images))
	}
}

func TestStrictLegacyDOCDecompressesVisibleOfficeArtPICTBlips(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "doc", "004564.doc")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 2 {
		t.Fatalf("strict DOC visible OfficeArt PICT occurrences = %d, want 2", len(result.Images))
	}
	for i, img := range result.Images {
		if img.Ext != ".pict" || !validPICTData(img.Data) {
			t.Fatalf("image %d = %q (%d bytes), want valid PICT", i, img.Ext, len(img.Data))
		}
	}
}

func TestStrictLegacyDOCUsesRootWordDocumentInsteadOfEmbeddedObject(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "doc", "002682-2.doc")
	result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(result.Text, "NIST Handbook 133") {
		t.Fatal("strict DOC extraction did not use the root WordDocument story")
	}
	if len(result.Images) != 1 {
		t.Fatalf("strict DOC image count = %d, want 1", len(result.Images))
	}
}

func TestStrictLegacyPPTUsesVisiblePictureShapeOccurrences(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// PowerPoint's Slide.Shapes exposes 25 msoPicture instances. Several
	// instances deliberately share a source blip, so content de-duplication
	// would be incorrect here.
	if len(result.Images) != 25 {
		t.Fatalf("strict PPT visible picture occurrences = %d, want 25", len(result.Images))
	}
}

func TestStrictLegacyPPTKeepsVisibleGroupPictureFrame(t *testing.T) {
	// PowerPoint records a selected/group child as FSP flags 0xa02. It remains
	// a visible PictureFrame in GroupItems and must not be dropped.
	fsp := make([]byte, 8)
	binary.LittleEndian.PutUint32(fsp[4:], 0x0a02)
	fopt := make([]byte, 6)
	binary.LittleEndian.PutUint16(fopt, 0x0104)
	binary.LittleEndian.PutUint32(fopt[2:], 7)
	shapePayload := append(
		pptRecordWithOptionsForTest(75<<4, 0xf00a, fsp),
		pptRecordWithOptionsForTest(1<<4, 0xf00b, fopt)...,
	)
	if got, ok := pptPictureFrameBlip(shapePayload); !ok || got != 7 {
		t.Fatalf("group picture frame = (%d, %v), want (7, true)", got, ok)
	}
}

func TestStrictLegacyPPTSkipsPseudoSlidePictureFrame(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000727.ppt")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 2 {
		t.Fatalf("strict PPT visible picture occurrences = %d, want 2", len(result.Images))
	}
}

func TestStrictLegacyPPTDecompressesVisibleEMFBlips(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "002419.ppt")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// PowerPoint exposes two msoPicture shapes.  Their source blips are zlib-
	// compressed EMFs, rather than directly stored raster bytes.
	if len(result.Images) != 2 {
		t.Fatalf("strict PPT visible compressed-EMF occurrences = %d, want 2", len(result.Images))
	}
	for i, img := range result.Images {
		if img.Ext != ".emf" || !validEMFData(img.Data) {
			t.Fatalf("image %d = %q (%d bytes), want valid EMF", i, img.Ext, len(img.Data))
		}
	}
}

func TestStrictLegacyPPTDecompressesVisibleWMFBlips(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000713.ppt")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Images) != 1 {
		t.Fatalf("strict PPT visible WMF occurrences = %d, want 1", len(result.Images))
	}
	if result.Images[0].Ext != ".wmf" || !validWMFData(result.Images[0].Data) {
		t.Fatalf("image = %q (%d bytes), want valid WMF", result.Images[0].Ext, len(result.Images[0].Data))
	}
}

func TestStrictLegacyPPTUsesActiveDocumentAndSlidesOnly(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{
		"A Tiered Approach for the Identification of a Human Fecal Pollution Source",

		"Compare ENT levels at each station measured during",
	} {
		if !strings.Contains(result.Text, want) {
			t.Fatalf("missing active PPT content %q in %q", want, result.Text)
		}
	}
	for _, unwanted := range []string{"Santa Monica (I)", "Log mean ENT", "Frequency of data"} {
		if strings.Contains(result.Text, unwanted) {
			t.Fatalf("kept chart-axis text %q in %q", unwanted, result.Text)
		}
	}
}

func TestStrictLegacyPPTIncludesVisibleGroupTextWithoutChartCacheText(t *testing.T) {
	groupFile := filepath.Join("testdata", "web-samples", "samples", "ppt", "000724.ppt")
	groupResult, err := Extract(groupFile, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(groupResult.Text, "Within group differences relate to factors of:") {
		t.Fatalf("missing visible GroupItems text: %q", groupResult.Text)
	}
	chartFile := filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt")
	chartResult, err := Extract(chartFile, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, hidden := range []string{"Santa Monica (I)", "Log mean ENT", "Frequency of data"} {
		if strings.Contains(chartResult.Text, hidden) {
			t.Fatalf("chart-cache text leaked into strict result %q", hidden)
		}
	}
}

func TestStrictLegacyPPTIncludesVisibleActiveDocumentText(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000720.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(result.Text, "Hygiene Water Prefilter Umpqua, Houston, TX") {
		t.Fatalf("missing visible text box from strict PPT extraction: %q", result.Text)
	}
}

func TestStrictLegacyPPTSkipsEmbeddedEquationObjectText(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "200104.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(result.Text, `\documentclass{`) || strings.Contains(result.Text, "EXTERNALNAME") {
		t.Fatalf("embedded equation-object payload leaked into strict text: %q", result.Text)
	}
	if !strings.Contains(result.Text, "ModelEvaluator Software Summary") {
		t.Fatal("visible PowerPoint slide text was lost")
	}
}

func TestStrictLegacyPPTMaterializesMainMasterFooterAndSlideNumber(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "013082.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if got := strings.Count(result.Text, "November 9, 2006"); got < 42 {
		t.Fatalf("main-master date footer occurrences = %d, want at least 42", got)
	}
	for _, want := range []string{"TUG 2006 Page 1", "TUG 2006 Page 42"} {
		if !strings.Contains(result.Text, want) {
			t.Fatalf("missing materialized main-master footer %q", want)
		}
	}
}

func TestPPTMaterializeSlideNumberFieldRequiresBrandedFooter(t *testing.T) {
	if got := pptMaterializeSlideNumberField("Page *", 1); got != "Page *" {
		t.Fatalf("bare master page placeholder = %q, want unchanged", got)
	}
	if got := pptMaterializeSlideNumberField("TUG 2006 Page *", 42); got != "TUG 2006 Page 42" {
		t.Fatalf("branded master page placeholder = %q", got)
	}
}

func TestStrictLegacyPPTDoesNotMaterializeBareMasterPagePlaceholder(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000165.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(result.Text, "Page 1") || strings.Contains(result.Text, "Page 18") || strings.Contains(result.Text, "Page *") {
		t.Fatalf("bare master page placeholder leaked into strict result: %q", result.Text)
	}
}

func TestPPTUniqueMasterFooterTextsSkipsAlreadyMaterializedFooter(t *testing.T) {
	master := []string{"© Crown copyright 2004", "Page 1"}
	slide := []string{"© Crown copyright 2004", "Visible slide text"}
	got := pptUniqueMasterFooterTexts(master, slide)
	if len(got) != 1 || got[0] != "Page 1" {
		t.Fatalf("unique master footer text = %#v, want only inherited Page 1", got)
	}
}

func TestPPTDedupePerSlideMasterFootersKeepsOnlyOneMaterializedPage(t *testing.T) {
	got := pptDedupePerSlideMasterFooters([]string{"Page 1", "Visible slide text", "Page 1", "Repeat"})
	want := []string{"Page 1", "Visible slide text", "Repeat"}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("deduped slide footer = %#v, want %#v", got, want)
	}
}

func TestStrictLegacyPPTSkipsEmbeddedPictureAndSoundMetadata(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "009407.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, unwanted := range []string{"Microsoft Word Picture", "Word.Picture.8", "applause.wav"} {
		if strings.Contains(result.Text, unwanted) {
			t.Fatalf("embedded object metadata leaked into strict text %q", unwanted)
		}
	}
	if !strings.Contains(result.Text, "FEMA HIGHER EDUCATION CONFERENCE") {
		t.Fatal("visible slide text was lost")
	}
}

func TestLegacyPPTDropsUnrenderedOutlinePlaceholders(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "009847.ppt")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, placeholder := range []string{"Second level", "Third level", "Fourth level", "Fifth level"} {
		if strings.Contains(result.Text, placeholder) {
			t.Fatalf("kept unrendered outline placeholder %q in %q", placeholder, result.Text)
		}
	}
}

func TestLegacyPPTDropsReplacedHiddenShapeText(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "004607.ppt")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(result.Text, "and slow spill (septa)\nBeam pipe at F17 lowered") {
		t.Fatalf("kept hidden text runs in %q", result.Text)
	}
	if !strings.Contains(result.Text, "and slow spill (septa) to SY absorber") || !strings.Contains(result.Text, "Beam pipe at F17 c-magnet raised") {
		t.Fatalf("lost rendered replacement text in %q", result.Text)
	}
}

func TestStrictLegacyPPTDoesNotMergeSmallDocumentContainerMetadata(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000712.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, unwanted := range []string{"Camera", "WAVE", "\n105\n"} {
		if strings.Contains(result.Text, unwanted) {
			t.Fatalf("kept non-shape document metadata %q in %q", unwanted, result.Text)
		}
	}
	if !strings.Contains(result.Text, "Milestones for Manned Exploration of Mars") {
		t.Fatalf("lost visible slide text in %q", result.Text)
	}
}

func TestLegacyPPTTextKeepsVisibleSlideDuplicatesAndExcludesNotes(t *testing.T) {
	var slide bytes.Buffer
	slide.Write(pptRecord(0x0fa8, []byte("Repeated visible title")))
	var deck bytes.Buffer
	deck.Write(pptContainerRecord(0x03ee, slide.Bytes()))
	deck.Write(pptContainerRecord(0x03ee, slide.Bytes()))
	deck.Write(pptContainerRecord(0x03f0, pptRecord(0x0fa8, []byte("Speaker note only"))))
	text := strings.Join(pptRecordText(deck.Bytes()), "\n")
	if strings.Count(text, "Repeated visible title") != 2 {
		t.Fatalf("visible slide text occurrences = %q", text)
	}
	if strings.Contains(text, "Speaker note only") {
		t.Fatalf("notes text leaked into slide text: %q", text)
	}
}
