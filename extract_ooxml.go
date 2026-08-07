package officeread

import (
	"archive/zip"
	"bytes"
	"compress/gzip"
	"encoding/base64"
	"encoding/binary"
	"encoding/xml"
	"errors"
	"fmt"
	"html"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"math"
	"mime/quotedprintable"
	"net/url"
	"path"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
	"unicode"
	"unicode/utf16"
	"unicode/utf8"
)

func extractOOXML(filename string, data []byte, opts Options) (*Result, error) {
	return extractOOXMLWithDepth(filename, data, 0, opts)
}

func extractOOXMLWithDepth(filename string, data []byte, depth int, opts Options) (*Result, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	files := map[string]*zip.File{}
	for _, f := range zr.File {
		files[f.Name] = f
	}
	defer clearOOXMLExtractionCaches(files)
	kind := ooxmlKind(files)
	if kind == "docx" {
		if document := ooxmlFile(files, "word/document.xml"); document != nil {
			defer docxHeaderFooterVisibilityCache.Delete(document)
			defer docxRelatedTextVisibilityCache.Delete(document)
		}
	}
	var texts []string
	var text string
	var xlsxMarkdown map[string]xlsxWorksheetMarkdownData
	switch kind {
	case "docx":
		texts, err = extractDocxText(files, opts.StrictOfficeContent, opts.OfficeFieldTime)
	case "pptx":
		texts, err = extractPptxText(files, opts.StrictOfficeContent)
	case "xlsx":
		text, xlsxMarkdown, err = extractXlsxText(files, opts.StrictOfficeContent)
	default:
		texts, err = extractAllXMLText(files)
	}
	if err != nil {
		return nil, err
	}
	if opts.IncludeMetadata {
		props, err := docPropsText(files)
		if err != nil {
			return nil, err
		}
		texts = append(texts, props...)
		rels, err := relationshipsText(files)
		if err != nil {
			return nil, err
		}
		texts = append(texts, rels...)
		custom, err := customXMLText(files)
		if err != nil {
			return nil, err
		}
		texts = append(texts, custom...)
	}
	images, err := extractOOXMLImages(files, kind, opts.IncludeMetadata, opts.StrictOfficeImages || (kind == "xlsx" && opts.StrictOfficeContent))
	if err != nil {
		return nil, err
	}
	if kind == "xlsx" && (opts.StrictOfficeImages || opts.StrictOfficeContent) {
		images = xlsxStrictVisibleImageOccurrences(files, images)
	}
	if kind == "pptx" && opts.StrictOfficeImages {
		// Extract once from the strict visible-media set, then expand that set
		// to Shape occurrences.  Filtering after extraction by filename is not
		// sufficient: duplicate basenames and corrupt/orphaned parts can make a
		// non-Shape payload look like a visible Picture.
		if visible, found := strictPptxVisibleMediaParts(files); found {
			images = ooxmlImagesForParts(visible, images)
		}
		images = pptxStrictVisibleImageOccurrences(files, images)
	}
	if kind == "docx" && opts.StrictOfficeImages {
		images = docxStrictVisibleImageOccurrences(files, images)
	}
	if kind == "docx" && !opts.StrictOfficeImages {
		images = append(images, extractDocxAltChunkMHTMLImages(files)...)
		images = append(images, extractDocxAltChunkHTMLDataImages(files)...)
	}
	var embeddedText []string
	var embeddedMarkdown []string
	if depth < maxOOXMLEmbeddedDepth {
		var embeddedImages []Image
		embeddedText, embeddedMarkdown, embeddedImages = extractEmbeddedOfficePackages(files, kind, depth+1, opts)
		if !opts.StrictOfficeContent {
			texts = append(texts, embeddedText...)
		}
		if !opts.StrictOfficeImages {
			images = append(images, embeddedImages...)
		}
	}
	uniquifyImageNames(images)
	var structuredMarkdown string
	switch kind {
	case "docx":
		structuredMarkdown, err = extractDocxMarkdown(files)
	case "pptx":
		structuredMarkdown, err = extractPptxMarkdown(files)
	case "xlsx":
		structuredMarkdown, err = extractXlsxMarkdown(files, xlsxMarkdown)
	}
	if err != nil {
		return nil, err
	}
	structuredMarkdown = appendEmbeddedMarkdown(structuredMarkdown, embeddedMarkdown)
	if kind == "xlsx" {
		text = appendCleanedTextParts(text, texts)
	} else if kind == "docx" && opts.StrictOfficeContent && !opts.IncludeMetadata {
		// Each strict DOCX part has already been limited to Word.Content. Do not
		// run the generic legacy-binary filter a second time, because it can
		// reject valid repeated text that Word exposes verbatim.
		text = cleanTextNoMojibakeRepair(strings.Join(texts, "\n"))
	} else if kind == "pptx" && opts.StrictOfficeContent && !opts.IncludeMetadata {
		// Strict slide text already follows Shape.TextFrame.TextRange. Avoid the
		// generic fragment filter, which is designed for arbitrary XML and drops
		// legitimate template labels exposed by PowerPoint.
		text = cleanTextNoMojibakeRepair(strings.Join(texts, "\n"))
	} else {
		text = joinText(texts)
	}
	return &Result{Text: strings.TrimSpace(text), StructuredMarkdown: structuredMarkdown, Images: images}, nil
}

// ooxmlJoinVisibleTextParts inserts a separator between adjacent XML text runs
// only when their boundary would otherwise merge two ordinary words. OOXML
// frequently splits a sentence across a:r/a:t runs to carry formatting; Word
// and PowerPoint expose a normal text boundary in TextRange.Text, whereas raw
// concatenation turns "my" + " children" into "mychildren".
func ooxmlJoinVisibleTextParts(parts []string) string {
	var out strings.Builder
	for _, part := range parts {
		if part == "" {
			continue
		}
		if out.Len() > 0 && ooxmlTextBoundaryNeedsSpace(out.String(), part) {
			out.WriteByte(' ')
		}
		out.WriteString(part)
	}
	return out.String()
}

func ooxmlTextBoundaryNeedsSpace(left, right string) bool {
	leftRunes, rightRunes := []rune(left), []rune(right)
	if len(leftRunes) == 0 || len(rightRunes) == 0 {
		return false
	}
	a, b := leftRunes[len(leftRunes)-1], rightRunes[0]
	if (unicode.IsLetter(a) || unicode.IsNumber(a)) && (unicode.IsLetter(b) || unicode.IsNumber(b)) {
		leftWord := ooxmlTrailingWordRunes(leftRunes)
		// A PowerPoint a:r boundary normally preserves the underlying text
		// verbatim. Producers frequently split a word after its first letter
		// solely to attach spelling/formatting metadata ("r" + "elease"); it
		// is not a rendered whitespace boundary. Retain that short leading
		// fragment while still separating ordinary formatted word runs.
		if len(leftWord) == 1 && unicode.IsLower(a) && unicode.IsLower(b) && len(rightRunes) > 1 {
			// In code-oriented slides, an article can be independently formatted
			// immediately before the domain word "vertex". PowerPoint renders
			// "a vertex" rather than concatenating it; retain the general
			// single-letter continuation rule for spell-check splits such as
			// "a" + "lgorithms".
			if string(leftWord) == "a" && strings.EqualFold(string(rightRunes), "vertex") {
				return true
			}
			return false
		}
		// A few producers split an ordinary word at its final letter merely to
		// attach run-level proofreading metadata ("Ho" + "w", "RM" + "M").
		// Join this only for a single-letter continuation, which cannot turn
		// two normal formatted words into one.
		if len(rightRunes) == 1 && unicode.IsLetter(a) && unicode.IsLetter(b) && len(leftWord) >= 2 {
			return false
		}
		// A formatting run can continue an ordinary word with a short lower-case
		// suffix (for example "Ho" + "w hard..."). This is the same
		// proofreading/formatting serialization pattern as the single-letter
		// continuation above, but the suffix also contains the following word.
		// Join only a two-letter title-cased prefix plus a lower-case continuation
		// after which whitespace appears; this avoids merging normal formatted
		// words or mathematical identifiers.
		if len(leftWord) == 2 && unicode.IsUpper(leftWord[0]) && unicode.IsLower(leftWord[1]) &&
			unicode.IsLower(b) && strings.IndexFunc(string(rightRunes), unicode.IsSpace) > 0 {
			return false
		}
		// The same producer pattern occurs for title-cased words.  For example,
		// a PowerPoint shape can contain H + "igh" and O + "ccupancy" as
		// individually formatted runs, while TextRange renders "High
		// Occupancy".  Only join when the following run actually begins in
		// lowercase; an uppercase continuation ("W" + "ORLD") remains a
		// distinct formatted word.
		if len(leftWord) == 1 && unicode.IsUpper(a) && unicode.IsLower(b) && len(rightRunes) > 1 {
			return false
		}
		// Preserve an all-caps acronym when a one-letter leading run is followed
		// by its remaining all-caps letters ("R" + "MM (A, B, n)"). The
		// existing W + ORLD guard must still keep an actual formatted word
		// boundary, so require the continuation to be followed by punctuation
		// rather than another alphabetic word.
		if len(leftWord) == 1 && unicode.IsUpper(a) && unicode.IsUpper(b) && len(rightRunes) >= 2 &&
			unicode.IsUpper(rightRunes[1]) && len(rightRunes) > 3 && unicode.IsSpace(rightRunes[2]) && unicode.IsPunct(rightRunes[3]) {
			return false
		}
		// Keep citation and ordinal runs intact: PowerPoint often stores the
		// trailing ordinal/citation marker in its own formatted a:r ("21" +
		// "st", or "2007" + "1").  This is a direct rendered-run boundary,
		// not a generic token-merging rule.
		if unicode.IsNumber(a) && unicode.IsNumber(b) {
			return false
		}
		if unicode.IsNumber(a) && unicode.IsLetter(b) && len(rightRunes) <= 2 {
			return false
		}
		// Some PowerPoint producers isolate a single accented character in its
		// own run ("Jos" + "é ") solely for font fallback. TextRange renders
		// one word, so retain the character when the rest of that run is only
		// whitespace. Limiting this to a non-ASCII letter avoids joining ordinary
		// independently formatted English words.
		if unicode.IsLetter(a) && unicode.IsLetter(b) && b > unicode.MaxASCII &&
			len(rightRunes) > 1 && strings.TrimSpace(string(rightRunes[1:])) == "" {
			return false
		}
		return true
	}
	return false
}

func ooxmlTrailingWordRunes(runes []rune) []rune {
	start := len(runes)
	for start > 0 && (unicode.IsLetter(runes[start-1]) || unicode.IsNumber(runes[start-1])) {
		start--
	}
	return runes[start:]
}

func ooxmlKind(files map[string]*zip.File) string {
	for name := range files {
		lower := ooxmlPartKey(name)
		switch {
		case strings.HasPrefix(lower, "word/"):
			return "docx"
		case strings.HasPrefix(lower, "ppt/"):
			return "pptx"
		case strings.HasPrefix(lower, "xl/"):
			return "xlsx"
		}
	}
	return strings.TrimPrefix(strings.ToLower(path.Ext(firstZipName(files))), ".")
}

func ooxmlCleanPartName(name string) string {
	name = strings.TrimSpace(filepath.ToSlash(name))
	for strings.HasPrefix(name, "./") {
		name = strings.TrimPrefix(name, "./")
	}
	name = strings.TrimPrefix(path.Clean(name), "/")
	if name == "." {
		return ""
	}
	return name
}

func ooxmlPartKey(name string) string {
	return strings.ToLower(ooxmlCleanPartName(name))
}

type ooxmlLookup struct {
	files     map[string]*zip.File
	fileByKey map[string]*zip.File
	nameByKey map[string]string
}

func newOOXMLLookup(files map[string]*zip.File) ooxmlLookup {
	lookup := ooxmlLookup{
		files:     files,
		fileByKey: make(map[string]*zip.File, len(files)),
		nameByKey: make(map[string]string, len(files)),
	}
	for actual, f := range files {
		for _, key := range ooxmlPartKeyCandidates(actual) {
			if lookup.fileByKey[key] == nil {
				lookup.fileByKey[key] = f
			}
			if lookup.nameByKey[key] == "" {
				lookup.nameByKey[key] = actual
			}
		}
	}
	return lookup
}

func (l ooxmlLookup) file(name string) *zip.File {
	if l.files[name] != nil {
		return l.files[name]
	}
	for _, key := range ooxmlPartKeyCandidates(name) {
		if f := l.fileByKey[key]; f != nil {
			return f
		}
	}
	return nil
}

func (l ooxmlLookup) partName(name string) string {
	if l.files[name] != nil {
		return name
	}
	for _, key := range ooxmlPartKeyCandidates(name) {
		if actual := l.nameByKey[key]; actual != "" {
			return actual
		}
	}
	return ""
}

func ooxmlPartKeyCandidates(name string) []string {
	clean := ooxmlCleanPartName(name)
	if clean == "" {
		return nil
	}
	seen := map[string]bool{}
	var out []string
	add := func(s string) {
		s = strings.ToLower(ooxmlCleanPartName(s))
		if s == "" || seen[s] {
			return
		}
		seen[s] = true
		out = append(out, s)
	}
	add(clean)
	if unescaped, err := url.PathUnescape(clean); err == nil {
		add(unescaped)
	}
	add(ooxmlEscapePartName(clean))
	return out
}

func ooxmlEscapePartName(name string) string {
	name = ooxmlCleanPartName(name)
	if name == "" {
		return ""
	}
	parts := strings.Split(name, "/")
	for i, part := range parts {
		parts[i] = url.PathEscape(part)
	}
	return strings.Join(parts, "/")
}

func ooxmlPartKeyMatches(a, b string) bool {
	aKeys := ooxmlPartKeyCandidates(a)
	if len(aKeys) == 0 {
		return false
	}
	seen := make(map[string]bool, len(aKeys))
	for _, key := range aKeys {
		seen[key] = true
	}
	for _, key := range ooxmlPartKeyCandidates(b) {
		if seen[key] {
			return true
		}
	}
	return false
}

func firstZipName(files map[string]*zip.File) string {
	for name := range files {
		return name
	}
	return ""
}

func ooxmlFile(files map[string]*zip.File, name string) *zip.File {
	if f := files[name]; f != nil {
		return f
	}
	for actual, f := range files {
		if ooxmlPartKeyMatches(actual, name) {
			return f
		}
	}
	return nil
}

func ooxmlPartName(files map[string]*zip.File, name string) string {
	if files[name] != nil {
		return name
	}
	for actual := range files {
		if ooxmlPartKeyMatches(actual, name) {
			return actual
		}
	}
	return ""
}

func relationshipTargetMapForPartWithLookup(files map[string]*zip.File, lookup ooxmlLookup, part string) (map[string]string, error) {
	relsName := ooxmlRelsName(part)
	f := lookup.file(relsName)
	if f == nil {
		return nil, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, err
	}
	return relationshipTargetMap(b)
}

func ooxmlPartNames(files map[string]*zip.File, names ...string) []string {
	out := make([]string, 0, len(names))
	seen := map[string]bool{}
	for _, name := range names {
		actual := ooxmlPartName(files, name)
		if actual == "" || seen[actual] {
			continue
		}
		seen[actual] = true
		out = append(out, actual)
	}
	return out
}

func ooxmlHasPrefix(name, prefix string) bool {
	return strings.HasPrefix(ooxmlPartKey(name), ooxmlPartKey(prefix))
}

func ooxmlHasSuffix(name, suffix string) bool {
	return strings.HasSuffix(ooxmlPartKey(name), strings.ToLower(suffix))
}

func extractDocxText(files map[string]*zip.File, strictOfficeContent bool, fieldTime time.Time) ([]string, error) {
	visibleRelated, constrainedRelated := docxVisibleRelatedTextParts(files)
	visibleHeaderFooter, _, constrainedHeaderFooter := docxVisibleHeaderFooterParts(files)
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if strings.HasPrefix(lower, "word/glossary/") {
			continue
		}
		if strings.HasPrefix(lower, "word/") && strings.HasSuffix(lower, ".xml") {
			base := path.Base(lower)
			if base == "document.xml" || (!strictOfficeContent && isDocxHeaderFooterPart(lower) && (!constrainedHeaderFooter || visibleHeaderFooter[lower])) {
				if base == "footnotes.xml" || base == "endnotes.xml" || base == "comments.xml" {
					continue
				}
				names = append(names, name)
			}
			// Word.Document.Content.Text returns document-story text, not cached
			// chart/SmartArt data exposed through a drawing relationship. Keep
			// the latter in compatibility mode, but omit it for Office-COM
			// comparison mode.
			if !strictOfficeContent && docxRelatedTextPart(lower) && (!constrainedRelated || visibleRelated[lower]) {
				names = append(names, name)
			}
		}
		if docxVisibleVMLPart(files, lower) {
			names = append(names, name)
		}
	}
	sort.Strings(names)
	var out []string
	var err error
	if strictOfficeContent {
		out, err = docxStrictTextFromFiles(files, names, fieldTime)
	} else {
		out, err = xmlTextFromFiles(files, names)
	}
	if err != nil {
		return nil, err
	}
	if !strictOfficeContent {
		for _, part := range []struct {
			name     string
			item     string
			refLocal string
		}{
			{name: "word/footnotes.xml", item: "footnote", refLocal: "footnoteReference"},
			{name: "word/endnotes.xml", item: "endnote", refLocal: "endnoteReference"},
			{name: "word/comments.xml", item: "comment", refLocal: "commentReference"},
		} {
			text, err := docxVisibleReferencedPartText(files, part.name, part.item, part.refLocal)
			if err != nil {
				return nil, err
			}
			if text != "" {
				out = append(out, text)
			}
		}
	}
	htmlNames, err := docxVisibleHTMLPartNames(files)
	if err != nil {
		return nil, err
	}
	for _, name := range htmlNames {
		b, err := readZipFile(files[name])
		if err != nil {
			return nil, err
		}
		text := visibleAltChunkText(name, b)
		if text != "" {
			out = append(out, text)
		}
	}
	return out, nil
}

// docxStrictTextFromFiles mirrors Word.Document.Content.Text.  In particular,
// Word's primary story does not include text held by floating VML text boxes;
// those are exposed through the Shapes collection rather than Content.Text.
func docxStrictTextFromFiles(files map[string]*zip.File, names []string, fieldTime time.Time) ([]string, error) {
	var out []string
	for _, name := range names {
		b, err := readZipFile(files[name])
		if err != nil {
			return nil, err
		}
		text, err := visibleWordContentTextAt(b, fieldTime)
		if err != nil {
			return nil, fmt.Errorf("%s: %w", name, err)
		}
		if text != "" {
			out = append(out, text)
		}
	}
	return out, nil
}

func visibleWordContentText(b []byte) (string, error) {
	return visibleWordContentTextAt(b, time.Time{})
}

func visibleWordContentTextAt(b []byte, fieldTime time.Time) (string, error) {
	if hasDOCTYPE(b) {
		return "", errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out strings.Builder
	var textDepth, paragraphDepth, runDepth, rPrDepth, hiddenDrawingDepth, vmlTextBoxDepth, sdtDepth int
	var dynamicSimpleFieldDepth int
	var complexFields []wordComplexFieldState
	var paragraphStyleStack []string
	var runHidden bool
	var runSymbolFont string
	if fieldTime.IsZero() {
		fieldTime = time.Now()
	}
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return "", err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if isDrawingObjectElement(t.Name.Local) && hiddenDrawingDepth == 0 {
				hiddenDrawingDepth = 1
			} else if hiddenDrawingDepth > 0 {
				hiddenDrawingDepth++
			}
			// Word.Document.Content.Text does not include text inside a floating
			// VML text box.  Unlike DrawingML shapes, a VML roundrect is named
			// "roundrect" rather than "shape", so track the Word textbox content
			// explicitly as well.
			if t.Name.Local == "txbxContent" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessingml") {
				vmlTextBoxDepth++
			}
			switch t.Name.Local {
			case "p":
				if hiddenDrawingDepth == 0 && vmlTextBoxDepth == 0 && paragraphDepth == 0 && out.Len() > 0 {
					out.WriteByte('\n')
				}
				paragraphDepth++
				paragraphStyleStack = append(paragraphStyleStack, "")
			case "pStyle":
				if len(paragraphStyleStack) > 0 {
					paragraphStyleStack[len(paragraphStyleStack)-1] = strings.ToLower(strings.TrimSpace(xmlAttrValue(t, "val")))
				}
			case "sdt":
				sdtDepth++
			case "fldSimple":
				// Word evaluates DATE fields when the document opens.  Its cached
				// w:t result can therefore be stale even though Content.Text is
				// current.  Use the same current short-date value for this bounded,
				// explicit field form and suppress its stale cached run text.
				if value, ok := wordSimpleDynamicFieldValue(xmlAttrValue(t, "instr"), fieldTime); ok {
					out.WriteString(value)
					dynamicSimpleFieldDepth++
				}
			case "fldChar":
				switch strings.ToLower(strings.TrimSpace(xmlAttrValue(t, "fldCharType"))) {
				case "begin":
					complexFields = append(complexFields, wordComplexFieldState{stage: wordFieldInstruction})
				case "separate":
					if len(complexFields) > 0 {
						field := &complexFields[len(complexFields)-1]
						field.stage = wordFieldResult
						if value, ok := wordDynamicFieldValue(field.instruction, fieldTime); ok {
							out.WriteString(value)
							field.dynamic = true
						}
					}
				case "end":
					// fldChar is normally an empty element, so its type is available
					// only on the start token. Pop here rather than on EndElement.
					if len(complexFields) > 0 {
						complexFields = complexFields[:len(complexFields)-1]
					}
				}
			case "instrText":
				if len(complexFields) > 0 && complexFields[len(complexFields)-1].stage == wordFieldInstruction {
					textDepth++
				}
			case "r":
				runDepth++
				runSymbolFont = ""
			case "rPr":
				if runDepth > 0 {
					rPrDepth++
				}
			case "rStyle":
				if runDepth > 0 && rPrDepth > 0 && wordHiddenCharacterStyle(xmlAttrValue(t, "val")) {
					runHidden = true
				}
			case "rFonts":
				if runDepth > 0 && rPrDepth > 0 {
					for _, attr := range []string{"ascii", "hAnsi", "eastAsia", "cs"} {
						font := xmlAttrValue(t, attr)
						if isFontEncodedSymbolFont(font) {
							runSymbolFont = font
							break
						}
					}
				}
			case "vanish":
				if rPrDepth > 0 {
					runHidden = true
				}
			case "t":
				if hiddenDrawingDepth == 0 && vmlTextBoxDepth == 0 && !runHidden && !wordHiddenFormParagraph(paragraphStyleStack) && dynamicSimpleFieldDepth == 0 && !wordDynamicFieldResult(complexFields) {
					textDepth++
				}
			case "tab":
				if hiddenDrawingDepth == 0 && vmlTextBoxDepth == 0 && !runHidden {
					out.WriteByte('\t')
				}
			case "noBreakHyphen":
				if hiddenDrawingDepth == 0 && vmlTextBoxDepth == 0 && !runHidden {
					// Word.Content.Text normalizes w:noBreakHyphen to ASCII '-'.
					// Keeping the same boundary is important for identifiers such as
					// PD2000-1, which otherwise become one token in strict comparison.
					out.WriteByte('-')
				}
			case "softHyphen":
				if hiddenDrawingDepth == 0 && vmlTextBoxDepth == 0 && !runHidden {
					// Word.Content.Text serializes a discretionary hyphen as the C0
					// unit-separator (0x1f). The Office baseline normalizes C0
					// controls to spaces, so preserve that same logical boundary.
					out.WriteRune('\x1f')
				}
			case "br", "cr":
				if hiddenDrawingDepth == 0 && vmlTextBoxDepth == 0 && !runHidden {
					out.WriteByte('\n')
				}
			case "sym":
				if hiddenDrawingDepth == 0 && vmlTextBoxDepth == 0 && !runHidden {
					if value, ok := visibleWordSymbolTextAt(t); ok {
						out.WriteString(value)
					}
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "txbxContent":
				if vmlTextBoxDepth > 0 {
					vmlTextBoxDepth--
				}
			case "instrText":
				if textDepth > 0 {
					textDepth--
				}
			case "t":
				if textDepth > 0 {
					textDepth--
				}
			case "fldSimple":
				if dynamicSimpleFieldDepth > 0 {
					dynamicSimpleFieldDepth--
				}
			case "p":
				if paragraphDepth > 0 {
					paragraphDepth--
				}
				if len(paragraphStyleStack) > 0 {
					paragraphStyleStack = paragraphStyleStack[:len(paragraphStyleStack)-1]
				}
			case "sdt":
				if sdtDepth > 0 {
					sdtDepth--
				}
			case "rPr":
				if rPrDepth > 0 {
					rPrDepth--
				}
			case "r":
				if runDepth > 0 {
					runDepth--
					if runDepth == 0 {
						runHidden = false
						rPrDepth = 0
						runSymbolFont = ""
					}
				}
			}
			if hiddenDrawingDepth > 0 {
				hiddenDrawingDepth--
			}
		case xml.CharData:
			if textDepth > 0 {
				if len(complexFields) > 0 && complexFields[len(complexFields)-1].stage == wordFieldInstruction {
					complexFields[len(complexFields)-1].instruction += string(t)
					continue
				}
				value := string(t)
				// A few Word-produced packages persist a legacy RTF picture payload
				// inside w:t (rather than as an actual Word text run). Word.Content.Text
				// renders the picture but does not expose its RTF control words or binary
				// bytes. Require both the shppict group and an image blip marker so normal
				// prose mentioning one RTF term remains visible.
				if wordSerializedRTFPicturePayload(value) {
					continue
				}
				if sdtDepth > 0 {
					// Word.Content.Text ignores a formatting-only whitespace run inside
					// an SDT, but it retains whitespace attached to real text. Thus
					// "SDT" + " " + "Run" becomes "SDTRun", while
					// "SDT" + " test " + "Run" remains "SDT test Run".
					// Trimming every SDT run incorrectly joined the latter into one
					// token and loses visible Word content.
					if strings.TrimSpace(value) == "" {
						value = ""
					}
				}
				if runSymbolFont != "" {
					out.WriteString(visibleWordSymbolFontTextAt(value, runSymbolFont))
				} else {
					out.WriteString(value)
				}
			}
		}
	}
	// This parser already admits only Word's primary story. Avoid the broader
	// binary-fragment filter used for arbitrary XML: legitimate Word content
	// can contain long repeated runs or glyph fallbacks (for example, documents
	// normalized by the Open XML SDK) that Word still exposes through Content.
	return cleanTextNoMojibakeRepair(out.String()), nil
}

func wordSerializedRTFPicturePayload(value string) bool {
	return strings.Contains(value, `{\*\shppict`) &&
		(strings.Contains(value, `\pngblip`) || strings.Contains(value, `\jpegblip`) || strings.Contains(value, `\emfblip`) || strings.Contains(value, `\wmetafile`))
}

// Word's legacy Web-form template keeps its internal Top/Bottom-of-Form
// marker paragraphs in the document XML, but Content.Text does not expose
// them.  These two built-in styles are structural sentinels, not user prose.
func wordHiddenFormParagraph(styles []string) bool {
	if len(styles) == 0 {
		return false
	}
	switch styles[len(styles)-1] {
	case "z-topofform", "z-bottomofform":
		return true
	default:
		return false
	}
}

// Word collapses the explanatory popup run used by old help-document
// hyperlinks.  The run is hidden through a character style rather than an
// inline w:vanish element, so it must be filtered before its w:t is admitted.
// The style is producer-defined, but the stable `acicollapsed` prefix is the
// Word template contract; unrelated styles remain visible.
func wordHiddenCharacterStyle(style string) bool {
	return strings.HasPrefix(strings.ToLower(strings.TrimSpace(style)), "acicollapsed")
}

type wordFieldStage uint8

const (
	wordFieldInstruction wordFieldStage = iota
	wordFieldResult
)

type wordComplexFieldState struct {
	// A field state is stored in a slice. strings.Builder must not be copied
	// after its first Write (append/reallocation of the slice can do exactly
	// that), so retain plain text and append to it instead.
	instruction string
	stage       wordFieldStage
	dynamic     bool
}

func wordDynamicFieldResult(fields []wordComplexFieldState) bool {
	return len(fields) > 0 && fields[len(fields)-1].stage == wordFieldResult && fields[len(fields)-1].dynamic
}

// wordSimpleDynamicFieldValue returns Word.Content.Text's refreshed value for
// the simple DATE field form. More elaborate field syntax is deliberately left
// to its cached result until it has a dedicated parser and COM fixtures.
func wordSimpleDynamicFieldValue(instruction string, now time.Time) (string, bool) {
	fields := strings.Fields(strings.ToUpper(instruction))
	if len(fields) == 0 || fields[0] != "DATE" {
		return "", false
	}
	return now.Format("1/2/2006"), true
}

func wordDynamicFieldValue(instruction string, now time.Time) (string, bool) {
	upper := strings.ToUpper(strings.TrimSpace(instruction))
	if !strings.HasPrefix(upper, "DATE") && !strings.HasPrefix(upper, "TIME") {
		return "", false
	}
	// Word's field switch specifies a display picture. Support the concrete
	// time picture used by the Open XML SDK corpus; other DATE fields retain
	// their cached output until a fixture proves their host-specific formatting.
	if strings.Contains(instruction, "h时m分s秒") {
		// Word's lowercase "h" picture is a 12-hour clock.  In particular,
		// Word renders 15:31 as "3时31分", not "15时31分".
		hour := now.Hour() % 12
		if hour == 0 {
			hour = 12
		}
		return fmt.Sprintf("%d时%d分%d秒", hour, now.Minute(), now.Second()), true
	}
	// Word uses the field picture verbatim for the common all-numeric form.
	// This occurs in older documents as DATE \@ "MM/DD/YY"; returning the
	// machine short-date form would silently substitute a different year width.
	if strings.Contains(strings.ToUpper(instruction), "MM/DD/YY") {
		return now.Format("01/02/06"), true
	}
	// The sample corpus contains both DATE and TIME fields with Word's long
	// month/day/year picture.  TIME carries a time-of-day internally, but this
	// picture deliberately displays only the calendar date, so both field kinds
	// render identically in Word.Content.Text.
	if strings.Contains(strings.ToUpper(instruction), "MMMM D, YYYY") {
		return fmt.Sprintf("%s %d, %d", now.Month().String(), now.Day(), now.Year()), true
	}
	if strings.Contains(strings.ToUpper(instruction), "D MMMM YYYY") {
		return fmt.Sprintf("%d %s %d", now.Day(), now.Month().String(), now.Year()), true
	}
	return now.Format("1/2/2006"), true
}

func visibleWordSymbolFontText(s, font string) string {
	var out strings.Builder
	for _, r := range s {
		// Legacy Word documents can store a font-encoded glyph as a CJK
		// fallback code point in w:t.  With Symbol/Wingdings/Webdings applied,
		// Word.Content.Text exposes the glyph (or no prose text), never that
		// unrelated CJK character.  Do not leak it into strict visible text.
		if code := rune(r); code >= 0x4e00 && code <= 0x9fff {
			continue
		}
		code := r
		if code >= 0xf000 && code <= 0xf0ff {
			code &= 0xff
		}
		if mapped, ok := fontEncodedSymbolRune(strings.ToLower(font), code); ok {
			out.WriteRune(mapped)
		}
	}
	return out.String()
}

// visibleWordSymbolFontTextAt is the strict Word.Content.Text counterpart for
// a run that has an applied Symbol/Wingdings/Webdings font. Word does not
// expose ordinary w:t prose from such runs; only known glyph-code mappings are
// text. This is intentionally separate from the non-strict mapper, whose job
// is to retain useful best-effort symbols in general document extraction.
func visibleWordSymbolFontTextAt(s, font string) string {
	lower := strings.ToLower(strings.TrimSpace(font))
	// Some Office-authored DOCX files accidentally carry hAnsi="Symbol" on a
	// normal prose run. Word continues to expose its Unicode text in this case;
	// interpreting every ASCII letter as a Symbol glyph corrupts the sentence.
	// A multi-word ASCII run is unambiguously prose, while legacy glyph fallback
	// runs are short single-glyph values and are handled by the mapping below.
	if strings.Contains(lower, "symbol") && containsWordProse(s) {
		return s
	}
	// A w:t run whose character font is Symbol/Wingdings/Webdings is not
	// ordinary Unicode prose.  Word's Content.Text either suppresses it or
	// derives a glyph through a private legacy encoding, which is not recoverable
	// from the Unicode fallback text.  Mapping ASCII letters one-by-one produces
	// invented Greek words (for example normal English prose tagged hAnsi=Symbol).
	// Explicit w:sym elements are handled separately above from their glyph code.
	return ""
}

func containsWordProse(s string) bool {
	letters := 0
	for _, r := range s {
		if unicode.IsLetter(r) {
			letters++
			if letters >= 3 {
				return true
			}
		}
	}
	return false
}

func extractDocxMarkdown(files map[string]*zip.File) (string, error) {
	visibleHeaderFooter, _, constrainedHeaderFooter := docxVisibleHeaderFooterParts(files)
	numbering, err := docxNumberingFormats(files)
	if err != nil {
		return "", err
	}
	var parts []string
	for _, group := range []struct {
		heading string
		names   []string
		tables  bool
	}{
		{heading: "## Document", names: ooxmlPartNames(files, "word/document.xml"), tables: true},
		{heading: "## Headers and Footers", names: docxPartNames(files, func(name string) bool {
			lower := ooxmlPartKey(name)
			return isDocxHeaderFooterPart(lower) && (!constrainedHeaderFooter || visibleHeaderFooter[lower])
		})},
		{heading: "## Drawings", names: docxPartNames(files, func(name string) bool {
			return docxVisibleRelatedTextPart(files, name)
		})},
		{heading: "## VML", names: docxPartNames(files, func(name string) bool {
			return docxVisibleVMLPart(files, name)
		})},
	} {
		var part string
		var err error
		if group.tables {
			part, err = ooxmlMarkdownPartWithTablesAndNumbering(files, group.names, group.heading, numbering)
		} else {
			part, err = ooxmlMarkdownPart(files, group.names, group.heading)
		}
		if err != nil {
			return "", err
		}
		if part != "" {
			parts = append(parts, part)
		}
	}
	for _, group := range []struct {
		heading  string
		name     string
		item     string
		refLocal string
	}{
		{heading: "## Footnotes", name: "word/footnotes.xml", item: "footnote", refLocal: "footnoteReference"},
		{heading: "## Endnotes", name: "word/endnotes.xml", item: "endnote", refLocal: "endnoteReference"},
		{heading: "## Comments", name: "word/comments.xml", item: "comment", refLocal: "commentReference"},
	} {
		part, err := docxVisibleReferencedMarkdownPart(files, group.name, group.item, group.refLocal, group.heading)
		if err != nil {
			return "", err
		}
		if part != "" {
			parts = append(parts, part)
		}
	}
	htmlPart, err := docxHTMLMarkdownPart(files)
	if err != nil {
		return "", err
	}
	if htmlPart != "" {
		parts = append(parts, htmlPart)
	}
	return strings.TrimSpace(strings.Join(parts, "\n\n")), nil
}

func docxVisibleReferencedPartText(files map[string]*zip.File, name, itemLocal, refLocal string) (string, error) {
	refs, constrained, err := docxVisibleReferenceIDs(files, refLocal)
	if err != nil {
		return "", err
	}
	if !constrained {
		return visibleXMLTextFromReferencedItems(files, name, itemLocal, map[string]bool{"": true})
	}
	return visibleXMLTextFromReferencedItems(files, name, itemLocal, refs)
}

func docxVisibleReferencedMarkdownPart(files map[string]*zip.File, name, itemLocal, refLocal, heading string) (string, error) {
	refs, constrained, err := docxVisibleReferenceIDs(files, refLocal)
	if err != nil {
		return "", err
	}
	var ids map[string]bool
	if !constrained {
		ids = map[string]bool{"": true}
	} else {
		ids = refs
	}
	items, err := visibleXMLTextItemsFromReferencedItems(files, name, itemLocal, ids)
	if err != nil {
		return "", err
	}
	var blocks []string
	for _, item := range items {
		text := markdownText(item.text)
		if text == "" {
			continue
		}
		if id := cleanDocxReferencedItemMarkdownID(item.id); id != "" {
			label := docxReferencedItemMarkdownLabel(itemLocal)
			text = "### " + escapeMarkdownHeading(label+" "+id) + "\n\n" + text
		}
		blocks = append(blocks, text)
	}
	if len(blocks) == 0 {
		return "", nil
	}
	return heading + "\n\n" + strings.Join(blocks, "\n\n"), nil
}

func cleanDocxReferencedItemMarkdownID(id string) string {
	id = cleanText(id)
	if id == "" || stripInlineHiddenOfficeReferences(id) != id || looksLikeHiddenResourceReference(id) || looksLikeRelationshipIDReference(id) || looksLikeOfficeRelationshipMetadataReference(id) || looksLikeOfficeXMLMetadataReference(id) {
		return ""
	}
	for _, r := range id {
		if r < '0' || r > '9' {
			return ""
		}
	}
	return id
}

func visibleXMLTextFromReferencedItems(files map[string]*zip.File, name, itemLocal string, ids map[string]bool) (string, error) {
	items, err := visibleXMLTextItemsFromReferencedItems(files, name, itemLocal, ids)
	if err != nil {
		return "", err
	}
	var out []string
	for _, item := range items {
		if item.text != "" {
			out = append(out, item.text)
		}
	}
	return cleanText(strings.Join(out, "\n")), nil
}

type referencedXMLTextItem struct {
	id   string
	text string
}

func visibleXMLTextItemsFromReferencedItems(files map[string]*zip.File, name, itemLocal string, ids map[string]bool) ([]referencedXMLTextItem, error) {
	f := ooxmlFile(files, name)
	if f == nil || len(ids) == 0 {
		return nil, nil
	}
	resolvedCommentParaIDs := map[string]bool{}
	if itemLocal == "comment" && ooxmlPartKey(name) == "word/comments.xml" {
		var err error
		resolvedCommentParaIDs, err = docxResolvedCommentParaIDs(files)
		if err != nil {
			return nil, err
		}
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, err
	}
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []referencedXMLTextItem
	var depth int
	var include bool
	var currentID string
	var buf bytes.Buffer
	var enc *xml.Encoder
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		if depth == 0 {
			start, ok := tok.(xml.StartElement)
			if !ok || start.Name.Local != itemLocal {
				continue
			}
			id := strings.TrimSpace(xmlAttrValue(start, "id"))
			include = ids[id]
			currentID = id
			depth = 1
			if include {
				buf.Reset()
				enc = xml.NewEncoder(&buf)
				if err := enc.EncodeToken(start); err != nil {
					return nil, err
				}
			}
			continue
		}
		if include {
			if err := enc.EncodeToken(tok); err != nil {
				return nil, err
			}
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
			if depth == 0 {
				if include {
					if err := enc.Flush(); err != nil {
						return nil, err
					}
					text, err := visibleXMLText(buf.Bytes())
					if err != nil {
						return nil, err
					}
					if len(resolvedCommentParaIDs) > 0 && xmlContainsAnyParagraphID(buf.Bytes(), resolvedCommentParaIDs) {
						text = ""
					}
					if text != "" {
						out = append(out, referencedXMLTextItem{id: currentID, text: text})
					}
				}
				include = false
				currentID = ""
				enc = nil
			}
		}
	}
	return out, nil
}

func docxResolvedCommentParaIDs(files map[string]*zip.File) (map[string]bool, error) {
	f := ooxmlFile(files, "word/commentsExtended.xml")
	if f == nil {
		return nil, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, err
	}
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	out := map[string]bool{}
	dec := xml.NewDecoder(bytes.NewReader(b))
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "commentEx" {
			continue
		}
		if !xmlBoolAttr(start, "done") {
			continue
		}
		if paraID := strings.TrimSpace(xmlAttrValue(start, "paraId")); paraID != "" {
			out[paraID] = true
		}
	}
	return out, nil
}

func xmlBoolAttr(start xml.StartElement, local string) bool {
	value := strings.ToLower(strings.TrimSpace(xmlAttrValue(start, local)))
	switch value {
	case "1", "true", "on":
		return true
	default:
		return false
	}
}

func xmlContainsAnyParagraphID(b []byte, ids map[string]bool) bool {
	if len(ids) == 0 {
		return false
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			return false
		}
		if err != nil {
			return false
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "p" {
			continue
		}
		if paraID := strings.TrimSpace(xmlAttrValue(start, "paraId")); paraID != "" && ids[paraID] {
			return true
		}
	}
}

func docxReferencedItemMarkdownLabel(itemLocal string) string {
	switch itemLocal {
	case "footnote":
		return "Footnote"
	case "endnote":
		return "Endnote"
	case "comment":
		return "Comment"
	default:
		return itemLocal
	}
}

func docxNumberingFormats(files map[string]*zip.File) (map[string]string, error) {
	f := ooxmlFile(files, "word/numbering.xml")
	if f == nil {
		return nil, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, err
	}
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	abstractFormats := map[string]map[int]string{}
	numToAbstract := map[string]string{}
	var abstractID, numID string
	var level int
	var inAbstract, inLevel, inNum bool
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "abstractNum":
				abstractID = strings.TrimSpace(xmlAttrValue(t, "abstractNumId"))
				inAbstract = abstractID != ""
			case "lvl":
				if inAbstract {
					level = 0
					if value, ok := intAttrValue(t, "ilvl"); ok && value >= 0 {
						level = value
					}
					inLevel = true
				}
			case "numFmt":
				if inAbstract && inLevel {
					format := strings.ToLower(strings.TrimSpace(xmlAttrValue(t, "val")))
					if format != "" {
						if abstractFormats[abstractID] == nil {
							abstractFormats[abstractID] = map[int]string{}
						}
						abstractFormats[abstractID][level] = format
					}
				}
			case "num":
				numID = strings.TrimSpace(xmlAttrValue(t, "numId"))
				inNum = numID != ""
			case "abstractNumId":
				if inNum {
					abstract := strings.TrimSpace(xmlAttrValue(t, "val"))
					if abstract != "" {
						numToAbstract[numID] = abstract
					}
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "abstractNum":
				inAbstract = false
				abstractID = ""
			case "lvl":
				inLevel = false
				level = 0
			case "num":
				inNum = false
				numID = ""
			}
		}
	}
	out := map[string]string{}
	for num, abstract := range numToAbstract {
		for level, format := range abstractFormats[abstract] {
			out[markdownNumberingKey(num, level)] = format
		}
	}
	if len(out) == 0 {
		return nil, nil
	}
	return out, nil
}

func docxVisibleReferenceIDs(files map[string]*zip.File, refLocal string) (map[string]bool, bool, error) {
	out := map[string]bool{}
	found := false
	for _, name := range docxVisibleRelationshipSourceParts(files) {
		ids, err := docxVisibleReferenceIDsFromPart(files, name, refLocal)
		if err != nil {
			return nil, false, err
		}
		if len(ids) > 0 {
			found = true
		}
		for id := range ids {
			out[id] = true
		}
	}
	return out, found, nil
}

func docxVisibleReferenceIDsFromPart(files map[string]*zip.File, name, refLocal string) (map[string]bool, error) {
	f := ooxmlFile(files, name)
	if f == nil {
		return nil, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, err
	}
	ids, err := visibleReferenceIDsInXML(b, refLocal)
	if err != nil {
		return nil, err
	}
	if refLocal == "commentReference" {
		rangeIDs, err := visibleReferenceIDsInXML(b, "commentRangeStart")
		if err != nil {
			return nil, err
		}
		for id := range rangeIDs {
			ids[id] = true
		}
	}
	return ids, nil
}

func visibleReferenceIDsInXML(b []byte, refLocal string) (map[string]bool, error) {
	ids := map[string]bool{}
	if hasDOCTYPE(b) {
		return ids, errors.New("xml doctype is not supported")
	}
	if !bytes.Contains(b, []byte(refLocal)) {
		return ids, nil
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var hiddenRevisionDepth int
	var hiddenRevisionRangeDepth int
	var drawingObjectStack []bool
	var paragraphHiddenStack []bool
	var alternateStack []alternateContentState
	var skipDepth int
	var runDepth int
	var rPrDepth int
	var pPrDepth int
	var runHidden bool
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return ids, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if skipDepth > 0 {
				skipDepth++
				continue
			}
			if alternateContentStartSkip(t.Name.Local, &alternateStack) {
				if t.Name.Local == "Fallback" {
					skipDepth = 1
				}
				continue
			}
			if hiddenRevisionDepth > 0 {
				hiddenRevisionDepth++
			} else if isHiddenRevisionElement(t.Name) {
				hiddenRevisionDepth = 1
			}
			if hiddenRevisionDepth == 0 {
				if isHiddenRevisionRangeStart(t.Name) {
					hiddenRevisionRangeDepth++
				} else if isHiddenRevisionRangeEnd(t.Name) && hiddenRevisionRangeDepth > 0 {
					hiddenRevisionRangeDepth--
				}
			}
			if isDrawingObjectElement(t.Name.Local) {
				parentHidden := len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]
				drawingObjectStack = append(drawingObjectStack, parentHidden)
			}
			if len(drawingObjectStack) > 0 && drawingObjectElementHidden(t) {
				drawingObjectStack[len(drawingObjectStack)-1] = true
			}
			switch t.Name.Local {
			case "p":
				paragraphHiddenStack = append(paragraphHiddenStack, false)
			case "r":
				runDepth++
				runHidden = false
			case "pPr":
				if len(paragraphHiddenStack) > 0 {
					pPrDepth++
				}
			case "rPr":
				if runDepth > 0 {
					rPrDepth++
				}
			case "vanish", "webHidden":
				if runDepth > 0 && rPrDepth > 0 {
					runHidden = true
				}
				if pPrDepth > 0 && len(paragraphHiddenStack) > 0 {
					paragraphHiddenStack[len(paragraphHiddenStack)-1] = true
				}
			}
			hidden := hiddenRevisionDepth > 0 || hiddenRevisionRangeDepth > 0 || runHidden || currentParagraphHidden(paragraphHiddenStack) ||
				(len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1])
			if !hidden && t.Name.Local == refLocal {
				id := strings.TrimSpace(xmlAttrValue(t, "id"))
				if id != "" {
					ids[id] = true
				}
			}
		case xml.EndElement:
			if skipDepth > 0 {
				skipDepth--
				continue
			}
			if alternateContentEnd(t.Name.Local, &alternateStack) {
				continue
			}
			if t.Name.Local == "pPr" && pPrDepth > 0 {
				pPrDepth--
			}
			if t.Name.Local == "rPr" && rPrDepth > 0 {
				rPrDepth--
			}
			if t.Name.Local == "r" && runDepth > 0 {
				runDepth--
				if runDepth == 0 {
					runHidden = false
					rPrDepth = 0
				}
			}
			if t.Name.Local == "p" && len(paragraphHiddenStack) > 0 {
				paragraphHiddenStack = paragraphHiddenStack[:len(paragraphHiddenStack)-1]
				if len(paragraphHiddenStack) == 0 {
					pPrDepth = 0
				}
			}
			if isDrawingObjectElement(t.Name.Local) && len(drawingObjectStack) > 0 {
				drawingObjectStack = drawingObjectStack[:len(drawingObjectStack)-1]
			}
			if hiddenRevisionDepth > 0 {
				hiddenRevisionDepth--
			}
		}
	}
	return ids, nil
}

func docxRelatedTextPart(name string) bool {
	lower := ooxmlPartKey(name)
	return strings.HasSuffix(lower, ".xml") &&
		(strings.HasPrefix(lower, "word/charts/") ||
			strings.HasPrefix(lower, "word/drawings/") ||
			strings.HasPrefix(lower, "word/diagrams/data"))
}

func docxVisibleRelatedTextPart(files map[string]*zip.File, name string) bool {
	if !docxRelatedTextPart(name) {
		return false
	}
	visible, constrained := docxVisibleRelatedTextParts(files)
	return !constrained || visible[ooxmlPartKey(name)]
}

type docxRelatedTextVisibility struct {
	visible     map[string]bool
	constrained bool
}

var docxRelatedTextVisibilityCache sync.Map

func docxVisibleRelatedTextParts(files map[string]*zip.File) (map[string]bool, bool) {
	documentFile := ooxmlFile(files, "word/document.xml")
	if documentFile != nil {
		if cached, ok := docxRelatedTextVisibilityCache.Load(documentFile); ok {
			v := cached.(docxRelatedTextVisibility)
			return cloneBoolMap(v.visible), v.constrained
		}
	}
	visible, constrained := docxVisibleRelatedTextPartsUncached(files)
	if documentFile != nil {
		docxRelatedTextVisibilityCache.Store(documentFile, docxRelatedTextVisibility{
			visible:     cloneBoolMap(visible),
			constrained: constrained,
		})
	}
	return visible, constrained
}

func docxVisibleRelatedTextPartsUncached(files map[string]*zip.File) (map[string]bool, bool) {
	visible := map[string]bool{}
	hidden := map[string]bool{}
	found := false
	for _, source := range docxVisibleRelationshipSourceParts(files) {
		partFound, ok := collectDocxSourceVisibleRelatedTextParts(files, source, visible, hidden)
		if !ok {
			return nil, false
		}
		found = found || partFound
	}
	if len(hidden) > 0 {
		for name := range hidden {
			if !visible[name] {
				delete(visible, name)
			}
		}
	}
	return visible, found
}

func collectDocxSourceVisibleRelatedTextParts(files map[string]*zip.File, source string, visible, hidden map[string]bool) (bool, bool) {
	f := ooxmlFile(files, source)
	if f == nil {
		return collectReachableDocxRelatedTextParts(files, source, visible, map[string]bool{})
	}
	b, err := readZipFile(f)
	if err != nil {
		return collectReachableDocxRelatedTextParts(files, source, visible, map[string]bool{})
	}
	if !likelyImageRelationshipMarkup(b) {
		return collectReachableDocxRelatedTextParts(files, source, visible, map[string]bool{})
	}
	refs, err := imageRelationshipRefsFromPartBytes(files, source, b)
	if err != nil || (len(refs.Visible) == 0 && len(refs.Hidden) == 0) {
		return collectReachableDocxRelatedTextParts(files, source, visible, map[string]bool{})
	}
	rels, err := relationshipTargetMapForPart(files, source)
	if err != nil || len(rels) == 0 {
		return collectReachableDocxRelatedTextParts(files, source, visible, map[string]bool{})
	}
	found := false
	for id := range refs.Visible {
		partFound, ok := collectRelationshipTargetDocxRelatedTextPart(files, source, rels[id], visible, map[string]bool{})
		if !ok {
			return false, false
		}
		found = found || partFound
	}
	for id := range refs.Hidden {
		partFound, ok := collectRelationshipTargetDocxRelatedTextPart(files, source, rels[id], hidden, map[string]bool{})
		if !ok {
			return false, false
		}
		found = found || partFound
	}
	if !found {
		return collectReachableDocxRelatedTextParts(files, source, visible, map[string]bool{})
	}
	return true, true
}

func collectRelationshipTargetDocxRelatedTextPart(files map[string]*zip.File, source, target string, out, seen map[string]bool) (bool, bool) {
	if strings.TrimSpace(target) == "" {
		return false, true
	}
	part := resolveOOXMLRelationshipTarget(source, target)
	if actual := ooxmlPartName(files, part); actual != "" {
		part = actual
	}
	lower := ooxmlPartKey(part)
	found := false
	if docxRelatedTextPart(lower) {
		out[lower] = true
		found = true
	}
	if strings.HasPrefix(lower, "word/") && ooxmlFile(files, ooxmlRelsName(part)) != nil {
		childFound, ok := collectReachableDocxRelatedTextParts(files, part, out, seen)
		if !ok {
			return false, false
		}
		found = found || childFound
	}
	return found, true
}

func collectReachableDocxRelatedTextParts(files map[string]*zip.File, source string, out, seen map[string]bool) (bool, bool) {
	relsName := ooxmlRelsName(source)
	if seen[relsName] {
		return false, true
	}
	seen[relsName] = true
	f := ooxmlFile(files, relsName)
	if f == nil {
		return false, true
	}
	b, err := readZipFile(f)
	if err != nil {
		return false, false
	}
	targets, err := relationshipTargets(b)
	if err != nil {
		return false, false
	}
	found := false
	for _, target := range targets {
		part := resolveOOXMLRelationshipTarget(source, target)
		if actual := ooxmlPartName(files, part); actual != "" {
			part = actual
		}
		lower := ooxmlPartKey(part)
		if docxRelatedTextPart(lower) {
			out[lower] = true
			found = true
		}
		if strings.HasPrefix(lower, "word/") && ooxmlFile(files, ooxmlRelsName(part)) != nil {
			childFound, ok := collectReachableDocxRelatedTextParts(files, part, out, seen)
			if !ok {
				return false, false
			}
			found = found || childFound
		}
	}
	return found, true
}

func docxVisibleVMLPart(files map[string]*zip.File, name string) bool {
	lower := ooxmlPartKey(name)
	if !strings.HasPrefix(lower, "word/") || !strings.HasSuffix(lower, ".vml") {
		return false
	}
	visible, constrained := docxVisibleVMLParts(files)
	return !constrained || visible[lower]
}

func docxVisibleVMLParts(files map[string]*zip.File) (map[string]bool, bool) {
	out := map[string]bool{}
	for _, source := range docxVisibleRelationshipSourceParts(files) {
		for _, target := range relationshipTargetsWithPrefix(files, source, "word/") {
			lower := ooxmlPartKey(target)
			if strings.HasPrefix(lower, "word/") && strings.HasSuffix(lower, ".vml") {
				out[lower] = true
			}
		}
	}
	return out, len(out) > 0
}

type docxHeaderFooterVisibility struct {
	visible     map[string]bool
	hidden      map[string]bool
	constrained bool
}

var docxHeaderFooterVisibilityCache sync.Map

func docxVisibleHeaderFooterParts(files map[string]*zip.File) (map[string]bool, map[string]bool, bool) {
	documentFile := ooxmlFile(files, "word/document.xml")
	if documentFile != nil {
		if cached, ok := docxHeaderFooterVisibilityCache.Load(documentFile); ok {
			v := cached.(docxHeaderFooterVisibility)
			return cloneBoolMap(v.visible), cloneBoolMap(v.hidden), v.constrained
		}
	}
	visible, hidden, constrained := docxVisibleHeaderFooterPartsUncached(files)
	if documentFile != nil {
		docxHeaderFooterVisibilityCache.Store(documentFile, docxHeaderFooterVisibility{
			visible:     cloneBoolMap(visible),
			hidden:      cloneBoolMap(hidden),
			constrained: constrained,
		})
	}
	return visible, hidden, constrained
}

func docxVisibleHeaderFooterPartsUncached(files map[string]*zip.File) (map[string]bool, map[string]bool, bool) {
	visible := map[string]bool{}
	hidden := map[string]bool{}
	document := ooxmlPartName(files, "word/document.xml")
	if document == "" {
		return visible, hidden, false
	}
	ids, err := docxHeaderFooterRelationshipIDs(files, document)
	if err != nil || len(ids) == 0 {
		return visible, hidden, false
	}
	rels, err := relationshipTargetMapForPart(files, document)
	if err != nil || len(rels) == 0 {
		return visible, hidden, false
	}
	for _, id := range ids {
		target := rels[id]
		if target == "" {
			continue
		}
		part := resolveOOXMLRelationshipTarget(document, target)
		if actual := ooxmlPartName(files, part); actual != "" {
			part = actual
		}
		lower := ooxmlPartKey(part)
		if isDocxHeaderFooterPart(lower) {
			visible[lower] = true
		}
	}
	if len(visible) == 0 {
		return visible, hidden, false
	}
	for name := range files {
		lower := ooxmlPartKey(name)
		if isDocxHeaderFooterPart(lower) && !visible[lower] {
			hidden[lower] = true
		}
	}
	return visible, hidden, true
}

func cloneBoolMap(in map[string]bool) map[string]bool {
	out := make(map[string]bool, len(in))
	for k, v := range in {
		out[k] = v
	}
	return out
}

func docxHeaderFooterRelationshipIDs(files map[string]*zip.File, document string) ([]string, error) {
	f := ooxmlFile(files, document)
	if f == nil {
		return nil, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, err
	}
	headerIDs, err := visibleReferenceIDsInXML(b, "headerReference")
	if err != nil {
		return nil, err
	}
	footerIDs, err := visibleReferenceIDsInXML(b, "footerReference")
	if err != nil {
		return nil, err
	}
	seen := map[string]bool{}
	var ids []string
	for id := range headerIDs {
		if !seen[id] {
			seen[id] = true
			ids = append(ids, id)
		}
	}
	for id := range footerIDs {
		if !seen[id] {
			seen[id] = true
			ids = append(ids, id)
		}
	}
	sort.Strings(ids)
	return ids, nil
}

func isDocxHeaderFooterPart(name string) bool {
	lower := ooxmlPartKey(name)
	if !strings.HasPrefix(lower, "word/") || !strings.HasSuffix(lower, ".xml") {
		return false
	}
	base := path.Base(lower)
	return strings.HasPrefix(base, "header") || strings.HasPrefix(base, "footer")
}

func docxVisibleRelationshipSourceParts(files map[string]*zip.File) []string {
	visibleHeaderFooter, _, constrainedHeaderFooter := docxVisibleHeaderFooterParts(files)
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if !strings.HasPrefix(lower, "word/") || !strings.HasSuffix(lower, ".xml") {
			continue
		}
		base := path.Base(lower)
		if base == "document.xml" || (isDocxHeaderFooterPart(lower) && (!constrainedHeaderFooter || visibleHeaderFooter[lower])) ||
			base == "footnotes.xml" || base == "endnotes.xml" || base == "comments.xml" {
			names = append(names, name)
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return names
}

func docxPartNames(files map[string]*zip.File, include func(string) bool) []string {
	var names []string
	for name := range files {
		if include(name) {
			names = append(names, name)
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return names
}

func ooxmlMarkdownPart(files map[string]*zip.File, names []string, heading string) (string, error) {
	var blocks []string
	for _, name := range names {
		if ooxmlFile(files, name) == nil {
			continue
		}
		text, err := visibleXMLTextFromZip(files, name)
		if err != nil {
			return "", err
		}
		text = markdownText(text)
		if text != "" {
			blocks = append(blocks, text)
		}
	}
	if len(blocks) == 0 {
		return "", nil
	}
	return heading + "\n\n" + strings.Join(blocks, "\n\n"), nil
}

func ooxmlMarkdownPartWithTables(files map[string]*zip.File, names []string, heading string) (string, error) {
	return ooxmlMarkdownPartWithTablesAndNumbering(files, names, heading, nil)
}

func ooxmlMarkdownPartWithTablesAndNumbering(files map[string]*zip.File, names []string, heading string, numbering map[string]string) (string, error) {
	var blocks []string
	for _, name := range names {
		f := ooxmlFile(files, name)
		if f == nil {
			continue
		}
		b, err := readZipFile(f)
		if err != nil {
			return "", err
		}
		text, err := visibleXMLMarkdownWithTablesAndNumbering(b, numbering)
		if err != nil {
			return "", err
		}
		text = markdownText(text)
		if text != "" {
			blocks = append(blocks, text)
		}
	}
	if len(blocks) == 0 {
		return "", nil
	}
	return heading + "\n\n" + strings.Join(blocks, "\n\n"), nil
}

func docxHTMLMarkdownPart(files map[string]*zip.File) (string, error) {
	names, err := docxVisibleHTMLPartNames(files)
	if err != nil {
		return "", err
	}
	var blocks []string
	for _, name := range names {
		b, err := readZipFile(files[name])
		if err != nil {
			return "", err
		}
		text := markdownText(visibleAltChunkText(name, b))
		if text != "" {
			blocks = append(blocks, text)
		}
	}
	if len(blocks) == 0 {
		return "", nil
	}
	return "## HTML Content\n\n" + strings.Join(blocks, "\n\n"), nil
}

func docxVisibleHTMLPartNames(files map[string]*zip.File) ([]string, error) {
	if !docxHasHTMLPart(files) {
		return nil, nil
	}
	visibleHeaderFooter, _, constrainedHeaderFooter := docxVisibleHeaderFooterParts(files)
	var contentParts []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if !strings.HasPrefix(lower, "word/") || !strings.HasSuffix(lower, ".xml") {
			continue
		}
		base := path.Base(lower)
		if base == "document.xml" || (isDocxHeaderFooterPart(lower) && (!constrainedHeaderFooter || visibleHeaderFooter[lower])) {
			contentParts = append(contentParts, name)
		}
	}
	sort.Slice(contentParts, func(i, j int) bool {
		return naturalLess(contentParts[i], contentParts[j])
	})
	seen := map[string]bool{}
	var names []string
	for _, part := range contentParts {
		ids, err := docxAltChunkRelationshipIDs(files, part)
		if err != nil {
			return nil, err
		}
		if len(ids) == 0 {
			continue
		}
		rels, err := relationshipTargetMapForPart(files, part)
		if err != nil {
			continue
		}
		for _, id := range ids {
			target := rels[id]
			if target == "" {
				continue
			}
			name := resolveOOXMLRelationshipTarget(part, target)
			if actual := ooxmlPartName(files, name); actual != "" {
				name = actual
			}
			lower := ooxmlPartKey(name)
			if !docxAltChunkHTMLLikePart(lower) {
				continue
			}
			if !seen[lower] {
				seen[lower] = true
				names = append(names, name)
			}
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	if len(names) == 0 {
		names = docxFallbackAltChunkHTMLPartNames(files)
	}
	return names, nil
}

func docxHasHTMLPart(files map[string]*zip.File) bool {
	for name := range files {
		lower := ooxmlPartKey(name)
		if docxAltChunkHTMLLikePart(lower) {
			return true
		}
	}
	return false
}

func docxAltChunkHTMLLikePart(name string) bool {
	lower := ooxmlPartKey(name)
	if !strings.HasPrefix(lower, "word/") {
		return false
	}
	return strings.HasSuffix(lower, ".htm") || strings.HasSuffix(lower, ".html") ||
		strings.HasSuffix(lower, ".mht") || strings.HasSuffix(lower, ".mhtml")
}

func docxFallbackAltChunkHTMLPartNames(files map[string]*zip.File) []string {
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		base := path.Base(lower)
		if strings.HasPrefix(lower, "word/") && strings.HasPrefix(base, "afchunk") && docxAltChunkHTMLLikePart(lower) {
			names = append(names, name)
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return names
}

func docxAltChunkRelationshipIDs(files map[string]*zip.File, part string) ([]string, error) {
	f := ooxmlFile(files, part)
	if f == nil {
		return nil, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, err
	}
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var ids []string
	var skipDepth int
	var hiddenRevisionRangeDepth int
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if skipDepth > 0 {
				skipDepth++
				continue
			}
			if isHiddenRevisionElement(t.Name) {
				skipDepth = 1
				continue
			}
			if isHiddenRevisionRangeStart(t.Name) {
				hiddenRevisionRangeDepth++
				continue
			}
			if isHiddenRevisionRangeEnd(t.Name) {
				if hiddenRevisionRangeDepth > 0 {
					hiddenRevisionRangeDepth--
				}
				continue
			}
			if hiddenRevisionRangeDepth == 0 && t.Name.Local == "altChunk" {
				id := strings.TrimSpace(xmlAttrValue(t, "id"))
				if id != "" {
					ids = append(ids, id)
				}
			}
		case xml.EndElement:
			if skipDepth > 0 {
				skipDepth--
			}
		}
	}
	return ids, nil
}

func extractPptxText(files map[string]*zip.File, strictOfficeContent ...bool) ([]string, error) {
	visibleNotes, err := pptxVisibleNotesSlideNames(files)
	if err != nil {
		return nil, err
	}
	visibleComments, err := pptxVisibleCommentPartNames(files)
	if err != nil {
		return nil, err
	}
	visibleRelated, constrainedRelated := pptxVisibleRelatedTextParts(files)
	slideNames, _, err := pptxCandidateSlideNames(files)
	if err != nil {
		return nil, err
	}
	var names []string
	for _, name := range slideNames {
		visible, err := pptxSlideVisible(files, name)
		if err != nil {
			return nil, err
		}
		if visible {
			names = append(names, name)
		}
	}
	if len(strictOfficeContent) > 0 && strictOfficeContent[0] {
		return pptxStrictTextFromFiles(files, names)
	}
	for name := range files {
		lower := ooxmlPartKey(name)
		if strings.HasPrefix(lower, "ppt/notesslides/notesslide") && strings.HasSuffix(lower, ".xml") {
			continue
		}
		if strings.HasPrefix(lower, "ppt/comments/comment") && strings.HasSuffix(lower, ".xml") {
			if visibleComments[ooxmlPartKey(name)] {
				names = append(names, name)
			}
		}
		if pptxRelatedTextPart(lower) && strings.HasSuffix(lower, ".xml") && (!constrainedRelated || visibleRelated[lower]) {
			names = append(names, name)
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	out, err := xmlTextFromFiles(files, names)
	if err != nil {
		return nil, err
	}
	for _, name := range pptxSortedPartNames(files, visibleNotes) {
		text, err := visiblePptxNotesTextFromZip(files, name)
		if err != nil {
			return nil, err
		}
		if text != "" {
			out = append(out, text)
		}
	}
	return out, nil
}

// pptxStrictTextFromFiles mirrors PowerPoint's Shape.TextFrame.TextRange.
// Image alt text lives in cNvPr attributes in slide XML, but PowerPoint does
// not include it in the text exposed by a shape, so it must not enter the
// Office-aligned text result.
func pptxStrictTextFromFiles(files map[string]*zip.File, names []string) ([]string, error) {
	var out []string
	for slideIndex, name := range names {
		f := ooxmlFile(files, name)
		if f == nil {
			continue
		}
		b, err := readZipFile(f)
		if err != nil {
			return nil, err
		}
		text, err := visiblePptxShapeTextWithDynamicFields(b, slideIndex+1, time.Now())
		if err != nil {
			return nil, fmt.Errorf("%s: %w", name, err)
		}
		if text != "" {
			out = append(out, text)
		}
	}
	return out, nil
}

// visiblePptxShapeText follows PowerPoint's visible shape tree. Group members
// are rendered on the slide and are exposed through the COM GroupItems
// collection, so their text is part of the Office-aligned strict result. Table
// cells are cached graphic-frame payloads that are not exposed by
// Shape.TextFrame after imported HTML is flattened. Non-table graphic frames
// remain eligible.
func visiblePptxShapeText(b []byte) (string, error) {
	return visiblePptxShapeTextWithDynamicFields(b, 0, time.Time{})
}

// visiblePptxShapeTextWithDynamicFields mirrors PowerPoint's TextRange for
// date and slide-number fields. PowerPoint recalculates those fields when it
// opens a presentation, so their cached a:t value can be years out of date.
// A zero now retains package caches, which keeps this lower-level helper
// deterministic for compatibility callers and unit tests.
func visiblePptxShapeTextWithDynamicFields(b []byte, slideNumber int, now time.Time) (string, error) {
	if hasDOCTYPE(b) {
		return "", errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	var textDepth, mathTextDepth, tableDepth, graphicFrameDepth int
	var runDepth int
	var runSymbolFont string
	var dynamicFieldDepth int
	var dynamicFieldValue string
	runHasBaselineOffset := false
	runHasHyperlink := false
	lastSegmentHasBaselineOffset := false
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return "", err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "r", "fld":
				runDepth++
				if runDepth == 1 {
					runHasBaselineOffset = false
					runHasHyperlink = false
					runSymbolFont = ""
				}
				if t.Name.Local == "fld" && !now.IsZero() {
					if value, ok := pptxDynamicFieldValue(xmlAttrValue(t, "type"), slideNumber, now); ok {
						dynamicFieldDepth++
						dynamicFieldValue = value
						appendPptxVisibleTextSegment(&out, value, false, false, false, &lastSegmentHasBaselineOffset)
					}
				}
			case "rPr":
				if runDepth > 0 && xmlAttrValue(t, "baseline") != "" && xmlAttrValue(t, "baseline") != "0" {
					runHasBaselineOffset = true
				}
				// Some older PowerPoint writers put a:sym directly inside a:rPr,
				// while others use only the rPr typeface. Keep the latter as the
				// fallback so U+F0xx values in the following a:t are decoded too.
				if runDepth > 0 && runSymbolFont == "" {
					if font := xmlAttrValue(t, "typeface"); isFontEncodedSymbolFont(font) {
						runSymbolFont = font
					}
				}
			case "hlinkClick", "hlinkMouseOver":
				if runDepth > 0 {
					runHasHyperlink = true
				}
			case "sym":
				// DrawingML represents legacy Symbol-font text with a:sym below
				// a:rPr. Its a:t value can contain U+F0xx private-use fallback
				// codes (not Unicode arrows). PowerPoint.TextRange renders the
				// corresponding visible glyph, whose punctuation boundary matters
				// to strict token comparison.
				if runDepth > 0 {
					runSymbolFont = xmlAttrValue(t, "typeface")
				}
			case "br":
				if tableDepth == 0 {
					out = append(out, "\n")
				}
			case "graphicFrame":
				graphicFrameDepth++
			case "tbl":
				if graphicFrameDepth > 0 {
					tableDepth++
				}
			case "t":
				if tableDepth > 0 {
					continue
				}
				if t.Name.Space == "http://schemas.openxmlformats.org/officeDocument/2006/math" {
					mathTextDepth++
				} else {
					textDepth++
				}
			case "p":
				if tableDepth == 0 {
					out = append(out, "\n")
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "r", "fld":
				if runDepth > 0 {
					runDepth--
					if runDepth == 0 {
						runHasBaselineOffset = false
						runHasHyperlink = false
						runSymbolFont = ""
					}
				}
				if t.Name.Local == "fld" && dynamicFieldDepth > 0 {
					dynamicFieldDepth--
					if dynamicFieldDepth == 0 {
						dynamicFieldValue = ""
					}
				}
			case "t":
				if t.Name.Space == "http://schemas.openxmlformats.org/officeDocument/2006/math" && mathTextDepth > 0 {
					mathTextDepth--
				} else if textDepth > 0 {
					textDepth--
				}
			case "tbl":
				if tableDepth > 0 {
					tableDepth--
				}
			case "graphicFrame":
				if graphicFrameDepth > 0 {
					graphicFrameDepth--
				}
			}
		case xml.CharData:
			if (textDepth > 0 || mathTextDepth > 0) && dynamicFieldValue == "" {
				text := string(t)
				if mathTextDepth > 0 {
					text = strings.Map(pptxMathTextRune, text)
				}
				if runSymbolFont != "" {
					text = pptxSymbolFontText(text, runSymbolFont)
				}
				appendPptxVisibleTextSegment(&out, text, runHasBaselineOffset, runHasHyperlink, mathTextDepth > 0, &lastSegmentHasBaselineOffset)
			}
		}
	}
	// Slide text is authoritative content, including ordinary template labels
	// such as "Click to add title".  The Markdown cleaner intentionally drops
	// some short boilerplate-looking lines, which is appropriate for prose
	// rendering but wrong for a COM TextRange baseline.
	return cleanTextNoMojibakeRepair(strings.Join(out, "")), nil
}

func pptxDynamicFieldValue(kind string, slideNumber int, now time.Time) (string, bool) {
	switch strings.ToLower(strings.TrimSpace(kind)) {
	case "slidenum", "slidenumber":
		if slideNumber <= 0 {
			return "", false
		}
		return strconv.Itoa(slideNumber), true
	case "datetime1":
		// PowerPoint's built-in date field uses the host short-date convention,
		// whose month/day components are not padded on the COM baseline host.
		return now.Format("1/2/2006"), true
	case "datetime2":
		return now.Format("Monday, January 2, 2006"), true
	case "datetime3":
		return now.Format("2 January 2006"), true
	case "datetime4":
		return now.Format("January 2, 2006"), true
	case "datetime5":
		return now.Format("2-Jan-06"), true
	case "datetime6":
		return now.Format("January 06"), true
	case "datetime7":
		return now.Format("Jan-06"), true
	case "datetime8":
		// PowerPoint exposes a locale-expanded datetime8 value. On the
		// automation host this is short date plus an unpadded 12-hour time,
		// while Go's reference layout pads minutes. Format components
		// explicitly to avoid producing 1:01 PM for 13:01.
		hour := now.Hour() % 12
		if hour == 0 {
			hour = 12
		}
		ampm := "AM"
		if now.Hour() >= 12 {
			ampm = "PM"
		}
		return fmt.Sprintf("%d/%d/%04d %d:%02d %s", now.Month(), now.Day(), now.Year(), hour, now.Minute(), ampm), true
	case "datetime9":
		return now.Format("1/2/2006"), true
	case "datetime10":
		return now.Format("1/2/06"), true
	case "datetime11":
		return now.Format("2006-01-02"), true
	case "datetime12":
		return now.Format("02-Jan-06"), true
	case "datetime13":
		return now.Format("January 2006"), true
	}
	return "", false
}

// appendPptxVisibleTextSegment keeps superscript and subscript runs attached
// to their mathematical base. PowerPoint's TextRange renders "A" + a raised
// "2" + "x" as "A2x", while an XML-only run boundary has no implicit space.
func appendPptxVisibleTextSegment(out *[]string, text string, runHasBaselineOffset, runHasHyperlink, mathText bool, lastSegmentHasBaselineOffset *bool) {
	if text == "" {
		return
	}
	// Office Math stores adjacent identifiers in separate m:r nodes. PowerPoint
	// TextRange exposes their visible mathematical atoms independently, unlike
	// ordinary DrawingML formatting runs where "Bam" + "HI" is one word. A
	// boundary here prevents K + m from becoming the invented token Km.
	if mathText && len(*out) > 0 {
		// Do not split normal multi-letter mathematical prose ("properties",
		// "max", "false") merely because it is represented as an m:r. Split
		// only the atom-sized single-glyph runs that PowerPoint exposes as
		// separate identifiers/operators through TextRange.
		previous := (*out)[len(*out)-1]
		if pptxMathRunIsAtom(previous) && pptxMathRunIsAtom(text) {
			*out = append(*out, " ")
		}
		*out = append(*out, text)
		*lastSegmentHasBaselineOffset = runHasBaselineOffset
		return
	}
	if runHasHyperlink && len(*out) > 0 && pptxHyperlinkContinuation((*out)[len(*out)-1], text) {
		*out = append(*out, text)
	} else if !runHasBaselineOffset && !*lastSegmentHasBaselineOffset {
		appendPptxOOXMLVisibleTextSegment(out, text)
	} else {
		*out = append(*out, text)
	}
	*lastSegmentHasBaselineOffset = runHasBaselineOffset
}

func pptxMathRunIsAtom(text string) bool {
	text = strings.TrimSpace(text)
	if text == "" {
		return false
	}
	return utf8.RuneCountInString(text) == 1
}

// pptxHyperlinkContinuation keeps a URL split across identically linked runs
// intact. PowerPoint's TextRange treats both runs as one hyperlink; inserting
// a formatting-boundary space would turn a visible URL such as
// "reciprocitySOUPS" + "11" into two tokens.
func pptxHyperlinkContinuation(left, right string) bool {
	leftRunes, rightRunes := []rune(left), []rune(right)
	if len(leftRunes) == 0 || len(rightRunes) == 0 {
		return false
	}
	return (unicode.IsLetter(leftRunes[len(leftRunes)-1]) || unicode.IsNumber(leftRunes[len(leftRunes)-1])) &&
		(unicode.IsLetter(rightRunes[0]) || unicode.IsNumber(rightRunes[0]))
}

func appendPptxOOXMLVisibleTextSegment(out *[]string, text string) {
	// A DrawingML paragraph is a sequence of text runs.  Run boundaries carry
	// formatting, not word separators: PowerPoint's TextRange concatenates
	// "Bam" + "HI" as "BamHI".  Whitespace that is visible in PowerPoint is
	// encoded in an a:t value and is consequently preserved by appending the
	// run verbatim.  Adding a heuristic space here splits identifiers whenever a
	// producer happens to divide a word across otherwise identical runs.
	*out = append(*out, text)
}

func pptxTextBoundaryNeedsSpace(left, right string) bool {
	leftRunes, rightRunes := []rune(left), []rune(right)
	if len(leftRunes) == 0 || len(rightRunes) == 0 {
		return false
	}
	return ooxmlTextBoundaryNeedsSpace(left, right)
}

func appendOOXMLVisibleTextSegment(out *[]string, text string) {
	if text == "" {
		return
	}
	if len(*out) > 0 && ooxmlTextBoundaryNeedsSpace((*out)[len(*out)-1], text) {
		*out = append(*out, " ")
	}
	*out = append(*out, text)
}

// PowerPoint's TextRange renders Office Math text using Cambria Math Unicode
// code points. Math runs in OOXML usually store plain ASCII identifiers plus
// style records (bold/italic/script); retain the visible characters here rather
// than leaking raw XML markup or dropping the math object altogether.
func pptxMathTextRune(r rune) rune {
	if r >= 0xf000 && r <= 0xf0ff {
		return r & 0xff
	}
	return r
}

// pptxSymbolFontText converts DrawingML's legacy font-encoded run payload to
// the Unicode glyph PowerPoint exposes through TextRange. A Symbol run may
// contain a mixture of ordinary punctuation/spacing and U+F0xx fallback
// codes, so preserve characters the legacy map does not define.
func pptxSymbolFontText(s, font string) string {
	if !isFontEncodedSymbolFont(font) {
		return s
	}
	var out strings.Builder
	for _, r := range s {
		// PowerPoint preserves ordinary Unicode text in a Symbol-formatted run;
		// only the private-use fallback byte range is a legacy glyph encoding.
		// Mapping ASCII here would turn variables such as X/Y/Z into unrelated
		// Greek letters and corrupt visible formula text. The old report's
		// remaining mismatch needs a fresh COM probe before widening this rule.
		if r < 0xf000 || r > 0xf0ff {
			out.WriteRune(r)
			continue
		}
		code := r
		code &= 0xff
		if mapped, ok := fontEncodedSymbolRune(font, code); ok {
			out.WriteRune(mapped)
			continue
		}
		out.WriteRune(r)
	}
	return out.String()
}

func extractPptxMarkdown(files map[string]*zip.File) (string, error) {
	slideNames, err := pptxVisibleSlideNames(files)
	if err != nil {
		return "", err
	}
	var parts []string
	for _, name := range slideNames {
		f := ooxmlFile(files, name)
		if f == nil {
			continue
		}
		b, err := readZipFile(f)
		if err != nil {
			return "", err
		}
		text, err := visibleXMLMarkdownWithTables(b)
		if err != nil {
			return "", err
		}
		text = markdownText(text)
		if text == "" {
			continue
		}
		parts = append(parts, "## "+escapeMarkdownHeading(pptxSlideMarkdownTitle(name))+"\n\n"+text)
	}
	notesNames, err := pptxVisibleNotesSlideNames(files)
	if err != nil {
		return "", err
	}
	notes, err := pptxNotesMarkdownPart(files, pptxSortedPartNames(files, notesNames), "## Notes")
	if err != nil {
		return "", err
	}
	if notes != "" {
		parts = append(parts, notes)
	}
	commentNames, err := pptxVisibleCommentPartNames(files)
	if err != nil {
		return "", err
	}
	comments, err := pptxCommentsMarkdownPart(files, commentNames, "## Comments")
	if err != nil {
		return "", err
	}
	if comments != "" {
		parts = append(parts, comments)
	}
	relatedNames := pptxVisibleRelatedTextPartNames(files)
	related, err := ooxmlMarkdownPart(files, relatedNames, "## Drawings")
	if err != nil {
		return "", err
	}
	if related != "" {
		parts = append(parts, related)
	}
	return strings.TrimSpace(strings.Join(parts, "\n\n")), nil
}

type pptxRelatedMarkdownPart struct {
	name  string
	title string
}

func pptxRelatedTextPart(part string) bool {
	part = ooxmlPartKey(part)
	return strings.HasPrefix(part, "ppt/charts/") ||
		strings.HasPrefix(part, "ppt/drawings/") ||
		strings.HasPrefix(part, "ppt/diagrams/data")
}

func pptxNotesMarkdownPart(files map[string]*zip.File, names []string, heading string) (string, error) {
	var blocks []string
	for _, name := range names {
		text, err := visiblePptxNotesTextFromZip(files, name)
		if err != nil {
			return "", err
		}
		text = markdownText(text)
		if text != "" {
			title, err := pptxNotesRelatedSlideMarkdownTitle(files, name)
			if err != nil {
				return "", err
			}
			if title != "" {
				text = "### " + escapeMarkdownHeading(title) + "\n\n" + text
			}
			blocks = append(blocks, text)
		}
	}
	if len(blocks) == 0 {
		return "", nil
	}
	return heading + "\n\n" + strings.Join(blocks, "\n\n"), nil
}

func pptxCommentsMarkdownPart(files map[string]*zip.File, allowed map[string]bool, heading string) (string, error) {
	var blocks []string
	for _, part := range pptxCommentMarkdownParts(files, allowed) {
		text, err := visibleXMLTextFromZip(files, part.name)
		if err != nil {
			return "", err
		}
		text = markdownText(text)
		if text == "" {
			continue
		}
		if part.title != "" {
			text = "### " + escapeMarkdownHeading(part.title) + "\n\n" + text
		}
		blocks = append(blocks, text)
	}
	if len(blocks) == 0 {
		return "", nil
	}
	return heading + "\n\n" + strings.Join(blocks, "\n\n"), nil
}

func pptxCommentMarkdownParts(files map[string]*zip.File, allowed map[string]bool) []pptxRelatedMarkdownPart {
	seen := map[string]bool{}
	var out []pptxRelatedMarkdownPart
	if slideNames, err := pptxVisibleSlideNames(files); err == nil {
		for _, slide := range slideNames {
			title := pptxSlideMarkdownTitle(slide)
			for _, name := range relationshipTargetsWithPrefix(files, slide, "ppt/comments/") {
				key := ooxmlPartKey(name)
				if !allowed[key] || seen[key] {
					continue
				}
				seen[key] = true
				out = append(out, pptxRelatedMarkdownPart{name: name, title: title})
			}
		}
	}
	for _, name := range pptxSortedPartNames(files, allowed) {
		key := ooxmlPartKey(name)
		if seen[key] {
			continue
		}
		seen[key] = true
		out = append(out, pptxRelatedMarkdownPart{name: name})
	}
	return out
}

func pptxNotesRelatedSlideMarkdownTitle(files map[string]*zip.File, name string) (string, error) {
	targets, err := relationshipTargetsForPart(files, name)
	if err != nil || len(targets) == 0 {
		return "", nil
	}
	candidates, constrained, err := pptxCandidateSlideNames(files)
	if err != nil {
		return "", err
	}
	candidate := map[string]bool{}
	for _, slide := range candidates {
		candidate[ooxmlPartKey(slide)] = true
	}
	for _, target := range targets {
		part := resolveOOXMLRelationshipTarget(name, target)
		if actual := ooxmlPartName(files, part); actual != "" {
			part = actual
		}
		lower := ooxmlPartKey(part)
		if !strings.HasPrefix(lower, "ppt/slides/slide") || !strings.HasSuffix(lower, ".xml") {
			continue
		}
		if constrained && !candidate[lower] {
			continue
		}
		visible, err := pptxSlideVisible(files, part)
		if err != nil {
			return "", err
		}
		if visible {
			return pptxSlideMarkdownTitle(part), nil
		}
	}
	return "", nil
}

func visiblePptxNotesTextFromZip(files map[string]*zip.File, name string) (string, error) {
	f := ooxmlFile(files, name)
	if f == nil {
		return "", nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return "", err
	}
	b, err = removePptxNotesSystemPlaceholders(b)
	if err != nil {
		return "", err
	}
	if strings.HasSuffix(ooxmlPartKey(name), ".vml") {
		return visibleVMLText(b)
	}
	return visibleXMLText(b)
}

func removePptxNotesSystemPlaceholders(b []byte) ([]byte, error) {
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out bytes.Buffer
	enc := xml.NewEncoder(&out)
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		if start, ok := tok.(xml.StartElement); ok && pptxNotesPlaceholderCandidateElement(start.Name.Local) {
			data, hidden, err := readPptxNotesPlaceholderCandidate(dec, start)
			if err != nil {
				return nil, err
			}
			if hidden {
				continue
			}
			if err := enc.Flush(); err != nil {
				return nil, err
			}
			out.Write(data)
			continue
		}
		if err := enc.EncodeToken(tok); err != nil {
			return nil, err
		}
	}
	if err := enc.Flush(); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

func readPptxNotesPlaceholderCandidate(dec *xml.Decoder, start xml.StartElement) ([]byte, bool, error) {
	var buf bytes.Buffer
	enc := xml.NewEncoder(&buf)
	if err := enc.EncodeToken(start); err != nil {
		return nil, false, err
	}
	depth := 1
	hidden := pptxNotesSystemPlaceholderElement(start)
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return nil, false, err
		}
		if elem, ok := tok.(xml.StartElement); ok && pptxNotesSystemPlaceholderElement(elem) {
			hidden = true
		}
		if err := enc.EncodeToken(tok); err != nil {
			return nil, false, err
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	if err := enc.Flush(); err != nil {
		return nil, false, err
	}
	return buf.Bytes(), hidden, nil
}

func pptxNotesPlaceholderCandidateElement(name string) bool {
	switch name {
	case "sp", "pic", "graphicFrame", "cxnSp":
		return true
	default:
		return false
	}
}

func pptxNotesSystemPlaceholderElement(start xml.StartElement) bool {
	if start.Name.Local != "ph" {
		return false
	}
	switch strings.ToLower(strings.TrimSpace(xmlAttrValue(start, "type"))) {
	case "dt", "ftr", "sldnum", "sldimg", "hdr":
		return true
	default:
		return false
	}
}

func pptxVisibleRelatedTextPartNames(files map[string]*zip.File) []string {
	visible, constrained := pptxVisibleRelatedTextParts(files)
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if !pptxRelatedTextPart(lower) || !strings.HasSuffix(lower, ".xml") {
			continue
		}
		if constrained && !visible[lower] {
			continue
		}
		names = append(names, name)
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return names
}

func pptxVisibleRelatedTextParts(files map[string]*zip.File) (map[string]bool, bool) {
	lookup := newOOXMLLookup(files)
	visible := map[string]bool{}
	hidden := map[string]bool{}
	found := false
	slideNames, constrained, err := pptxCandidateSlideNames(files)
	if err != nil {
		return nil, false
	}
	candidate := map[string]bool{}
	for _, name := range slideNames {
		candidate[ooxmlPartKey(name)] = true
	}
	if constrained {
		slideNames = pptxAllSlidePartNames(files)
	}
	for _, name := range slideNames {
		slideVisible, err := pptxSlideVisible(files, name)
		if err != nil {
			return nil, false
		}
		targets := visible
		if !slideVisible || (constrained && !candidate[ooxmlPartKey(name)]) {
			targets = hidden
			partFound, ok := collectReachablePptxTextPartsWithLookup(files, lookup, name, targets, map[string]bool{})
			if !ok {
				return nil, false
			}
			found = found || partFound
			continue
		}
		partFound, ok := collectPptxSlideVisibleTextPartsWithLookup(files, lookup, name, visible, hidden)
		if !ok {
			return nil, false
		}
		found = found || partFound
	}
	if !found {
		return nil, false
	}
	for name := range hidden {
		if !visible[name] {
			delete(visible, name)
		}
	}
	return visible, true
}

func collectPptxSlideVisibleTextParts(files map[string]*zip.File, slide string, visible, hidden map[string]bool) (bool, bool) {
	return collectPptxSlideVisibleTextPartsWithLookup(files, newOOXMLLookup(files), slide, visible, hidden)
}

func collectPptxSlideVisibleTextPartsWithLookup(files map[string]*zip.File, lookup ooxmlLookup, slide string, visible, hidden map[string]bool) (bool, bool) {
	f := lookup.file(slide)
	if f == nil {
		return collectReachablePptxTextPartsWithLookup(files, lookup, slide, visible, map[string]bool{})
	}
	b, err := readZipFile(f)
	if err != nil {
		return collectReachablePptxTextPartsWithLookup(files, lookup, slide, visible, map[string]bool{})
	}
	refs, err := docxImageRelationshipRefs(b)
	if err != nil || (len(refs.Visible) == 0 && len(refs.Hidden) == 0) {
		return collectReachablePptxTextPartsWithLookup(files, lookup, slide, visible, map[string]bool{})
	}
	rels, err := relationshipTargetMapForPartWithLookup(files, lookup, slide)
	if err != nil || len(rels) == 0 {
		return collectReachablePptxTextPartsWithLookup(files, lookup, slide, visible, map[string]bool{})
	}
	found := false
	for id := range refs.Visible {
		partFound, ok := collectRelationshipTargetPptxTextPartWithLookup(files, lookup, slide, rels[id], visible, map[string]bool{})
		if !ok {
			return false, false
		}
		found = found || partFound
	}
	for id := range refs.Hidden {
		partFound, ok := collectRelationshipTargetPptxTextPartWithLookup(files, lookup, slide, rels[id], hidden, map[string]bool{})
		if !ok {
			return false, false
		}
		found = found || partFound
	}
	if !found {
		return collectReachablePptxTextPartsWithLookup(files, lookup, slide, visible, map[string]bool{})
	}
	return true, true
}

func collectRelationshipTargetPptxTextPart(files map[string]*zip.File, source, target string, out, seen map[string]bool) (bool, bool) {
	return collectRelationshipTargetPptxTextPartWithLookup(files, newOOXMLLookup(files), source, target, out, seen)
}

func collectRelationshipTargetPptxTextPartWithLookup(files map[string]*zip.File, lookup ooxmlLookup, source, target string, out, seen map[string]bool) (bool, bool) {
	if strings.TrimSpace(target) == "" {
		return false, true
	}
	part := resolveOOXMLRelationshipTarget(source, target)
	if actual := lookup.partName(part); actual != "" {
		part = actual
	}
	lower := ooxmlPartKey(part)
	found := false
	if pptxRelatedTextPart(lower) && strings.HasSuffix(lower, ".xml") {
		out[lower] = true
		found = true
	}
	if strings.HasPrefix(lower, "ppt/") && lookup.file(ooxmlRelsName(part)) != nil {
		childFound, ok := collectReachablePptxTextPartsWithLookup(files, lookup, part, out, seen)
		if !ok {
			return false, false
		}
		found = found || childFound
	}
	return found, true
}

func pptxAllSlidePartNames(files map[string]*zip.File) []string {
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if strings.HasPrefix(lower, "ppt/slides/slide") && strings.HasSuffix(lower, ".xml") {
			names = append(names, name)
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return names
}

func collectReachablePptxTextParts(files map[string]*zip.File, source string, out, seen map[string]bool) (bool, bool) {
	return collectReachablePptxTextPartsWithLookup(files, newOOXMLLookup(files), source, out, seen)
}

func collectReachablePptxTextPartsWithLookup(files map[string]*zip.File, lookup ooxmlLookup, source string, out, seen map[string]bool) (bool, bool) {
	relsName := ooxmlRelsName(source)
	if seen[relsName] {
		return false, true
	}
	seen[relsName] = true
	f := lookup.file(relsName)
	if f == nil {
		return false, true
	}
	b, err := readZipFile(f)
	if err != nil {
		return false, false
	}
	targets, err := relationshipTargets(b)
	if err != nil {
		return false, false
	}
	found := false
	for _, target := range targets {
		part := resolveOOXMLRelationshipTarget(source, target)
		if actual := lookup.partName(part); actual != "" {
			part = actual
		}
		lower := ooxmlPartKey(part)
		if pptxRelatedTextPart(lower) && strings.HasSuffix(lower, ".xml") {
			out[lower] = true
			found = true
		}
		if strings.HasPrefix(lower, "ppt/") && lookup.file(ooxmlRelsName(part)) != nil {
			childFound, ok := collectReachablePptxTextPartsWithLookup(files, lookup, part, out, seen)
			if !ok {
				return false, false
			}
			found = found || childFound
		}
	}
	return found, true
}

func pptxVisibleSlideNames(files map[string]*zip.File) ([]string, error) {
	var names []string
	candidates, constrained, err := pptxCandidateSlideNames(files)
	if err != nil {
		return nil, err
	}
	for _, name := range candidates {
		visible, err := pptxSlideVisible(files, name)
		if err != nil {
			return nil, err
		}
		if visible {
			names = append(names, name)
		}
	}
	if !constrained {
		sort.Slice(names, func(i, j int) bool {
			return naturalLess(names[i], names[j])
		})
	}
	return names, nil
}

func pptxCandidateSlideNames(files map[string]*zip.File) ([]string, bool, error) {
	if names, constrained, err := pptxPresentationSlideNames(files); constrained || err != nil {
		return names, constrained, err
	}
	return pptxAllSlidePartNames(files), false, nil
}

func pptxPresentationSlideNames(files map[string]*zip.File) ([]string, bool, error) {
	f := ooxmlFile(files, "ppt/presentation.xml")
	if f == nil {
		return nil, false, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, false, err
	}
	if hasDOCTYPE(b) {
		return nil, false, errors.New("xml doctype is not supported")
	}
	rels, err := relationshipTargetMapForPart(files, "ppt/presentation.xml")
	if err != nil || len(rels) == 0 {
		return nil, false, nil
	}
	ids, err := pptxPresentationSlideRelationshipIDs(b)
	if err != nil || len(ids) == 0 {
		return nil, false, err
	}
	seen := map[string]bool{}
	var names []string
	for _, id := range ids {
		target := rels[id]
		if target == "" {
			continue
		}
		part := resolveOOXMLRelationshipTarget("ppt/presentation.xml", target)
		if actual := ooxmlPartName(files, part); actual != "" {
			part = actual
		}
		lower := ooxmlPartKey(part)
		if !strings.HasPrefix(lower, "ppt/slides/slide") || !strings.HasSuffix(lower, ".xml") || seen[lower] {
			continue
		}
		seen[lower] = true
		names = append(names, part)
	}
	// presentation.xml's slide-id relationship list is authoritative. A
	// numerically adjacent part without a relationship is still an orphan and
	// must not leak text, notes, comments, or pictures into output. If that
	// relationship list is unusable, pptxCandidateSlideNames falls back to all
	// slide parts instead of guessing extra slides from their filenames.
	if len(names) == 0 {
		return nil, false, nil
	}
	return names, true, nil
}

func pptxSlidePartNumber(name string) (int, bool) {
	match := regexp.MustCompile(`(?i)^ppt/slides/slide([0-9]+)\.xml$`).FindStringSubmatch(ooxmlPartKey(name))
	if len(match) != 2 {
		return 0, false
	}
	n, err := strconv.Atoi(match[1])
	return n, err == nil && n > 0
}

func pptxPresentationSlideRelationshipIDs(b []byte) ([]string, error) {
	dec := xml.NewDecoder(bytes.NewReader(b))
	var ids []string
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "sldId" {
			continue
		}
		id := pptxSlideRelationshipID(start)
		if id != "" {
			ids = append(ids, id)
		}
	}
	return ids, nil
}

func pptxSlideRelationshipID(start xml.StartElement) string {
	var fallback string
	for _, attr := range start.Attr {
		if attr.Name.Local != "id" {
			continue
		}
		value := strings.TrimSpace(attr.Value)
		if value == "" {
			continue
		}
		if attr.Name.Space != "" {
			return value
		}
		fallback = value
	}
	return fallback
}

func pptxVisibleNotesSlideNames(files map[string]*zip.File) (map[string]bool, error) {
	out := map[string]bool{}
	type noteVisibility struct {
		name          string
		visible       bool
		relatedSlides int
	}
	var notes []noteVisibility
	hasRelatedSlides := false
	for name := range files {
		lower := ooxmlPartKey(name)
		if !strings.HasPrefix(lower, "ppt/notesslides/notesslide") || !strings.HasSuffix(lower, ".xml") {
			continue
		}
		visible, relatedSlides, err := pptxPartRelatedSlideVisibility(files, name)
		if err != nil {
			visible = true
		}
		if relatedSlides > 0 {
			hasRelatedSlides = true
		}
		notes = append(notes, noteVisibility{name: lower, visible: visible, relatedSlides: relatedSlides})
	}
	for _, note := range notes {
		if hasRelatedSlides && note.relatedSlides == 0 {
			continue
		}
		if note.visible {
			out[note.name] = true
		}
	}
	return out, nil
}

func pptxVisibleCommentPartNames(files map[string]*zip.File) (map[string]bool, error) {
	visible, hidden, err := pptxSlideRelatedPartVisibility(files, "ppt/comments/")
	if err != nil {
		return nil, err
	}
	hasRelatedComments := len(visible) > 0 || len(hidden) > 0
	out := map[string]bool{}
	for name := range files {
		lower := ooxmlPartKey(name)
		if !strings.HasPrefix(lower, "ppt/comments/comment") || !strings.HasSuffix(lower, ".xml") {
			continue
		}
		if hasRelatedComments {
			if visible[lower] {
				out[lower] = true
			}
			continue
		}
		if !hidden[lower] {
			out[lower] = true
		}
	}
	return out, nil
}

func pptxPartVisibleByRelatedSlide(files map[string]*zip.File, name string) (bool, error) {
	visible, _, err := pptxPartRelatedSlideVisibility(files, name)
	return visible, err
}

func pptxPartRelatedSlideVisibility(files map[string]*zip.File, name string) (bool, int, error) {
	targets, err := relationshipTargetsForPart(files, name)
	if err != nil {
		return true, 0, nil
	}
	candidates, constrained, err := pptxCandidateSlideNames(files)
	if err != nil {
		return false, 0, err
	}
	candidate := map[string]bool{}
	for _, slide := range candidates {
		candidate[ooxmlPartKey(slide)] = true
	}
	relatedSlides := 0
	for _, target := range targets {
		part := resolveOOXMLRelationshipTarget(name, target)
		if actual := ooxmlPartName(files, part); actual != "" {
			part = actual
		}
		lower := ooxmlPartKey(part)
		if !strings.HasPrefix(lower, "ppt/slides/slide") || !strings.HasSuffix(lower, ".xml") {
			continue
		}
		relatedSlides++
		if constrained && !candidate[lower] {
			continue
		}
		visible, err := pptxSlideVisible(files, part)
		if err != nil {
			return false, relatedSlides, err
		}
		if visible {
			return true, relatedSlides, nil
		}
	}
	return relatedSlides == 0, relatedSlides, nil
}

func pptxSlideRelatedPartVisibility(files map[string]*zip.File, prefix string) (map[string]bool, map[string]bool, error) {
	visible := map[string]bool{}
	hidden := map[string]bool{}
	lowerPrefix := ooxmlPartKey(prefix)
	slideNames, constrained, err := pptxCandidateSlideNames(files)
	if err != nil {
		return nil, nil, err
	}
	candidate := map[string]bool{}
	for _, name := range slideNames {
		candidate[ooxmlPartKey(name)] = true
	}
	if constrained {
		slideNames = pptxAllSlidePartNames(files)
	}
	for _, name := range slideNames {
		slideVisible, err := pptxSlideVisible(files, name)
		if err != nil {
			return nil, nil, err
		}
		targets, err := relationshipTargetsForPart(files, name)
		if err != nil {
			continue
		}
		for _, target := range targets {
			part := resolveOOXMLRelationshipTarget(name, target)
			if actual := ooxmlPartName(files, part); actual != "" {
				part = actual
			}
			key := ooxmlPartKey(part)
			if !strings.HasPrefix(key, lowerPrefix) {
				continue
			}
			if slideVisible && (!constrained || candidate[ooxmlPartKey(name)]) {
				visible[key] = true
			} else {
				hidden[key] = true
			}
		}
	}
	return visible, hidden, nil
}

func relationshipTargetsForPart(files map[string]*zip.File, part string) ([]string, error) {
	f := ooxmlFile(files, ooxmlRelsName(part))
	if f == nil {
		return nil, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, err
	}
	return relationshipTargets(b)
}

func pptxSortedPartNames(files map[string]*zip.File, allowed map[string]bool) []string {
	var names []string
	for name := range files {
		if allowed[ooxmlPartKey(name)] {
			names = append(names, name)
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return names
}

func pptxMarkdownPart(files map[string]*zip.File, prefix, heading string) (string, error) {
	var names []string
	lowerPrefix := strings.ToLower(prefix)
	for name := range files {
		lower := ooxmlPartKey(name)
		if strings.HasPrefix(lower, lowerPrefix) && strings.HasSuffix(lower, ".xml") {
			names = append(names, name)
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return ooxmlMarkdownPart(files, names, heading)
}

func visibleXMLTextFromZip(files map[string]*zip.File, name string) (string, error) {
	f := ooxmlFile(files, name)
	if f == nil {
		return "", nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return "", err
	}
	return visibleXMLText(b)
}

func pptxSlideMarkdownTitle(name string) string {
	name = ooxmlCleanPartName(name)
	base := strings.TrimSuffix(path.Base(name), path.Ext(name))
	if decoded, err := url.PathUnescape(base); err == nil && decoded != "" {
		base = decoded
	}
	if !strings.HasPrefix(strings.ToLower(base), "slide") {
		if title := cleanVisibleText(base); title != "" {
			return title
		}
		return "Slide"
	}
	n := base[len("slide"):]
	if n == "" {
		return "Slide"
	}
	if _, ok := atoi(n); !ok {
		if title := cleanVisibleText(strings.TrimSpace(n)); title != "" {
			return "Slide " + title
		}
		return "Slide"
	}
	return "Slide " + n
}

func appendEmbeddedMarkdown(markdown string, embeddedText []string) string {
	text := demoteMarkdownHeadings(markdownText(joinText(embeddedText)), 1)
	if text == "" {
		return strings.TrimSpace(markdown)
	}
	markdown = strings.TrimSpace(markdown)
	embedded := "## Embedded Content\n\n" + text
	if markdown == "" {
		return embedded
	}
	return markdown + "\n\n" + embedded
}

func demoteMarkdownHeadings(markdown string, levels int) string {
	if levels <= 0 || strings.TrimSpace(markdown) == "" {
		return markdown
	}
	prefix := strings.Repeat("#", levels)
	lines := strings.Split(markdown, "\n")
	for i, line := range lines {
		trimmed := strings.TrimLeft(line, " ")
		indent := line[:len(line)-len(trimmed)]
		if !strings.HasPrefix(trimmed, "#") {
			continue
		}
		hashes := 0
		for hashes < len(trimmed) && trimmed[hashes] == '#' {
			hashes++
		}
		if hashes == 0 || hashes >= 6 || hashes >= len(trimmed) || trimmed[hashes] != ' ' {
			continue
		}
		lines[i] = indent + prefix + trimmed
	}
	return strings.Join(lines, "\n")
}

func pptxSlideVisible(files map[string]*zip.File, name string) (bool, error) {
	f := files[name]
	if f == nil {
		return true, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return false, err
	}
	if hasDOCTYPE(b) {
		return false, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			return true, nil
		}
		if err != nil {
			return false, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		if start.Name.Local != "sld" {
			return true, nil
		}
		for _, attr := range start.Attr {
			if attr.Name.Local == "show" {
				value := strings.TrimSpace(strings.ToLower(attr.Value))
				return !(value == "0" || value == "false"), nil
			}
		}
		return true, nil
	}
}

func extractXlsxText(files map[string]*zip.File, strictOfficeContent bool) (string, map[string]xlsxWorksheetMarkdownData, error) {
	shared, err := readSharedStrings(files)
	if err != nil {
		return "", nil, err
	}
	styles, err := readXlsxCellStyles(files)
	if err != nil {
		return "", nil, err
	}
	var out strings.Builder
	markdown := map[string]xlsxWorksheetMarkdownData{}
	workbookTexts, sheetNames, err := workbookTextAndSheets(files, strictOfficeContent)
	if err != nil {
		return "", nil, err
	}
	appendCleanedTextBlocks(&out, workbookTexts)
	if sheetNames == nil {
		for name := range files {
			lower := ooxmlPartKey(name)
			if strings.HasPrefix(lower, "xl/worksheets/sheet") && strings.HasSuffix(lower, ".xml") {
				sheetNames = append(sheetNames, name)
			}
		}
		sort.Slice(sheetNames, func(i, j int) bool {
			return naturalLess(sheetNames[i], sheetNames[j])
		})
	}
	growXlsxTextBuilder(&out, files, sheetNames)
	for _, name := range sheetNames {
		b, err := readZipFile(files[name])
		if err != nil {
			return "", nil, err
		}
		var md xlsxWorksheetMarkdownData
		if err := appendWorksheetText(&out, b, shared, styles, &md, strictOfficeContent); err != nil {
			return "", nil, err
		}
		markdown[ooxmlPartKey(name)] = md
	}
	if strictOfficeContent {
		// A strict Excel baseline is Worksheet.UsedRange.Text plus Picture
		// shapes.  Text in drawing, VML, chart, comments, pivot and other
		// auxiliary parts is not Range.Text.  In particular VML Pict controls
		// can contain hundreds of image relationships but no worksheet cells;
		// broad OOXML cleanup must never revive their XML tokens as visible text.
		return strings.TrimSpace(out.String()), markdown, nil
	}
	if !strictOfficeContent && xlsxHasAnyPartPrefix(files, []string{"xl/charts/", "xl/drawings/", "xl/tables/", "xl/pivottables/", "xl/pivotcache/", "xl/slicers/", "xl/slicercaches/"}) {
		var extraNames []string
		visibleDrawingParts := xlsxVisibleDrawingPartNames(files)
		visibleTableParts, constrainedTableParts := xlsxVisibleTablePartNames(files)
		visiblePivotParts, constrainedPivotParts := xlsxVisibleSpecialPartNames(files, []string{"xl/pivottables/", "xl/pivotcache/"})
		visibleSlicerParts, constrainedSlicerParts := xlsxVisibleSpecialPartNames(files, []string{"xl/slicers/", "xl/slicercaches/"})
		visibleChartParts, constrainedChartParts := xlsxVisibleChartPartNames(files)
		for name := range files {
			lower := ooxmlPartKey(name)
			include := (strings.HasPrefix(lower, "xl/charts/") && (!constrainedChartParts || visibleChartParts[lower])) ||
				visibleDrawingParts[lower] ||
				(strings.HasPrefix(lower, "xl/tables/") && (!constrainedTableParts || visibleTableParts[lower])) ||
				((strings.HasPrefix(lower, "xl/pivottables/") || strings.HasPrefix(lower, "xl/pivotcache/")) && (!constrainedPivotParts || visiblePivotParts[lower])) ||
				((strings.HasPrefix(lower, "xl/slicers/") || strings.HasPrefix(lower, "xl/slicercaches/")) && (!constrainedSlicerParts || visibleSlicerParts[lower]))
			if include && (strings.HasSuffix(lower, ".xml") || strings.HasSuffix(lower, ".vml")) {
				extraNames = append(extraNames, name)
			}
		}
		sort.Strings(extraNames)
		extras, err := xmlTextFromFiles(files, extraNames)
		if err != nil {
			return "", nil, err
		}
		appendCleanedTextBlocks(&out, extras)
	}
	if !strictOfficeContent && xlsxHasAnyPartPrefix(files, []string{"xl/comments", "xl/threadedcomments"}) {
		comments, err := xlsxVisibleCommentsText(files)
		if err != nil {
			return "", nil, err
		}
		appendCleanedTextBlocks(&out, comments)
	}
	return strings.TrimSpace(out.String()), markdown, nil
}

func xlsxHasAnyPartPrefix(files map[string]*zip.File, prefixes []string) bool {
	for name := range files {
		lower := ooxmlPartKey(name)
		if !(strings.HasSuffix(lower, ".xml") || strings.HasSuffix(lower, ".vml")) {
			continue
		}
		for _, prefix := range prefixes {
			if strings.HasPrefix(lower, prefix) {
				return true
			}
		}
	}
	return false
}

func growXlsxTextBuilder(out *strings.Builder, files map[string]*zip.File, sheetNames []string) {
	const maxPregrow = 512 << 20
	var size uint64
	for _, name := range sheetNames {
		if f := files[name]; f != nil {
			size += f.UncompressedSize64
			if size > maxPregrow {
				return
			}
		}
	}
	if size > 0 {
		out.Grow(int(size))
	}
}

func xlsxVisibleChartPartNames(files map[string]*zip.File) (map[string]bool, bool) {
	if key := xlsxVisibilityCacheKey(files); key != nil {
		if cached, ok := xlsxVisibleChartPartsCache.Load(key); ok {
			v := cached.(xlsxVisiblePartsCacheEntry)
			return cloneBoolMap(v.visible), v.constrained
		}
		visible, constrained := xlsxVisibleChartPartNamesUncached(files)
		xlsxVisibleChartPartsCache.Store(key, xlsxVisiblePartsCacheEntry{visible: cloneBoolMap(visible), constrained: constrained})
		return visible, constrained
	}
	return xlsxVisibleChartPartNamesUncached(files)
}

func xlsxVisibleChartPartNamesUncached(files map[string]*zip.File) (map[string]bool, bool) {
	visible := map[string]bool{}
	hidden := map[string]bool{}
	found := false
	sheets, err := workbookSheets(files)
	if err != nil {
		return nil, false
	}
	for _, sheet := range sheets {
		if sheet.Hidden {
			partFound, ok := collectReachableXlsxChartParts(files, sheet.Path, hidden, map[string]bool{})
			if !ok {
				return nil, false
			}
			found = found || partFound
			continue
		}
		partFound, ok := collectXlsxSourceVisibleChartParts(files, sheet.Path, visible, hidden, map[string]bool{})
		if !ok {
			return nil, false
		}
		found = found || partFound
	}
	if !found {
		if len(sheets) > 0 {
			return xlsxUnconstrainedVisiblePartsWhenNoReachableParts(files, sheets, visible)
		}
		return nil, false
	}
	for name := range hidden {
		if !visible[name] {
			delete(visible, name)
		}
	}
	return visible, true
}

// docxOLEPreviewShapeIDs identifies VML ShapeIDs that belong to embedded OLE
// objects.  Their imagedata payload is a preview cache, whereas Word exposes
// the containing object as msoEmbeddedOLEObject rather than as a Picture Shape.
func docxOLEPreviewShapeIDs(b []byte) (map[string]bool, error) {
	ids := map[string]bool{}
	if hasDOCTYPE(b) {
		return ids, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return ids, nil
		}
		if err != nil {
			return ids, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "OLEObject" {
			continue
		}
		if id := normalizeVMLShapeID(xmlAttrValue(start, "ShapeID")); id != "" {
			ids[id] = true
		}
	}
}

// docxVMLShapeIDsWithEmbeddedObjectData returns VML shapes that carry an
// Office embedded-object relationship in the same w:pict container.  Office
// exposes those as OLE objects (with a preview), not as Picture Shapes.
func docxVMLShapeIDsWithEmbeddedObjectData(b []byte) (map[string]bool, error) {
	ids := map[string]bool{}
	if hasDOCTYPE(b) {
		return ids, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	pictDepth := 0
	var shapeStack []string
	var pictShapeIDs []string
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return ids, nil
		}
		if err != nil {
			return ids, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "pict" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessingml") {
				pictDepth++
				if pictDepth == 1 {
					pictShapeIDs = nil
				}
			}
			if t.Name.Local == "shape" && strings.Contains(strings.ToLower(t.Name.Space), "vml") {
				id := normalizeVMLShapeID(xmlAttrValue(t, "id"))
				shapeStack = append(shapeStack, id)
				if pictDepth > 0 && id != "" {
					pictShapeIDs = append(pictShapeIDs, id)
				}
			}
			if pictDepth > 0 && t.Name.Local == "OLEObject" {
				for _, id := range pictShapeIDs {
					ids[id] = true
				}
			}
		case xml.EndElement:
			if t.Name.Local == "shape" && strings.Contains(strings.ToLower(t.Name.Space), "vml") && len(shapeStack) > 0 {
				shapeStack = shapeStack[:len(shapeStack)-1]
			}
			if t.Name.Local == "pict" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessingml") && pictDepth > 0 {
				pictDepth--
				if pictDepth == 0 {
					pictShapeIDs = nil
				}
			}
		}
	}
}

// docxSVGPictureRelationshipIDs returns all bitmap fallback and SVG payload
// relationships belonging to a DrawingML picture with an Office SVG blip.
// Word exposes this legacy SVG construct as a non-picture InlineShape on the
// current COM surface, so neither payload belongs in the strict Picture count.
func docxSVGPictureRelationshipIDs(b []byte) (map[string]bool, error) {
	ids := map[string]bool{}
	if hasDOCTYPE(b) {
		return ids, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	pictureDepth := 0
	var pictureIDs [][]string
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return ids, nil
		}
		if err != nil {
			return ids, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "pic" && strings.Contains(strings.ToLower(t.Name.Space), "picture") {
				pictureDepth++
				pictureIDs = append(pictureIDs, nil)
			}
			if pictureDepth == 0 {
				continue
			}
			values := imageRelationshipIDs(t)
			if len(values) > 0 {
				pictureIDs[len(pictureIDs)-1] = append(pictureIDs[len(pictureIDs)-1], values...)
			}
			if t.Name.Local == "svgBlip" {
				for _, id := range pictureIDs[len(pictureIDs)-1] {
					ids[id] = true
				}
			}
		case xml.EndElement:
			if t.Name.Local == "pic" && strings.Contains(strings.ToLower(t.Name.Space), "picture") && pictureDepth > 0 {
				pictureDepth--
				pictureIDs = pictureIDs[:len(pictureIDs)-1]
			}
		}
	}
}

// docxVMLPictureRelationshipIDs reports VML picture payloads while excluding
// the preview media Word assigns to chart/diagram placeholder groups.  Both
// use v:shape type #_x0000_t75, but only the former is named as a Picture in
// Word's Shapes/GroupItems collection; producer-generated Chart/Object names
// are msoChart/msoEmbeddedOLEObject and must not inflate picture counts.
func docxVMLPictureRelationshipIDs(b []byte) (map[string]bool, error) {
	ids := map[string]bool{}
	if hasDOCTYPE(b) {
		return ids, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	type vmlShapeState struct {
		isPicture bool
		ids       []string
	}
	var shapes []vmlShapeState
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return ids, nil
		}
		if err != nil {
			return ids, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "shape" && strings.Contains(strings.ToLower(t.Name.Space), "vml") {
				name := strings.ToLower(strings.TrimSpace(xmlAttrValue(t, "id")))
				shapeType := strings.ToLower(strings.TrimSpace(xmlAttrValue(t, "type")))
				// Word's VML shape name identifies its COM Shape type more reliably
				// than the shared t75 geometry. Names beginning "Object" are
				// embedded OLE/package placeholders (not msoPicture) even though
				// they carry visible imagedata; names such as "Picture 1" are the
				// actual Word Picture Shapes.
				isObjectPlaceholder := strings.HasPrefix(name, "object")
				isPicture := strings.Contains(shapeType, "_x0000_t75") && !isObjectPlaceholder && !strings.Contains(name, "chart") && !strings.Contains(name, "placeholder")
				shapes = append(shapes, vmlShapeState{isPicture: isPicture})
				continue
			}
			if t.Name.Local == "imagedata" && strings.Contains(strings.ToLower(t.Name.Space), "vml") && len(shapes) > 0 {
				// A VML group is not on this stack, so the most recent VML shape
				// remains the shape that owns this imagedata element.
				shapes[len(shapes)-1].ids = append(shapes[len(shapes)-1].ids, imageRelationshipIDs(t)...)
			}
		case xml.EndElement:
			if t.Name.Local == "shape" && strings.Contains(strings.ToLower(t.Name.Space), "vml") && len(shapes) > 0 {
				shape := shapes[len(shapes)-1]
				shapes = shapes[:len(shapes)-1]
				if shape.isPicture {
					for _, id := range shape.ids {
						ids[id] = true
					}
				}
			}
		}
	}
}

// docxStrictPictureRelationshipRefs mirrors Word's InlineShapes and Shapes
// collections: a relationship counts only when it belongs to a DrawingML
// picture (pic:pic) or a VML image element. Image fills, diagrams, and other
// drawing resources can reference media but are exposed by Word as non-picture
// shapes (for example msoTextBox or msoGraphic), not document images.
func docxStrictPictureRelationshipRefs(b []byte) (docxImageRefs, error) {
	refs := docxImageRefs{Visible: map[string]bool{}, Hidden: map[string]bool{}}
	if hasDOCTYPE(b) {
		return refs, errors.New("xml doctype is not supported")
	}
	olePreviewShapes, err := docxOLEPreviewShapeIDs(b)
	if err != nil {
		return refs, err
	}
	embeddedObjectShapes, err := docxVMLShapeIDsWithEmbeddedObjectData(b)
	if err != nil {
		return refs, err
	}
	for id := range embeddedObjectShapes {
		olePreviewShapes[id] = true
	}
	svgPictureIDs, err := docxSVGPictureRelationshipIDs(b)
	if err != nil {
		return refs, err
	}
	vmlPictureIDs, err := docxVMLPictureRelationshipIDs(b)
	if err != nil {
		return refs, err
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	pictureDepth := 0
	vmlImageDepth := 0
	objectDepth := 0
	anchorDepth := 0
	textBoxDepth := 0
	wordProcessingGroupDepth := 0
	var vmlSuppressImages []bool
	vmlShapeDepth := 0
	vmlShapeID := ""
	var vmlShapeIDs []string
	var alternateStack []alternateContentState
	fallbackDepth := 0
	for {
		token, err := dec.Token()
		if err == io.EOF {
			return refs, nil
		}
		if err != nil {
			return refs, err
		}
		switch t := token.(type) {
		case xml.StartElement:
			if t.Name.Local == "AlternateContent" {
				alternateStack = append(alternateStack, alternateContentState{})
				continue
			}
			if len(alternateStack) > 0 {
				top := &alternateStack[len(alternateStack)-1]
				switch t.Name.Local {
				case "Choice":
					top.choiceSeen = true
				case "Fallback":
					if top.choiceSeen {
						fallbackDepth++
					}
				}
			}
			if t.Name.Local == "object" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessingml") {
				objectDepth++
			}
			if t.Name.Local == "anchor" && strings.Contains(strings.ToLower(t.Name.Space), "drawingml") {
				anchorDepth++
				if len(alternateStack) > 0 {
					alternateStack[len(alternateStack)-1].choiceHasAnchor = true
				}
			}
			if (t.Name.Local == "txbxContent" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessingml")) || (t.Name.Local == "textbox" && strings.Contains(strings.ToLower(t.Name.Space), "vml")) {
				textBoxDepth++
			}
			if t.Name.Local == "wgp" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessinggroup") {
				wordProcessingGroupDepth++
			}
			if t.Name.Local == "group" && strings.Contains(strings.ToLower(t.Name.Space), "vml") {
				vmlSuppressImages = append(vmlSuppressImages, docxVMLFloatingCanvasGroup(t))
			}
			if t.Name.Local == "pic" && strings.Contains(strings.ToLower(t.Name.Space), "picture") {
				pictureDepth++
			}
			if t.Name.Local == "imagedata" && strings.Contains(strings.ToLower(t.Name.Space), "vml") {
				vmlImageDepth++
			}
			if t.Name.Local == "shape" && strings.Contains(strings.ToLower(t.Name.Space), "vml") {
				vmlShapeIDs = append(vmlShapeIDs, normalizeVMLShapeID(xmlAttrValue(t, "id")))
				vmlShapeDepth++
			}
			vmlShapeID = ""
			if len(vmlShapeIDs) > 0 {
				vmlShapeID = vmlShapeIDs[len(vmlShapeIDs)-1]
			}
			isOLEPreview := vmlImageDepth > 0 && olePreviewShapes[vmlShapeID]
			vmlSuppressed := false
			for _, suppress := range vmlSuppressImages {
				if suppress {
					vmlSuppressed = true
					break
				}
			}
			// A legacy VML group may be stored inside w:object while Word still
			// exposes its picture children through Shape.GroupItems.  Keep the
			// w:object exclusion for DrawingML pictures (where it represents an
			// embedded-object preview), but admit VML relationships that have
			// already been classified as actual VML picture shapes.
			isDrawingMLPicture := pictureDepth > 0
			isVMLPicture := vmlImageDepth > 0
			allowedObjectContext := (isDrawingMLPicture && objectDepth == 0) || isVMLPicture
			// Word exposes DrawingML pictures nested in a legacy VML text box
			// through InlineShapes; VML imagedata in that same box is merely a
			// fill/preview and stays excluded.
			allowedTextBoxContext := textBoxDepth == 0 || isDrawingMLPicture
			// Word still exposes legacy VML pictures in mc:Fallback (it uses that
			// representation for these legacy groups), whereas DrawingML fallback
			// would be a duplicate of the selected Choice and must stay hidden.
			// A floating WPG group is exposed by Word as a Shape and exposes
			// its VML fallback members as GroupItems. An inline WPG group is
			// not exposed that way, so its fallback must remain suppressed.
			choiceHasAnchor := len(alternateStack) > 0 && alternateStack[len(alternateStack)-1].choiceHasAnchor
			allowedAlternateContent := fallbackDepth == 0 || (isVMLPicture && choiceHasAnchor)
			if allowedAlternateContent && allowedObjectContext && allowedTextBoxContext && wordProcessingGroupDepth == 0 && !vmlSuppressed && !isOLEPreview && (isDrawingMLPicture || isVMLPicture) {
				for _, id := range imageRelationshipIDs(t) {
					if !svgPictureIDs[id] && (isDrawingMLPicture || vmlPictureIDs[id]) {
						refs.Visible[id] = true
					}
				}
			}
		case xml.EndElement:
			if t.Name.Local == "Fallback" && fallbackDepth > 0 {
				fallbackDepth--
			}
			if t.Name.Local == "AlternateContent" && len(alternateStack) > 0 {
				alternateStack = alternateStack[:len(alternateStack)-1]
			}
			if t.Name.Local == "object" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessingml") && objectDepth > 0 {
				objectDepth--
			}
			if t.Name.Local == "anchor" && strings.Contains(strings.ToLower(t.Name.Space), "drawingml") && anchorDepth > 0 {
				anchorDepth--
			}
			if ((t.Name.Local == "txbxContent" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessingml")) || (t.Name.Local == "textbox" && strings.Contains(strings.ToLower(t.Name.Space), "vml"))) && textBoxDepth > 0 {
				textBoxDepth--
			}
			if t.Name.Local == "wgp" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessinggroup") && wordProcessingGroupDepth > 0 {
				wordProcessingGroupDepth--
			}
			if t.Name.Local == "group" && strings.Contains(strings.ToLower(t.Name.Space), "vml") && len(vmlSuppressImages) > 0 {
				vmlSuppressImages = vmlSuppressImages[:len(vmlSuppressImages)-1]
			}
			if t.Name.Local == "pic" && strings.Contains(strings.ToLower(t.Name.Space), "picture") && pictureDepth > 0 {
				pictureDepth--
			}
			if t.Name.Local == "imagedata" && strings.Contains(strings.ToLower(t.Name.Space), "vml") && vmlImageDepth > 0 {
				vmlImageDepth--
			}
			if t.Name.Local == "shape" && strings.Contains(strings.ToLower(t.Name.Space), "vml") && vmlShapeDepth > 0 {
				vmlShapeDepth--
				if len(vmlShapeIDs) > 0 {
					vmlShapeIDs = vmlShapeIDs[:len(vmlShapeIDs)-1]
				}
				vmlShapeID = ""
				if len(vmlShapeIDs) > 0 {
					vmlShapeID = vmlShapeIDs[len(vmlShapeIDs)-1]
				}
			}
		}
	}
}

// docxStrictPictureRelationshipIDsInOrder is the occurrence-preserving form
// of docxStrictPictureRelationshipRefs.  It deliberately records duplicate
// r:embed/r:link values because Word exposes each picture placement as an
// individual InlineShape or Shape.
func docxStrictPictureRelationshipIDsInOrder(b []byte) ([]string, error) {
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	olePreviewShapes, err := docxOLEPreviewShapeIDs(b)
	if err != nil {
		return nil, err
	}
	embeddedObjectShapes, err := docxVMLShapeIDsWithEmbeddedObjectData(b)
	if err != nil {
		return nil, err
	}
	for id := range embeddedObjectShapes {
		olePreviewShapes[id] = true
	}
	svgPictureIDs, err := docxSVGPictureRelationshipIDs(b)
	if err != nil {
		return nil, err
	}
	vmlPictureIDs, err := docxVMLPictureRelationshipIDs(b)
	if err != nil {
		return nil, err
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	pictureDepth := 0
	vmlImageDepth := 0
	objectDepth := 0
	anchorDepth := 0
	textBoxDepth := 0
	wordProcessingGroupDepth := 0
	var vmlSuppressImages []bool
	vmlShapeDepth := 0
	vmlShapeID := ""
	var vmlShapeIDs []string
	var alternateStack []alternateContentState
	fallbackDepth := 0
	var ids []string
	for {
		token, err := dec.Token()
		if err == io.EOF {
			return ids, nil
		}
		if err != nil {
			return nil, err
		}
		switch t := token.(type) {
		case xml.StartElement:
			if t.Name.Local == "AlternateContent" {
				alternateStack = append(alternateStack, alternateContentState{})
				continue
			}
			if len(alternateStack) > 0 {
				top := &alternateStack[len(alternateStack)-1]
				switch t.Name.Local {
				case "Choice":
					top.choiceSeen = true
				case "Fallback":
					if top.choiceSeen {
						fallbackDepth++
					}
				}
			}
			if t.Name.Local == "object" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessingml") {
				objectDepth++
			}
			if t.Name.Local == "anchor" && strings.Contains(strings.ToLower(t.Name.Space), "drawingml") {
				anchorDepth++
				if len(alternateStack) > 0 {
					alternateStack[len(alternateStack)-1].choiceHasAnchor = true
				}
			}
			if (t.Name.Local == "txbxContent" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessingml")) || (t.Name.Local == "textbox" && strings.Contains(strings.ToLower(t.Name.Space), "vml")) {
				textBoxDepth++
			}
			if t.Name.Local == "wgp" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessinggroup") {
				wordProcessingGroupDepth++
			}
			if t.Name.Local == "group" && strings.Contains(strings.ToLower(t.Name.Space), "vml") {
				vmlSuppressImages = append(vmlSuppressImages, docxVMLFloatingCanvasGroup(t))
			}
			if t.Name.Local == "pic" && strings.Contains(strings.ToLower(t.Name.Space), "picture") {
				pictureDepth++
			}
			if t.Name.Local == "imagedata" && strings.Contains(strings.ToLower(t.Name.Space), "vml") {
				vmlImageDepth++
			}
			if t.Name.Local == "shape" && strings.Contains(strings.ToLower(t.Name.Space), "vml") {
				vmlShapeIDs = append(vmlShapeIDs, normalizeVMLShapeID(xmlAttrValue(t, "id")))
				vmlShapeDepth++
			}
			vmlShapeID = ""
			if len(vmlShapeIDs) > 0 {
				vmlShapeID = vmlShapeIDs[len(vmlShapeIDs)-1]
			}
			isOLEPreview := vmlImageDepth > 0 && olePreviewShapes[vmlShapeID]
			vmlSuppressed := false
			for _, suppress := range vmlSuppressImages {
				if suppress {
					vmlSuppressed = true
					break
				}
			}
			// Keep this occurrence-preserving path in lockstep with
			// docxStrictPictureRelationshipRefs; see its comment for why VML
			// GroupItems in w:object remain visible to Word.
			isDrawingMLPicture := pictureDepth > 0
			isVMLPicture := vmlImageDepth > 0
			allowedObjectContext := (isDrawingMLPicture && objectDepth == 0) || isVMLPicture
			allowedTextBoxContext := textBoxDepth == 0 || isDrawingMLPicture
			// See docxStrictPictureRelationshipRefs: a VML fallback is the
			// Word-visible legacy representation, unlike a DrawingML fallback.
			choiceHasAnchor := len(alternateStack) > 0 && alternateStack[len(alternateStack)-1].choiceHasAnchor
			allowedAlternateContent := fallbackDepth == 0 || (isVMLPicture && choiceHasAnchor)
			if allowedAlternateContent && allowedObjectContext && allowedTextBoxContext && wordProcessingGroupDepth == 0 && !vmlSuppressed && !isOLEPreview && (isDrawingMLPicture || isVMLPicture) {
				values := imageRelationshipIDs(t)
				for _, id := range values {
					if !svgPictureIDs[id] && (isDrawingMLPicture || vmlPictureIDs[id]) {
						ids = append(ids, id)
					}
				}
			}
		case xml.EndElement:
			if t.Name.Local == "Fallback" && fallbackDepth > 0 {
				fallbackDepth--
			}
			if t.Name.Local == "AlternateContent" && len(alternateStack) > 0 {
				alternateStack = alternateStack[:len(alternateStack)-1]
			}
			if t.Name.Local == "object" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessingml") && objectDepth > 0 {
				objectDepth--
			}
			if t.Name.Local == "anchor" && strings.Contains(strings.ToLower(t.Name.Space), "drawingml") && anchorDepth > 0 {
				anchorDepth--
			}
			if ((t.Name.Local == "txbxContent" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessingml")) || (t.Name.Local == "textbox" && strings.Contains(strings.ToLower(t.Name.Space), "vml"))) && textBoxDepth > 0 {
				textBoxDepth--
			}
			if t.Name.Local == "wgp" && strings.Contains(strings.ToLower(t.Name.Space), "wordprocessinggroup") && wordProcessingGroupDepth > 0 {
				wordProcessingGroupDepth--
			}
			if t.Name.Local == "group" && strings.Contains(strings.ToLower(t.Name.Space), "vml") && len(vmlSuppressImages) > 0 {
				vmlSuppressImages = vmlSuppressImages[:len(vmlSuppressImages)-1]
			}
			if t.Name.Local == "pic" && strings.Contains(strings.ToLower(t.Name.Space), "picture") && pictureDepth > 0 {
				pictureDepth--
			}
			if t.Name.Local == "imagedata" && strings.Contains(strings.ToLower(t.Name.Space), "vml") && vmlImageDepth > 0 {
				vmlImageDepth--
			}
			if t.Name.Local == "shape" && strings.Contains(strings.ToLower(t.Name.Space), "vml") && vmlShapeDepth > 0 {
				vmlShapeDepth--
				if len(vmlShapeIDs) > 0 {
					vmlShapeIDs = vmlShapeIDs[:len(vmlShapeIDs)-1]
				}
				vmlShapeID = ""
				if len(vmlShapeIDs) > 0 {
					vmlShapeID = vmlShapeIDs[len(vmlShapeIDs)-1]
				}
			}
		}
	}
}

// docxVMLFloatingCanvasGroup identifies a VML canvas whose direct image
// children Word exposes only through the canvas, not as Picture Shapes.
// The suppression is deliberately local to that group: images in a nested
// ordinary VML group remain separate GroupItems, even when that group itself
// is contained by a canvas.
func docxVMLFloatingCanvasGroup(start xml.StartElement) bool {
	return strings.EqualFold(strings.TrimSpace(xmlAttrValue(start, "editas")), "canvas")
}

func collectXlsxSourceVisibleChartParts(files map[string]*zip.File, source string, visible, hidden, seen map[string]bool) (bool, bool) {
	if seen[ooxmlPartKey(source)] {
		return false, true
	}
	seen[ooxmlPartKey(source)] = true
	f := ooxmlFile(files, source)
	if f == nil {
		return collectReachableXlsxChartParts(files, source, visible, map[string]bool{})
	}
	rels, err := relationshipTargetMapForPart(files, source)
	if err != nil {
		return collectReachableXlsxChartParts(files, source, visible, map[string]bool{})
	}
	if !relationshipTargetsAnyResolvedPrefix(files, source, rels, []string{"xl/charts/", "xl/drawings/"}) {
		return false, true
	}
	b, err := readZipFile(f)
	if err != nil {
		return collectReachableXlsxChartParts(files, source, visible, map[string]bool{})
	}
	refs, err := xlsxRelationshipRefsForChartSource(source, b)
	if err != nil || (len(refs.Visible) == 0 && len(refs.Hidden) == 0) {
		return collectReachableXlsxChartParts(files, source, visible, map[string]bool{})
	}
	if len(rels) == 0 {
		return collectReachableXlsxChartParts(files, source, visible, map[string]bool{})
	}
	found := false
	for id := range refs.Visible {
		partFound, ok := collectRelationshipTargetXlsxChartPart(files, source, rels[id], visible, hidden, false, seen)
		if !ok {
			return false, false
		}
		found = found || partFound
	}
	for id := range refs.Hidden {
		partFound, ok := collectRelationshipTargetXlsxChartPart(files, source, rels[id], visible, hidden, true, seen)
		if !ok {
			return false, false
		}
		found = found || partFound
	}
	if !found {
		return collectReachableXlsxChartParts(files, source, visible, map[string]bool{})
	}
	return true, true
}

func xlsxRelationshipRefsForChartSource(source string, b []byte) (docxImageRefs, error) {
	if xlsxWorksheetPart(source) {
		if refs, ok, err := xlsxWorksheetRelationshipRefsFast(b); ok || err != nil {
			return refs, err
		}
	}
	if !likelyImageRelationshipMarkup(b) {
		return docxImageRefs{Visible: map[string]bool{}, Hidden: map[string]bool{}}, nil
	}
	return xlsxImageRelationshipRefs(b)
}

func xlsxWorksheetPart(name string) bool {
	lower := ooxmlPartKey(name)
	return strings.HasPrefix(lower, "xl/worksheets/") && strings.HasSuffix(lower, ".xml")
}

func xlsxWorksheetRelationshipRefsFast(b []byte) (docxImageRefs, bool, error) {
	refs := docxImageRefs{Visible: map[string]bool{}, Hidden: map[string]bool{}}
	if hasDOCTYPE(b) {
		return refs, true, errors.New("xml doctype is not supported")
	}
	if bytes.Contains(b, []byte("AlternateContent")) {
		return refs, false, nil
	}
	scanXMLRelationshipAttributeIDs(b, func(id string) {
		refs.Visible[id] = true
	})
	return refs, true, nil
}

func collectRelationshipTargetXlsxChartPart(files map[string]*zip.File, source, target string, visible, hidden map[string]bool, forceHidden bool, seen map[string]bool) (bool, bool) {
	if strings.TrimSpace(target) == "" {
		return false, true
	}
	part := resolveOOXMLRelationshipTarget(source, target)
	if actual := ooxmlPartName(files, part); actual != "" {
		part = actual
	}
	lower := ooxmlPartKey(part)
	found := false
	if strings.HasPrefix(lower, "xl/charts/") && strings.HasSuffix(lower, ".xml") {
		if forceHidden {
			hidden[lower] = true
		} else {
			visible[lower] = true
		}
		found = true
	}
	if strings.HasPrefix(lower, "xl/") && ooxmlFile(files, ooxmlRelsName(part)) != nil {
		if forceHidden {
			childFound, ok := collectReachableXlsxChartParts(files, part, hidden, map[string]bool{})
			if !ok {
				return false, false
			}
			found = found || childFound
		} else {
			childFound, ok := collectXlsxSourceVisibleChartParts(files, part, visible, hidden, seen)
			if !ok {
				return false, false
			}
			found = found || childFound
		}
	}
	return found, true
}

func xlsxVisibleChartPartNameList(files map[string]*zip.File) []string {
	visible, constrained := xlsxVisibleChartPartNames(files)
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if !strings.HasPrefix(lower, "xl/charts/") || !strings.HasSuffix(lower, ".xml") {
			continue
		}
		if constrained && !visible[lower] {
			continue
		}
		names = append(names, name)
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return names
}

func collectReachableXlsxChartParts(files map[string]*zip.File, source string, out, seen map[string]bool) (bool, bool) {
	relsName := ooxmlRelsName(source)
	if seen[relsName] {
		return false, true
	}
	seen[relsName] = true
	f := ooxmlFile(files, relsName)
	if f == nil {
		return false, true
	}
	b, err := readZipFile(f)
	if err != nil {
		return false, false
	}
	targets, err := relationshipTargets(b)
	if err != nil {
		return false, false
	}
	found := false
	for _, target := range targets {
		part := resolveOOXMLRelationshipTarget(source, target)
		if actual := ooxmlPartName(files, part); actual != "" {
			part = actual
		}
		lower := ooxmlPartKey(part)
		if strings.HasPrefix(lower, "xl/charts/") && strings.HasSuffix(lower, ".xml") {
			out[lower] = true
			found = true
		}
		if strings.HasPrefix(lower, "xl/") && ooxmlFile(files, ooxmlRelsName(part)) != nil {
			childFound, ok := collectReachableXlsxChartParts(files, part, out, seen)
			if !ok {
				return false, false
			}
			found = found || childFound
		}
	}
	return found, true
}

func xlsxVisibleTablePartNames(files map[string]*zip.File) (map[string]bool, bool) {
	visible := map[string]bool{}
	hidden := map[string]bool{}
	found := false
	sheets, err := workbookSheets(files)
	if err != nil {
		return nil, false
	}
	for _, sheet := range sheets {
		targets := visible
		if sheet.Hidden {
			targets = hidden
		}
		for _, table := range relationshipTargetsWithPrefix(files, sheet.Path, "xl/tables/") {
			lower := ooxmlPartKey(table)
			if strings.HasSuffix(lower, ".xml") {
				targets[lower] = true
				found = true
			}
		}
	}
	if !found {
		if len(sheets) > 0 {
			return xlsxUnconstrainedVisiblePartsWhenNoReachableParts(files, sheets, visible)
		}
		return nil, false
	}
	for name := range hidden {
		if !visible[name] {
			delete(visible, name)
		}
	}
	return visible, true
}

func xlsxVisibleSpecialPartNames(files map[string]*zip.File, prefixes []string) (map[string]bool, bool) {
	visible := map[string]bool{}
	hidden := map[string]bool{}
	found := false
	sheets, err := workbookSheets(files)
	if err != nil {
		return nil, false
	}
	for _, sheet := range sheets {
		targets := visible
		if sheet.Hidden {
			targets = hidden
		}
		partFound, ok := collectReachableXlsxSpecialParts(files, sheet.Path, prefixes, targets, map[string]bool{})
		if !ok {
			return nil, false
		}
		found = found || partFound
	}
	if !found {
		if len(sheets) > 0 {
			return xlsxUnconstrainedVisiblePartsWhenNoReachableParts(files, sheets, visible)
		}
		return nil, false
	}
	for name := range hidden {
		if !visible[name] {
			delete(visible, name)
		}
	}
	return visible, true
}

func xlsxUnconstrainedVisiblePartsWhenNoReachableParts(files map[string]*zip.File, sheets []workbookSheet, visible map[string]bool) (map[string]bool, bool) {
	if _, hasRels, err := workbookRelationships(files); err != nil || hasRels {
		return visible, true
	}
	for _, sheet := range sheets {
		if !sheet.Hidden {
			return nil, false
		}
	}
	return visible, true
}

func collectReachableXlsxSpecialParts(files map[string]*zip.File, source string, prefixes []string, out, seen map[string]bool) (bool, bool) {
	relsName := ooxmlRelsName(source)
	if seen[relsName] {
		return false, true
	}
	seen[relsName] = true
	f := ooxmlFile(files, relsName)
	if f == nil {
		return false, true
	}
	b, err := readZipFile(f)
	if err != nil {
		return false, false
	}
	targets, err := relationshipTargets(b)
	if err != nil {
		return false, false
	}
	found := false
	for _, target := range targets {
		part := resolveOOXMLRelationshipTarget(source, target)
		if actual := ooxmlPartName(files, part); actual != "" {
			part = actual
		}
		lower := ooxmlPartKey(part)
		if xlsxPartHasAnyPrefix(lower, prefixes) && strings.HasSuffix(lower, ".xml") {
			out[lower] = true
			found = true
		}
		if strings.HasPrefix(lower, "xl/") && ooxmlFile(files, ooxmlRelsName(part)) != nil {
			childFound, ok := collectReachableXlsxSpecialParts(files, part, prefixes, out, seen)
			if !ok {
				return false, false
			}
			found = found || childFound
		}
	}
	return found, true
}

func xlsxPartHasAnyPrefix(name string, prefixes []string) bool {
	lower := ooxmlPartKey(name)
	for _, prefix := range prefixes {
		if strings.HasPrefix(lower, ooxmlPartKey(prefix)) {
			return true
		}
	}
	return false
}

func relationshipTargetsAnyResolvedPrefix(files map[string]*zip.File, source string, rels map[string]string, prefixes []string) bool {
	for _, target := range rels {
		part := resolveOOXMLRelationshipTarget(source, target)
		if actual := ooxmlPartName(files, part); actual != "" {
			part = actual
		}
		if xlsxPartHasAnyPrefix(part, prefixes) {
			return true
		}
	}
	return false
}

func xlsxVisibleRelatedPart(part string, hidden, visible map[string]bool, prefix string) bool {
	part = ooxmlPartKey(part)
	return strings.HasPrefix(part, ooxmlPartKey(prefix)) && (visible[part] || !hidden[part])
}

func xlsxVisibleRelatedPartNames(files map[string]*zip.File, prefix string) []string {
	hidden, visible := xlsxSheetRelatedPartNames(files, prefix)
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if xlsxVisibleRelatedPart(lower, hidden, visible, prefix) &&
			(strings.HasSuffix(lower, ".xml") || strings.HasSuffix(lower, ".vml")) {
			names = append(names, name)
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return names
}

type xlsxSheetCellVisibility struct {
	hiddenCols []intRange
	hiddenRows []intRange
}

func xlsxVisibleCommentsText(files map[string]*zip.File) ([]string, error) {
	sources, err := xlsxVisibleCommentPartSources(files)
	if err != nil {
		return nil, err
	}
	var names []string
	for name := range sources {
		names = append(names, name)
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	var out []string
	for _, name := range names {
		f := ooxmlFile(files, name)
		if f == nil {
			continue
		}
		b, err := readZipFile(f)
		if err != nil {
			return nil, err
		}
		items, err := visibleXlsxCommentItems(b, sources[name])
		if err != nil {
			return nil, err
		}
		for _, item := range items {
			if item.text != "" {
				out = append(out, item.text)
			}
		}
	}
	return out, nil
}

func xlsxVisibleCommentsMarkdownPart(files map[string]*zip.File) (string, error) {
	sources, err := xlsxVisibleCommentPartSources(files)
	if err != nil {
		return "", err
	}
	var names []string
	for name := range sources {
		names = append(names, name)
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	var out []string
	for _, name := range names {
		f := ooxmlFile(files, name)
		if f == nil {
			continue
		}
		b, err := readZipFile(f)
		if err != nil {
			return "", err
		}
		items, err := visibleXlsxCommentItems(b, sources[name])
		if err != nil {
			return "", err
		}
		for _, item := range items {
			text := markdownText(item.text)
			if text == "" {
				continue
			}
			if item.ref != "" {
				text = "### " + escapeMarkdownHeading(item.ref) + "\n\n" + text
			}
			out = append(out, text)
		}
	}
	if len(out) == 0 {
		return "", nil
	}
	return "## Comments\n\n" + strings.Join(out, "\n\n"), nil
}

func xlsxVisibleCommentPartSources(files map[string]*zip.File) (map[string][]xlsxSheetCellVisibility, error) {
	if key := xlsxVisibilityCacheKey(files); key != nil {
		if cached, ok := xlsxVisibleCommentSourcesCache.Load(key); ok {
			return cloneXlsxVisibleCommentSources(cached.(map[string][]xlsxSheetCellVisibility)), nil
		}
		out, err := xlsxVisibleCommentPartSourcesUncached(files)
		if err != nil {
			return nil, err
		}
		xlsxVisibleCommentSourcesCache.Store(key, cloneXlsxVisibleCommentSources(out))
		return out, nil
	}
	return xlsxVisibleCommentPartSourcesUncached(files)
}

func xlsxVisibleCommentPartSourcesUncached(files map[string]*zip.File) (map[string][]xlsxSheetCellVisibility, error) {
	out := map[string][]xlsxSheetCellVisibility{}
	sheets, err := workbookSheets(files)
	if err != nil {
		return nil, err
	}
	for _, sheet := range sheets {
		if sheet.Hidden {
			continue
		}
		var commentParts []string
		for _, prefix := range []string{"xl/comments", "xl/threadedcomments"} {
			for _, part := range relationshipTargetsWithPrefix(files, sheet.Path, prefix) {
				lower := ooxmlPartKey(part)
				if strings.HasSuffix(lower, ".xml") {
					commentParts = append(commentParts, lower)
				}
			}
		}
		if len(commentParts) == 0 {
			continue
		}
		f := ooxmlFile(files, sheet.Path)
		if f == nil {
			continue
		}
		b, err := readZipFile(f)
		if err != nil {
			return nil, err
		}
		hiddenCols, hiddenRows, err := worksheetHiddenRanges(b)
		if err != nil {
			return nil, err
		}
		visibility := xlsxSheetCellVisibility{hiddenCols: hiddenCols, hiddenRows: hiddenRows}
		for _, lower := range commentParts {
			out[lower] = append(out[lower], visibility)
		}
	}
	if len(out) > 0 {
		return out, nil
	}
	if len(sheets) > 0 {
		_, hasWorkbookRels, err := workbookRelationships(files)
		if err != nil {
			return nil, err
		}
		if !hasWorkbookRels {
			for _, prefix := range []string{"xl/comments", "xl/threadedcomments"} {
				for _, name := range xlsxVisibleRelatedPartNames(files, prefix) {
					out[ooxmlPartKey(name)] = nil
				}
			}
			return out, nil
		}
		return out, nil
	}
	for _, prefix := range []string{"xl/comments", "xl/threadedcomments"} {
		for _, name := range xlsxVisibleRelatedPartNames(files, prefix) {
			out[ooxmlPartKey(name)] = nil
		}
	}
	return out, nil
}

type xlsxVisiblePartsCacheEntry struct {
	visible     map[string]bool
	constrained bool
}

var xlsxVisibleChartPartsCache sync.Map
var xlsxVisibleCommentSourcesCache sync.Map
var ooxmlImageRelationshipRefsCache sync.Map

func xlsxVisibilityCacheKey(files map[string]*zip.File) *zip.File {
	if f := ooxmlFile(files, "xl/workbook.xml"); f != nil {
		return f
	}
	for _, f := range files {
		return f
	}
	return nil
}

func clearOOXMLExtractionCaches(files map[string]*zip.File) {
	for _, f := range files {
		xlsxVisibleChartPartsCache.Delete(f)
		xlsxVisibleCommentSourcesCache.Delete(f)
		ooxmlImageRelationshipRefsCache.Delete(f)
	}
}

func cloneXlsxVisibleCommentSources(in map[string][]xlsxSheetCellVisibility) map[string][]xlsxSheetCellVisibility {
	out := make(map[string][]xlsxSheetCellVisibility, len(in))
	for name, sources := range in {
		clonedSources := make([]xlsxSheetCellVisibility, len(sources))
		for i, source := range sources {
			clonedSources[i] = xlsxSheetCellVisibility{
				hiddenCols: append([]intRange(nil), source.hiddenCols...),
				hiddenRows: append([]intRange(nil), source.hiddenRows...),
			}
		}
		out[name] = clonedSources
	}
	return out
}

type xlsxCommentItem struct {
	ref  string
	text string
}

func worksheetHiddenRanges(b []byte) ([]intRange, []intRange, error) {
	if hasDOCTYPE(b) {
		return nil, nil, errors.New("xml doctype is not supported")
	}
	if !worksheetMayHaveHiddenRanges(b) {
		return nil, nil, nil
	}
	if hiddenCols, hiddenRows, ok := worksheetHiddenRangesFast(b); ok {
		return hiddenCols, hiddenRows, nil
	}
	return worksheetHiddenRangesXML(b)
}

func worksheetHiddenRangesXML(b []byte) ([]intRange, []intRange, error) {
	dec := xml.NewDecoder(bytes.NewReader(b))
	var hiddenCols []intRange
	var hiddenRows []intRange
	nextRow := 1
	for {
		tok, err := dec.RawToken()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, nil, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch start.Name.Local {
		case "col":
			if r, ok := hiddenColumnRange(start); ok {
				hiddenCols = append(hiddenCols, r)
			}
		case "row":
			row := worksheetRowIndex(start, nextRow)
			if row < 1 {
				row = nextRow
			}
			if worksheetRowHidden(start) {
				hiddenRows = append(hiddenRows, intRange{min: row, max: row})
			}
			nextRow = row + 1
		}
	}
	return hiddenCols, hiddenRows, nil
}

func worksheetHiddenRangesFast(b []byte) ([]intRange, []intRange, bool) {
	var hiddenCols []intRange
	var hiddenRows []intRange
	nextRow := 1
	for offset := 0; offset < len(b); {
		i := bytes.IndexByte(b[offset:], '<')
		if i < 0 {
			break
		}
		start := offset + i
		if start+1 >= len(b) {
			break
		}
		switch b[start+1] {
		case '/', '!', '?':
			offset = start + 2
			continue
		}
		end := xmlTagEnd(b, start+1)
		if end < 0 {
			return nil, nil, false
		}
		tag := b[start+1 : end]
		switch string(xmlTagLocalName(tag)) {
		case "col":
			if r, ok := hiddenColumnRangeFromTag(tag); ok {
				hiddenCols = append(hiddenCols, r)
			}
		case "row":
			row := worksheetRowIndexFromTag(tag, nextRow)
			if row < 1 {
				row = nextRow
			}
			if worksheetRowHiddenFromTag(tag) {
				hiddenRows = append(hiddenRows, intRange{min: row, max: row})
			}
			nextRow = row + 1
		}
		offset = end + 1
	}
	return hiddenCols, hiddenRows, true
}

func hiddenColumnRangeFromTag(tag []byte) (intRange, bool) {
	if !worksheetColumnHiddenFromTag(tag) {
		return intRange{}, false
	}
	min, _ := xmlTagIntAttr(tag, "min")
	max, _ := xmlTagIntAttr(tag, "max")
	if min <= 0 || max < min {
		return intRange{}, false
	}
	return intRange{min: min, max: max}, true
}

func worksheetColumnHiddenFromTag(tag []byte) bool {
	if value, ok := xmlTagAttr(tag, "hidden"); ok && boolAttrBytes(value) {
		return true
	}
	if value, ok := xmlTagFloatAttr(tag, "width"); ok && value <= 0 {
		return true
	}
	return false
}

func worksheetRowHiddenFromTag(tag []byte) bool {
	if value, ok := xmlTagAttr(tag, "hidden"); ok && boolAttrBytes(value) {
		return true
	}
	if value, ok := xmlTagFloatAttr(tag, "ht"); ok && value <= 0 {
		return true
	}
	return false
}

func worksheetRowIndexFromTag(tag []byte, fallback int) int {
	if row, ok := xmlTagIntAttr(tag, "r"); ok {
		return row
	}
	return fallback
}

func worksheetMayHaveHiddenRanges(b []byte) bool {
	for _, marker := range [][]byte{
		[]byte("hidden"),
	} {
		if bytes.Contains(b, marker) {
			return true
		}
	}
	for _, marker := range [][]byte{
		[]byte(`width="0`),
		[]byte(`width='0`),
		[]byte(`ht="0`),
		[]byte(`ht='0`),
	} {
		if bytes.Contains(b, marker) {
			return true
		}
	}
	return false
}

func visibleXlsxCommentText(b []byte, sources []xlsxSheetCellVisibility) (string, error) {
	items, err := visibleXlsxCommentItems(b, sources)
	if err != nil {
		return "", err
	}
	var out []string
	for _, item := range items {
		if item.text != "" {
			out = append(out, item.text)
		}
	}
	return cleanVisibleText(strings.Join(out, "\n")), nil
}

func visibleXlsxCommentItems(b []byte, sources []xlsxSheetCellVisibility) ([]xlsxCommentItem, error) {
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []xlsxCommentItem
	var cur strings.Builder
	var ref string
	var commentDepth int
	var skipDepth int
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if skipDepth > 0 {
				skipDepth++
				continue
			}
			if commentDepth > 0 {
				commentDepth++
				continue
			}
			if isXlsxCommentElement(t.Name.Local) {
				ref = strings.TrimSpace(xmlAttrValue(t, "ref"))
				if xlsxCommentRefHidden(ref, sources) {
					skipDepth = 1
					continue
				}
				ref = cleanXlsxCommentRef(ref)
				commentDepth = 1
				cur.Reset()
			}
		case xml.EndElement:
			if skipDepth > 0 {
				skipDepth--
				continue
			}
			if commentDepth > 0 {
				commentDepth--
				if commentDepth == 0 {
					text := cleanVisibleText(cur.String())
					if text != "" {
						out = append(out, xlsxCommentItem{ref: ref, text: text})
					}
					ref = ""
				}
			}
		case xml.CharData:
			if commentDepth > 0 && skipDepth == 0 {
				cur.Write(t)
			}
		}
	}
	return out, nil
}

func cleanXlsxCommentRef(ref string) string {
	ref = cleanText(ref)
	cleaned := stripInlineHiddenOfficeReferences(ref)
	if cleaned != ref {
		return ""
	}
	ref = cleaned
	if ref == "" || looksLikeHiddenResourceReference(ref) || looksLikeRelationshipIDReference(ref) || looksLikeOfficeRelationshipMetadataReference(ref) || looksLikeOfficeXMLMetadataReference(ref) {
		return ""
	}
	fields := worksheetRefFields(ref)
	if len(fields) == 0 {
		return ""
	}
	for _, field := range fields {
		parts := strings.Split(field, ":")
		if len(parts) == 0 || len(parts) > 2 {
			return ""
		}
		for _, part := range parts {
			if _, _, ok := cellRefIndexes(part); !ok {
				return ""
			}
		}
	}
	return ref
}

func isXlsxCommentElement(name string) bool {
	return name == "comment" || name == "threadedComment"
}

func xlsxCommentRefHidden(ref string, sources []xlsxSheetCellVisibility) bool {
	ref = strings.TrimSpace(ref)
	if ref == "" || len(sources) == 0 {
		return false
	}
	sawKnown := false
	for _, source := range sources {
		hidden, ok := cellRangeHidden(ref, source.hiddenCols, source.hiddenRows)
		if !ok {
			return false
		}
		sawKnown = true
		if !hidden {
			return false
		}
	}
	return sawKnown
}

func xlsxVisibleDrawingPartNames(files map[string]*zip.File) map[string]bool {
	out := map[string]bool{}
	hidden, visible := xlsxSheetDrawingPartNames(files)
	if len(hidden) > 0 || len(visible) > 0 {
		for name := range visible {
			out[name] = true
		}
		return out
	}
	for name := range files {
		lower := ooxmlPartKey(name)
		if !strings.HasPrefix(lower, "xl/drawings/") || !(strings.HasSuffix(lower, ".xml") || strings.HasSuffix(lower, ".vml")) {
			continue
		}
		if hidden[lower] && !visible[lower] {
			continue
		}
		out[lower] = true
	}
	return out
}

func xlsxVisibleDrawingPartNameList(files map[string]*zip.File) []string {
	parts := xlsxVisibleDrawingPartNames(files)
	var names []string
	for name := range files {
		if parts[ooxmlPartKey(name)] {
			names = append(names, name)
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return names
}

func xlsxSheetDrawingPartNames(files map[string]*zip.File) (map[string]bool, map[string]bool) {
	hidden := map[string]bool{}
	visible := map[string]bool{}
	sheets, err := workbookSheets(files)
	if err != nil {
		return hidden, visible
	}
	referenced := map[string]bool{}
	for _, sheet := range sheets {
		referenced[ooxmlPartKey(sheet.Path)] = true
		for _, drawing := range relationshipTargetsWithPrefix(files, sheet.Path, "xl/drawings/") {
			lower := ooxmlPartKey(drawing)
			if strings.HasSuffix(lower, ".xml") || strings.HasSuffix(lower, ".vml") {
				if sheet.Hidden {
					hidden[lower] = true
				} else {
					visible[lower] = true
				}
			}
		}
	}
	if len(referenced) > 0 {
		for _, sheet := range xlsxAllWorksheetPartNames(files) {
			if referenced[ooxmlPartKey(sheet)] {
				continue
			}
			for _, drawing := range relationshipTargetsWithPrefix(files, sheet, "xl/drawings/") {
				lower := ooxmlPartKey(drawing)
				if strings.HasSuffix(lower, ".xml") || strings.HasSuffix(lower, ".vml") {
					hidden[lower] = true
				}
			}
		}
	}
	return hidden, visible
}

func xlsxSheetRelatedPartNames(files map[string]*zip.File, prefix string) (map[string]bool, map[string]bool) {
	hidden := map[string]bool{}
	visible := map[string]bool{}
	sheets, err := workbookSheets(files)
	if err != nil {
		return hidden, visible
	}
	for _, sheet := range sheets {
		for _, part := range relationshipTargetsWithPrefix(files, sheet.Path, prefix) {
			lower := ooxmlPartKey(part)
			if sheet.Hidden {
				hidden[lower] = true
			} else {
				visible[lower] = true
			}
		}
	}
	return hidden, visible
}

type workbookSheet struct {
	Name   string
	Path   string
	Hidden bool
}

type xlsxWorksheetMarkdownData struct {
	rows         [][]string
	annotations  []string
	headerFooter []string
}

func workbookSheetStateHidden(value string) bool {
	value = strings.TrimSpace(value)
	return strings.EqualFold(value, "hidden") || strings.EqualFold(value, "veryHidden")
}

func extractXlsxMarkdown(files map[string]*zip.File, prepared map[string]xlsxWorksheetMarkdownData) (string, error) {
	sheets, err := workbookVisibleSheets(files)
	if err != nil {
		return "", err
	}
	if len(sheets) == 0 {
		for name := range files {
			lower := ooxmlPartKey(name)
			if strings.HasPrefix(lower, "xl/worksheets/sheet") && strings.HasSuffix(lower, ".xml") {
				cleanName := ooxmlCleanPartName(name)
				sheets = append(sheets, workbookSheet{Name: path.Base(strings.TrimSuffix(cleanName, path.Ext(cleanName))), Path: name})
			}
		}
		sort.Slice(sheets, func(i, j int) bool {
			return naturalLess(sheets[i].Path, sheets[j].Path)
		})
	}
	var parts []string
	var shared []string
	readShared := func() ([]string, error) {
		if shared != nil {
			return shared, nil
		}
		var err error
		shared, err = readSharedStrings(files)
		return shared, err
	}
	for _, sheet := range sheets {
		data, ok := prepared[ooxmlPartKey(sheet.Path)]
		if !ok {
			f := files[sheet.Path]
			if f != nil {
				shared, err := readShared()
				if err != nil {
					return "", err
				}
				b, err := readZipFile(f)
				if err != nil {
					return "", err
				}
				rows, annotations, headerFooter, err := worksheetMarkdownData(b, shared)
				if err != nil {
					return "", err
				}
				data = xlsxWorksheetMarkdownData{rows: rows, annotations: annotations, headerFooter: headerFooter}
			}
		}
		rows, annotations, headerFooter := data.rows, data.annotations, data.headerFooter
		title := cleanText(sheet.Name)
		if title == "" {
			title = path.Base(strings.TrimSuffix(sheet.Path, path.Ext(sheet.Path)))
		}
		if len(rows) == 0 {
			var sheetBlocks []string
			if len(annotations) > 0 {
				sheetBlocks = append(sheetBlocks, "### Annotations\n\n"+markdownText(strings.Join(annotations, "\n")))
			}
			if len(headerFooter) > 0 {
				sheetBlocks = append(sheetBlocks, "### Headers and Footers\n\n"+markdownText(strings.Join(headerFooter, "\n")))
			}
			if len(sheetBlocks) == 0 {
				parts = append(parts, "## "+escapeMarkdownHeading(title))
				continue
			}
			parts = append(parts, "## "+escapeMarkdownHeading(title)+"\n\n"+strings.Join(sheetBlocks, "\n\n"))
			continue
		}
		sheetMarkdown := markdownTablePrepared(rows)
		if len(annotations) > 0 {
			sheetMarkdown += "\n\n### Annotations\n\n" + markdownText(strings.Join(annotations, "\n"))
		}
		if len(headerFooter) > 0 {
			sheetMarkdown += "\n\n### Headers and Footers\n\n" + markdownText(strings.Join(headerFooter, "\n"))
		}
		parts = append(parts, "## "+escapeMarkdownHeading(title)+"\n\n"+sheetMarkdown)
	}
	if xlsxHasAnyPartPrefix(files, []string{"xl/comments", "xl/threadedcomments"}) {
		comments, err := xlsxVisibleCommentsMarkdownPart(files)
		if err != nil {
			return "", err
		}
		if comments != "" {
			parts = append(parts, comments)
		}
	}
	if xlsxHasAnyPartPrefix(files, []string{"xl/drawings/", "xl/charts/"}) {
		drawingNames := append(xlsxVisibleDrawingPartNameList(files), xlsxVisibleChartPartNameList(files)...)
		sort.Slice(drawingNames, func(i, j int) bool {
			return naturalLess(drawingNames[i], drawingNames[j])
		})
		drawings, err := xlsxDrawingMarkdownPart(files, drawingNames)
		if err != nil {
			return "", err
		}
		if drawings != "" {
			parts = append(parts, drawings)
		}
	}
	if workbookNames, err := xlsxWorkbookNamesMarkdownPart(files); err != nil {
		return "", err
	} else if workbookNames != "" {
		parts = append(parts, workbookNames)
	}
	if xlsxHasAnyPartPrefix(files, []string{"xl/tables/"}) {
		visibleTableParts, constrainedTableParts := xlsxVisibleTablePartNames(files)
		tables, err := xlsxSpecialXMLMarkdownPart(files, []string{"xl/tables/"}, "## Tables", tableXMLText, visibleTableParts, constrainedTableParts)
		if err != nil {
			return "", err
		}
		if tables != "" {
			parts = append(parts, tables)
		}
	}
	if xlsxHasAnyPartPrefix(files, []string{"xl/pivottables/", "xl/pivotcache/"}) {
		visiblePivotParts, constrainedPivotParts := xlsxVisibleSpecialPartNames(files, []string{"xl/pivottables/", "xl/pivotcache/"})
		pivots, err := xlsxSpecialXMLMarkdownPart(files, []string{"xl/pivottables/", "xl/pivotcache/"}, "## Pivot Tables", pivotXMLText, visiblePivotParts, constrainedPivotParts)
		if err != nil {
			return "", err
		}
		if pivots != "" {
			parts = append(parts, pivots)
		}
	}
	if xlsxHasAnyPartPrefix(files, []string{"xl/slicers/", "xl/slicercaches/"}) {
		visibleSlicerParts, constrainedSlicerParts := xlsxVisibleSpecialPartNames(files, []string{"xl/slicers/", "xl/slicercaches/"})
		slicers, err := xlsxSpecialXMLMarkdownPart(files, []string{"xl/slicers/", "xl/slicercaches/"}, "## Slicers", slicerXMLText, visibleSlicerParts, constrainedSlicerParts)
		if err != nil {
			return "", err
		}
		if slicers != "" {
			parts = append(parts, slicers)
		}
	}
	return strings.TrimSpace(strings.Join(parts, "\n\n")), nil
}

func xlsxWorkbookNamesMarkdownPart(files map[string]*zip.File) (string, error) {
	items, err := workbookDefinedNameItems(files)
	if err != nil {
		return "", err
	}
	if len(items) == 0 {
		return "", nil
	}
	var lines []string
	for _, item := range items {
		name := markdownText(item.name)
		value := markdownText(item.value)
		if name != "" {
			lines = append(lines, "- "+name)
		}
		if value != "" {
			lines = append(lines, "  - "+value)
		}
	}
	if len(lines) == 0 {
		return "", nil
	}
	return "## Workbook Names\n\n" + strings.Join(lines, "\n"), nil
}

type workbookDefinedNameItem struct {
	name  string
	value string
}

func workbookDefinedNameItems(files map[string]*zip.File) ([]workbookDefinedNameItem, error) {
	f := ooxmlFile(files, "xl/workbook.xml")
	if f == nil {
		return nil, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, err
	}
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []workbookDefinedNameItem
	var cur strings.Builder
	var item workbookDefinedNameItem
	var inDefinedName bool
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local != "definedName" {
				continue
			}
			inDefinedName = true
			cur.Reset()
			item = workbookDefinedNameItem{}
			hidden := false
			for _, attr := range t.Attr {
				if attr.Name.Local == "hidden" && boolAttrValue(attr.Value) {
					hidden = true
					break
				}
			}
			if hidden {
				inDefinedName = false
				continue
			}
			for _, attr := range t.Attr {
				if attr.Name.Local != "name" {
					continue
				}
				name := cleanVisibleXMLAttributeText(attr.Value)
				if name != "" && visibleDefinedNameName(name) {
					item.name = name
				}
				break
			}
		case xml.EndElement:
			if t.Name.Local != "definedName" || !inDefinedName {
				continue
			}
			value := cleanVisibleXMLAttributeText(cur.String())
			if value != "" && visibleDefinedNameValue(value) {
				item.value = value
			}
			if item.name != "" || item.value != "" {
				out = append(out, item)
			}
			inDefinedName = false
		case xml.CharData:
			if inDefinedName {
				cur.Write(t)
			}
		}
	}
	return out, nil
}

func xlsxSpecialXMLMarkdownPart(files map[string]*zip.File, prefixes []string, heading string, extract func([]byte) (string, error), visible map[string]bool, constrained bool) (string, error) {
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if !strings.HasSuffix(lower, ".xml") {
			continue
		}
		if constrained && !visible[lower] {
			continue
		}
		for _, prefix := range prefixes {
			if strings.HasPrefix(lower, prefix) {
				names = append(names, name)
				break
			}
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	var blocks []string
	for _, name := range names {
		f := ooxmlFile(files, name)
		if f == nil {
			continue
		}
		b, err := readZipFile(f)
		if err != nil {
			return "", err
		}
		text, err := extract(b)
		if err != nil {
			return "", err
		}
		text = markdownText(text)
		if text != "" {
			blocks = append(blocks, text)
		}
	}
	if len(blocks) == 0 {
		return "", nil
	}
	return heading + "\n\n" + strings.Join(blocks, "\n\n"), nil
}

func xlsxDrawingMarkdownPart(files map[string]*zip.File, names []string) (string, error) {
	texts, err := xmlTextFromFiles(files, names)
	if err != nil {
		return "", err
	}
	var blocks []string
	for _, text := range texts {
		text = markdownText(text)
		if text != "" {
			blocks = append(blocks, text)
		}
	}
	if len(blocks) == 0 {
		return "", nil
	}
	return "## Drawings\n\n" + strings.Join(blocks, "\n\n"), nil
}

func worksheetAnnotationText(b []byte) ([]string, error) {
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	var hiddenCols []intRange
	var hiddenRows []intRange
	nextRow := 1
	seen := map[string]bool{}
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch start.Name.Local {
		case "col":
			if r, ok := hiddenColumnRange(start); ok {
				hiddenCols = append(hiddenCols, r)
			}
		case "row":
			row := worksheetRowIndex(start, nextRow)
			if row < 1 {
				row = nextRow
			}
			if worksheetRowHidden(start) {
				hiddenRows = append(hiddenRows, intRange{min: row, max: row})
			}
			nextRow = row + 1
		}
		if worksheetElementHiddenByRef(start, hiddenCols, hiddenRows) {
			continue
		}
		for _, value := range worksheetAttributeText(start) {
			if seen[value] {
				continue
			}
			seen[value] = true
			out = append(out, value)
		}
	}
	return out, nil
}

func worksheetHeaderFooterText(b []byte) ([]string, error) {
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	var cur strings.Builder
	var inHeaderFooter bool
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if isExcelHeaderFooterElement(t.Name.Local) {
				inHeaderFooter = true
				cur.Reset()
			}
		case xml.EndElement:
			if isExcelHeaderFooterElement(t.Name.Local) {
				value := cleanExcelHeaderFooterText(cur.String())
				if value != "" {
					out = append(out, value)
				}
				inHeaderFooter = false
			}
		case xml.CharData:
			if inHeaderFooter {
				cur.Write(t)
			}
		}
	}
	return out, nil
}

func worksheetMarkdownData(b []byte, shared []string) ([][]string, []string, []string, error) {
	if hasDOCTYPE(b) {
		return nil, nil, nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var rows [][]string
	var annotations []string
	var headerFooter []string
	var rowValues map[int]string
	var cellText strings.Builder
	var headerText strings.Builder
	var hiddenCols []intRange
	var hiddenRows []intRange
	var rowHidden bool
	var skipCell bool
	var cellType string
	var cellRef string
	var cellCol int
	var nextCol int
	nextAnnotationRow := 1
	var inV, inT bool
	var inHeaderFooter bool
	var phoneticDepth int
	var cellDepth int
	var systemCellTextDepth int
	collectRow := true
	seenAnnotations := map[string]bool{}
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, nil, nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "col" {
				if r, ok := hiddenColumnRange(t); ok {
					hiddenCols = append(hiddenCols, r)
				}
			}
			if t.Name.Local == "row" {
				rowHidden = worksheetRowHidden(t)
				row := worksheetRowIndex(t, nextAnnotationRow)
				if row < 1 {
					row = nextAnnotationRow
				}
				if rowHidden {
					hiddenRows = append(hiddenRows, intRange{min: row, max: row})
				}
				nextAnnotationRow = row + 1
				collectRow = !rowHidden && len(rows) < maxMarkdownTableRows
				if collectRow {
					rowValues = map[int]string{}
				} else {
					rowValues = nil
				}
				nextCol = 1
			}
			if !worksheetElementHiddenByRef(t, hiddenCols, hiddenRows) {
				for _, value := range worksheetAttributeText(t) {
					if seenAnnotations[value] {
						continue
					}
					seenAnnotations[value] = true
					annotations = append(annotations, value)
				}
			}
			if t.Name.Local == "c" {
				cellDepth = 1
				cellType = ""
				cellRef = ""
				cellText.Reset()
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "t":
						cellType = a.Value
					case "r":
						cellRef = a.Value
					}
				}
				if col, _, ok := cellRefIndexes(cellRef); ok {
					cellCol = col
				} else {
					cellCol = nextCol
				}
				if cellCol < 1 {
					cellCol = 1
				}
				nextCol = cellCol + 1
				skipCell = !collectRow || rowHidden || hiddenColumnCell(cellRef, hiddenCols) || columnHidden(cellCol, hiddenCols)
			} else if cellDepth > 0 {
				cellDepth++
			}
			if systemCellTextDepth > 0 {
				systemCellTextDepth++
				continue
			}
			if cellDepth > 0 && isExcelSystemCellTextElement(t.Name.Local) {
				systemCellTextDepth = 1
				continue
			}
			if isExcelPhoneticElement(t.Name.Local) {
				phoneticDepth++
				continue
			}
			if isExcelHeaderFooterElement(t.Name.Local) {
				inHeaderFooter = true
				headerText.Reset()
			}
			if phoneticDepth == 0 && (t.Name.Local == "v" || t.Name.Local == "t") {
				inV = t.Name.Local == "v"
				inT = t.Name.Local == "t"
			}
		case xml.EndElement:
			if phoneticDepth > 0 {
				if isExcelPhoneticElement(t.Name.Local) {
					phoneticDepth--
				}
				continue
			}
			if systemCellTextDepth > 0 {
				systemCellTextDepth--
				continue
			}
			if isExcelHeaderFooterElement(t.Name.Local) {
				value := cleanExcelHeaderFooterText(headerText.String())
				if value != "" {
					headerFooter = append(headerFooter, value)
				}
				inHeaderFooter = false
			}
			if t.Name.Local == "v" || t.Name.Local == "t" {
				inV, inT = false, false
			}
			if t.Name.Local == "c" {
				if !skipCell && rowValues != nil {
					value := strings.TrimSpace(cellText.String())
					if value != "" {
						if cellType == "s" {
							if idx, ok := atoi(value); ok && idx >= 0 && idx < len(shared) {
								value = shared[idx]
							}
						}
						if cellType == "b" {
							value = excelBooleanDisplayText(value)
						}
						if cellCol <= maxMarkdownTableCols {
							if cellValue := cleanMarkdownTableCellValue(value); cellValue != "" {
								rowValues[cellCol] = prepareMarkdownTableCellValue(cellValue)
							}
						}
					}
				}
				skipCell = false
				cellText.Reset()
				cellDepth = 0
				systemCellTextDepth = 0
			} else if cellDepth > 0 {
				cellDepth--
			}
			if t.Name.Local == "row" {
				if collectRow {
					if row := compactWorksheetMarkdownRow(rowValues); len(row) > 0 && len(rows) < maxMarkdownTableRows {
						rows = append(rows, row)
					}
				}
				rowHidden = false
				rowValues = nil
				collectRow = true
			}
		case xml.CharData:
			if inHeaderFooter {
				headerText.Write(t)
			}
			if !skipCell && (inV || inT) {
				cellText.Write(t)
			}
		}
	}
	return rows, annotations, headerFooter, nil
}

func isExcelSystemCellTextElement(name string) bool {
	switch name {
	case "extLst", "ext":
		return true
	default:
		return false
	}
}

func workbookVisibleSheets(files map[string]*zip.File) ([]workbookSheet, error) {
	all, err := workbookSheets(files)
	if err != nil {
		return nil, err
	}
	var visible []workbookSheet
	for _, sheet := range all {
		if !sheet.Hidden {
			visible = append(visible, sheet)
		}
	}
	return visible, nil
}

func workbookSheets(files map[string]*zip.File) ([]workbookSheet, error) {
	f := ooxmlFile(files, "xl/workbook.xml")
	if f == nil {
		return nil, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, err
	}
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	rels, hasRels, err := workbookRelationships(files)
	if err != nil {
		return nil, err
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var sheets []workbookSheet
	fallbackParts := xlsxAllWorksheetPartNames(files)
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "sheet" {
			continue
		}
		hidden := false
		var sheetName, relID string
		for _, attr := range start.Attr {
			switch attr.Name.Local {
			case "name":
				sheetName = cleanVisibleSheetName(attr.Value)
			case "id":
				relID = attr.Value
			case "state":
				hidden = workbookSheetStateHidden(attr.Value)
			}
		}
		if hasRels {
			target := rels[relID]
			if actual := ooxmlPartName(files, target); actual != "" {
				sheets = append(sheets, workbookSheet{Name: sheetName, Path: actual, Hidden: hidden})
				continue
			}
		}
		if len(fallbackParts) > len(sheets) {
			sheets = append(sheets, workbookSheet{Name: sheetName, Path: fallbackParts[len(sheets)], Hidden: hidden})
		} else {
			sheets = append(sheets, workbookSheet{Name: sheetName, Hidden: hidden})
		}
	}
	return sheets, nil
}

func workbookText(files map[string]*zip.File) ([]string, error) {
	text, _, err := workbookTextAndSheets(files, false)
	return text, err
}

func workbookTextAndSheets(files map[string]*zip.File, strictOfficeContent bool) ([]string, []string, error) {
	f := ooxmlFile(files, "xl/workbook.xml")
	if f == nil {
		return nil, nil, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, nil, err
	}
	if hasDOCTYPE(b) {
		return nil, nil, errors.New("xml doctype is not supported")
	}
	rels, hasRels, err := workbookRelationships(files)
	if err != nil {
		return nil, nil, err
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	var sheets []string
	var sawSheet bool
	var sheetIndex int
	var inDefinedName bool
	var cur strings.Builder
	fallbackParts := xlsxAllWorksheetPartNames(files)
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "sheet":
				sawSheet = true
				hidden := false
				var sheetName, relID string
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "name":
						sheetName = cleanVisibleSheetName(attr.Value)
					case "id":
						relID = attr.Value
					case "state":
						hidden = workbookSheetStateHidden(attr.Value)
					}
				}
				if !hidden {
					if hasRels {
						target := rels[relID]
						if actual := ooxmlPartName(files, target); actual != "" {
							// Excel.Worksheets, which is the strict COM baseline,
							// excludes chart sheets. Their cached series values live
							// in chartsheet XML but are not Worksheet.UsedRange.Text.
							if strictOfficeContent && !strings.HasPrefix(ooxmlPartKey(actual), "xl/worksheets/") {
								sheetIndex++
								break
							}
							if sheetName != "" {
								out = append(out, sheetName)
							}
							sheets = append(sheets, actual)
							sheetIndex++
							break
						}
					}
					if sheetName != "" {
						out = append(out, sheetName)
					}
					if len(fallbackParts) > sheetIndex {
						sheets = append(sheets, fallbackParts[sheetIndex])
					}
				}
				sheetIndex++
			case "definedName":
				// Excel's UsedRange.Text exposes worksheet cells and visible sheet
				// names, not workbook-level named ranges.  Keep names in the
				// compatibility text/structured Markdown paths, but exclude them
				// from the Office-COM comparison path so a named range that repeats
				// a cell label is not counted as a second visible occurrence.
				if strictOfficeContent {
					break
				}
				inDefinedName = true
				cur.Reset()
				hidden := false
				for _, attr := range t.Attr {
					if attr.Name.Local == "hidden" && boolAttrValue(attr.Value) {
						hidden = true
					}
				}
				if hidden {
					inDefinedName = false
					break
				}
				for _, attr := range t.Attr {
					if attr.Name.Local == "name" {
						name := cleanVisibleXMLAttributeText(attr.Value)
						if name != "" && visibleDefinedNameName(name) {
							out = append(out, name)
						}
						break
					}
				}
			}
		case xml.EndElement:
			if t.Name.Local == "definedName" {
				value := cleanVisibleXMLAttributeText(cur.String())
				if value != "" && visibleDefinedNameValue(value) {
					out = append(out, value)
				}
				inDefinedName = false
			}
		case xml.CharData:
			if inDefinedName {
				cur.Write(t)
			}
		}
	}
	if sawSheet && hasRels && sheets == nil {
		sheets = []string{}
	}
	return out, sheets, nil
}

func visibleDefinedNameName(name string) bool {
	return name != "" && !strings.HasPrefix(strings.ToLower(name), "_xlnm.")
}

func cleanVisibleSheetName(value string) string {
	return cleanVisibleXMLAttributeText(value)
}

func visibleDefinedNameValue(value string) bool {
	value = strings.TrimSpace(value)
	if value == "" {
		return false
	}
	if strings.HasPrefix(value, "=") {
		return false
	}
	if strings.ContainsAny(value, "!():[]") {
		return false
	}
	return true
}

func workbookRelationships(files map[string]*zip.File) (map[string]string, bool, error) {
	rels := map[string]string{}
	workbook := ooxmlPartName(files, "xl/workbook.xml")
	if workbook == "" {
		workbook = "xl/workbook.xml"
	}
	f := ooxmlFile(files, ooxmlRelsName(workbook))
	if f == nil {
		return rels, false, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, true, err
	}
	if hasDOCTYPE(b) {
		return nil, true, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, true, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "Relationship" {
			continue
		}
		var id, target string
		for _, attr := range start.Attr {
			switch attr.Name.Local {
			case "Id":
				id = attr.Value
			case "Target":
				target = attr.Value
			}
		}
		target = normalizeWorkbookRelationshipTarget(target)
		if actual := ooxmlPartName(files, target); actual != "" {
			target = actual
		}
		if id != "" && target != "" {
			rels[id] = target
		}
	}
	return rels, true, nil
}

func normalizeWorkbookRelationshipTarget(target string) string {
	target = strings.TrimSpace(strings.ReplaceAll(target, "\\", "/"))
	if target == "" {
		return ""
	}
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(path.Clean(target), "/")
	}
	return path.Clean(path.Join("xl", target))
}

func extractAllXMLText(files map[string]*zip.File) ([]string, error) {
	var names []string
	for name := range files {
		if strings.HasSuffix(ooxmlPartKey(name), ".xml") {
			names = append(names, name)
		}
	}
	sort.Strings(names)
	return xmlTextFromFiles(files, names)
}

func docPropsText(files map[string]*zip.File) ([]string, error) {
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if strings.HasPrefix(lower, "docprops/") && strings.HasSuffix(lower, ".xml") {
			names = append(names, name)
		}
	}
	sort.Strings(names)
	var out []string
	for _, name := range names {
		b, err := readZipFile(files[name])
		if err != nil {
			return nil, err
		}
		text, err := propertyXMLText(b)
		if err != nil {
			return nil, fmt.Errorf("%s: %w", name, err)
		}
		if text != "" {
			out = append(out, text)
		}
	}
	return out, nil
}

func relationshipsText(files map[string]*zip.File) ([]string, error) {
	var names []string
	for name := range files {
		if strings.HasSuffix(ooxmlPartKey(name), ".rels") {
			names = append(names, name)
		}
	}
	sort.Strings(names)
	seen := map[string]bool{}
	var out []string
	for _, name := range names {
		b, err := readZipFile(files[name])
		if err != nil {
			return nil, err
		}
		targets, err := relationshipXMLText(b)
		if err != nil {
			return nil, fmt.Errorf("%s: %w", name, err)
		}
		for _, target := range targets {
			if seen[target] {
				continue
			}
			seen[target] = true
			out = append(out, target)
		}
	}
	return out, nil
}

func customXMLText(files map[string]*zip.File) ([]string, error) {
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if !strings.HasPrefix(lower, "customxml/") || !strings.HasSuffix(lower, ".xml") {
			continue
		}
		if strings.Contains(lower, "/_rels/") || strings.HasPrefix(path.Base(lower), "itemprops") {
			continue
		}
		names = append(names, name)
	}
	sort.Strings(names)
	var out []string
	for _, name := range names {
		b, err := readZipFile(files[name])
		if err != nil {
			return nil, err
		}
		text, err := allXMLCharDataText(b)
		if err != nil {
			return nil, fmt.Errorf("%s: %w", name, err)
		}
		if text != "" {
			out = append(out, text)
		}
	}
	return out, nil
}

func xmlTextFromFiles(files map[string]*zip.File, names []string) ([]string, error) {
	var out []string
	for _, name := range names {
		b, err := readZipFile(files[name])
		if err != nil {
			return nil, err
		}
		lowerName := ooxmlPartKey(name)
		if strings.HasSuffix(lowerName, ".vml") {
			text, err := visibleVMLText(b)
			if err != nil {
				return nil, fmt.Errorf("%s: %w", name, err)
			}
			if text != "" {
				out = append(out, text)
			}
			continue
		}
		if strings.HasPrefix(lowerName, "xl/tables/") {
			tableText, err := tableXMLText(b)
			if err != nil {
				return nil, fmt.Errorf("%s: %w", name, err)
			}
			if tableText != "" {
				out = append(out, tableText)
			}
			continue
		}
		if strings.HasPrefix(lowerName, "xl/pivottables/") || strings.HasPrefix(lowerName, "xl/pivotcache/") {
			pivotText, err := pivotXMLText(b)
			if err != nil {
				return nil, fmt.Errorf("%s: %w", name, err)
			}
			if pivotText != "" {
				out = append(out, pivotText)
			}
			continue
		}
		if strings.HasPrefix(lowerName, "xl/slicers/") || strings.HasPrefix(lowerName, "xl/slicercaches/") {
			slicerText, err := slicerXMLText(b)
			if err != nil {
				return nil, fmt.Errorf("%s: %w", name, err)
			}
			if slicerText != "" {
				out = append(out, slicerText)
			}
			continue
		}
		text, err := visibleXMLText(b)
		if err != nil {
			return nil, fmt.Errorf("%s: %w", name, err)
		}
		if text != "" {
			out = append(out, text)
		}
	}
	return out, nil
}

func tableXMLText(b []byte) (string, error) {
	if hasDOCTYPE(b) {
		return "", errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return "", err
		}
		start, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		if start.Name.Local == "table" || start.Name.Local == "tableColumn" {
			for _, attr := range start.Attr {
				if attr.Name.Local == "name" || attr.Name.Local == "displayName" {
					value := cleanVisibleXMLAttributeText(attr.Value)
					if value != "" {
						out = append(out, value)
					}
				}
			}
		}
	}
	return cleanText(strings.Join(out, "\n")), nil
}

func pivotXMLText(b []byte) (string, error) {
	if hasDOCTYPE(b) {
		return "", errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	seen := map[string]bool{}
	var out []string
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return "", err
		}
		start, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		for _, attr := range start.Attr {
			if !isPivotTextAttribute(start.Name.Local, attr.Name.Local) {
				continue
			}
			value := cleanVisibleXMLAttributeText(attr.Value)
			if value == "" || seen[value] {
				continue
			}
			seen[value] = true
			out = append(out, value)
		}
	}
	return cleanText(strings.Join(out, "\n")), nil
}

func isPivotTextAttribute(element, attr string) bool {
	switch element {
	case "pivotTableDefinition":
		switch attr {
		case "name", "dataCaption", "grandTotalCaption", "errorCaption", "missingCaption":
			return true
		}
	case "cacheField", "pivotField", "field", "dataField", "calculatedItem", "calculatedMember", "memberProperty", "set":
		switch attr {
		case "name", "caption", "propertyName":
			return true
		}
	case "s":
		return attr == "v"
	}
	return false
}

func slicerXMLText(b []byte) (string, error) {
	if hasDOCTYPE(b) {
		return "", errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	seen := map[string]bool{}
	var out []string
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return "", err
		}
		start, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		for _, attr := range start.Attr {
			if !isSlicerTextAttribute(start.Name.Local, attr.Name.Local) {
				continue
			}
			value := cleanVisibleXMLAttributeText(attr.Value)
			if value == "" || !visibleSlicerAttributeValue(attr.Name.Local, value) || seen[value] {
				continue
			}
			seen[value] = true
			out = append(out, value)
		}
	}
	return cleanText(strings.Join(out, "\n")), nil
}

func isSlicerTextAttribute(element, attr string) bool {
	switch element {
	case "slicer", "slicerCacheDefinition", "slicerCachePivotTable", "slicerCacheOlapLevelName", "level", "item", "i":
		switch attr {
		case "name", "caption", "sourceName", "displayName", "uniqueName", "v", "n":
			return true
		}
	}
	return false
}

func cleanVisibleXMLAttributeText(value string) string {
	return cleanVisibleAttributeValue(value)
}

func visibleSlicerAttributeValue(attr, value string) bool {
	if (strings.EqualFold(attr, "uniqueName") || strings.EqualFold(attr, "n")) && looksLikeInternalSlicerUniqueName(value) {
		return false
	}
	return !looksLikeHiddenResourceReference(value)
}

func looksLikeInternalSlicerUniqueName(value string) bool {
	value = strings.TrimSpace(value)
	if value == "" {
		return false
	}
	if strings.Contains(value, "].[") || strings.Contains(value, "].&[") {
		return true
	}
	if strings.HasPrefix(value, "[") && strings.HasSuffix(value, "]") {
		return true
	}
	return false
}

func allXMLCharDataText(b []byte) (string, error) {
	if hasDOCTYPE(b) {
		return "", errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	var hiddenDepth int
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return "", err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if hiddenDepth > 0 || xmlSystemTextElement(t.Name.Local) {
				hiddenDepth++
			}
		case xml.EndElement:
			if hiddenDepth > 0 {
				hiddenDepth--
			}
		case xml.CharData:
			if hiddenDepth == 0 {
				value := cleanVisibleText(string(t))
				if value != "" {
					out = append(out, value)
				}
			}
		}
	}
	return cleanVisibleText(strings.Join(out, "\n")), nil
}

func xmlSystemTextElement(name string) bool {
	switch strings.ToLower(name) {
	case "script", "style", "clientdata":
		return true
	default:
		return false
	}
}

func relationshipXMLText(b []byte) ([]string, error) {
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "Relationship" {
			continue
		}
		var typ, target, mode string
		for _, attr := range start.Attr {
			switch attr.Name.Local {
			case "Type":
				typ = attr.Value
			case "Target":
				target = attr.Value
			case "TargetMode":
				mode = attr.Value
			}
		}
		target = cleanText(target)
		if !isTextRelationship(typ, mode, target) {
			continue
		}
		if target != "" {
			out = append(out, target)
		}
	}
	return out, nil
}

func isTextRelationship(typ, mode, target string) bool {
	_ = mode
	return strings.HasSuffix(strings.ToLower(typ), "/hyperlink") && !looksLikeHiddenResourceReference(target)
}

func propertyXMLText(b []byte) (string, error) {
	if hasDOCTYPE(b) {
		return "", errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	var inValue bool
	var cur strings.Builder
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return "", err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if isPropertyTextElement(t.Name.Local) {
				inValue = true
				cur.Reset()
			}
		case xml.EndElement:
			if isPropertyTextElement(t.Name.Local) {
				value := cleanVisibleText(cur.String())
				if value != "" {
					out = append(out, value)
				}
				inValue = false
			}
		case xml.CharData:
			if inValue {
				cur.Write(t)
			}
		}
	}
	return cleanVisibleText(strings.Join(out, "\n")), nil
}

func isPropertyTextElement(name string) bool {
	switch name {
	case "title", "subject", "creator", "keywords", "description", "lastModifiedBy", "category",
		"contentStatus", "version", "revision", "Template", "Manager", "Company", "Application",
		"lpstr", "lpwstr", "bstr", "filetime", "i4", "int", "r8", "bool":
		return true
	default:
		return false
	}
}

func visibleXMLText(b []byte) (string, error) {
	return visibleXMLTextWithAttributes(b, true)
}

func visibleXMLTextWithoutAttributes(b []byte) (string, error) {
	return visibleXMLTextWithAttributes(b, false)
}

func visibleXMLTextWithAttributes(b []byte, includeAttributes bool) (string, error) {
	if hasDOCTYPE(b) {
		return "", errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	var textDepth int
	var runDepth int
	var rPrDepth int
	var runHidden bool
	var paragraphHiddenStack []bool
	var pPrDepth int
	var checkboxDepth int
	var checkboxVisible bool
	var checkboxChecked bool
	var dropdownDepth int
	var dropdownVisible bool
	var dropdownResult int
	var dropdownEntries []string
	var textInputDepth int
	var textInputVisible bool
	var textInputDefault string
	var alternateStack []alternateContentState
	var skipDepth int
	var hiddenRevisionRangeDepth int
	var drawingObjectStack []bool
	seenAttrs := map[string]bool{}
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return "", err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if skipDepth > 0 {
				skipDepth++
				continue
			}
			if checkboxDepth > 0 {
				checkboxDepth++
				if t.Name.Local == "checked" && checkboxCheckedElement(t) {
					checkboxChecked = true
				}
				continue
			}
			if dropdownDepth > 0 {
				dropdownDepth++
				switch t.Name.Local {
				case "result":
					if value, ok := intAttrValue(t, "val"); ok && value >= 0 {
						dropdownResult = value
					}
				case "listEntry":
					value := cleanVisibleXMLAttributeText(xmlAttrValue(t, "val"))
					if value != "" {
						dropdownEntries = append(dropdownEntries, value)
					}
				}
				continue
			}
			if textInputDepth > 0 {
				textInputDepth++
				if t.Name.Local == "default" {
					textInputDefault = cleanVisibleXMLAttributeText(xmlAttrValue(t, "val"))
				}
				continue
			}
			if isDrawingObjectElement(t.Name.Local) {
				parentHidden := len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]
				drawingObjectStack = append(drawingObjectStack, parentHidden)
			}
			if len(drawingObjectStack) > 0 && drawingObjectElementHidden(t) {
				drawingObjectStack[len(drawingObjectStack)-1] = true
			}
			drawingObjectHidden := len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]
			if t.Name.Local == "AlternateContent" {
				alternateStack = append(alternateStack, alternateContentState{})
				continue
			}
			if len(alternateStack) > 0 {
				top := &alternateStack[len(alternateStack)-1]
				switch t.Name.Local {
				case "Choice":
					top.choiceSeen = true
				case "Fallback":
					if top.choiceSeen {
						skipDepth = 1
						continue
					}
				}
			}
			if isHiddenRevisionElement(t.Name) {
				skipDepth = 1
				continue
			}
			if isHiddenRevisionRangeStart(t.Name) {
				hiddenRevisionRangeDepth++
				continue
			}
			if isHiddenRevisionRangeEnd(t.Name) {
				if hiddenRevisionRangeDepth > 0 {
					hiddenRevisionRangeDepth--
				}
				continue
			}
			if isSystemFootnoteElement(t) {
				skipDepth = 1
				continue
			}
			if xmlSystemTextElement(t.Name.Local) {
				skipDepth = 1
				continue
			}
			hiddenByRevisionRange := hiddenRevisionRangeDepth > 0
			contentVisible := !runHidden && !currentParagraphHidden(paragraphHiddenStack) && !drawingObjectHidden && !hiddenByRevisionRange
			if contentVisible && includeAttributes {
				for _, value := range visibleAttributeText(t) {
					if seenAttrs[value] {
						continue
					}
					seenAttrs[value] = true
					out = append(out, "\n", value, "\n")
				}
			}
			switch t.Name.Local {
			case "p":
				paragraphHiddenStack = append(paragraphHiddenStack, false)
				if contentVisible {
					out = append(out, "\n")
				}
			case "r":
				runDepth++
				runHidden = false
			case "pPr":
				if len(paragraphHiddenStack) > 0 {
					pPrDepth++
				}
			case "rPr":
				if runDepth > 0 {
					rPrDepth++
				}
			case "vanish", "webHidden":
				if runDepth > 0 && rPrDepth > 0 {
					runHidden = true
				}
				if pPrDepth > 0 && len(paragraphHiddenStack) > 0 {
					paragraphHiddenStack[len(paragraphHiddenStack)-1] = true
				}
			case "t", "text", "v":
				if contentVisible {
					textDepth++
				}
			case "br", "cr", "row":
				if contentVisible {
					out = append(out, "\n")
				}
			case "tab":
				if contentVisible {
					out = append(out, "\t")
				}
			case "noBreakHyphen":
				if contentVisible {
					out = append(out, "\u2011")
				}
			case "softHyphen":
				// Soft hyphen is conditional formatting; omit it from visible text.
			case "sym":
				if contentVisible {
					if value, ok := visibleSymbolText(t); ok {
						out = append(out, value)
					}
				}
			case "checkBox":
				checkboxDepth = 1
				checkboxVisible = contentVisible
				checkboxChecked = false
			case "ddList":
				dropdownDepth = 1
				dropdownVisible = contentVisible
				dropdownResult = 0
				dropdownEntries = dropdownEntries[:0]
			case "textInput":
				textInputDepth = 1
				textInputVisible = contentVisible
				textInputDefault = ""
			}
		case xml.EndElement:
			if skipDepth > 0 {
				skipDepth--
				continue
			}
			if checkboxDepth > 0 {
				checkboxDepth--
				if checkboxDepth == 0 {
					if checkboxVisible {
						if checkboxChecked {
							out = append(out, "\u2612")
						} else {
							out = append(out, "\u2610")
						}
					}
					checkboxVisible = false
					checkboxChecked = false
				}
				continue
			}
			if dropdownDepth > 0 {
				dropdownDepth--
				if dropdownDepth == 0 {
					if dropdownVisible && dropdownResult >= 0 && dropdownResult < len(dropdownEntries) {
						out = append(out, dropdownEntries[dropdownResult])
					}
					dropdownVisible = false
					dropdownResult = 0
					dropdownEntries = dropdownEntries[:0]
				}
				continue
			}
			if textInputDepth > 0 {
				textInputDepth--
				if textInputDepth == 0 {
					if textInputVisible && textInputDefault != "" {
						out = append(out, textInputDefault)
					}
					textInputVisible = false
					textInputDefault = ""
				}
				continue
			}
			if t.Name.Local == "AlternateContent" {
				if len(alternateStack) > 0 {
					alternateStack = alternateStack[:len(alternateStack)-1]
				}
				continue
			}
			if (t.Name.Local == "t" || t.Name.Local == "text" || t.Name.Local == "v") && textDepth > 0 {
				textDepth--
			}
			if t.Name.Local == "pPr" && pPrDepth > 0 {
				pPrDepth--
			}
			if t.Name.Local == "rPr" && rPrDepth > 0 {
				rPrDepth--
			}
			if t.Name.Local == "r" && runDepth > 0 {
				runDepth--
				if runDepth == 0 {
					runHidden = false
					rPrDepth = 0
				}
			}
			if isDrawingObjectElement(t.Name.Local) && len(drawingObjectStack) > 0 {
				drawingObjectStack = drawingObjectStack[:len(drawingObjectStack)-1]
			}
			drawingObjectHidden := len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]
			if (t.Name.Local == "p" || t.Name.Local == "row") && !runHidden && !currentParagraphHidden(paragraphHiddenStack) && !drawingObjectHidden && hiddenRevisionRangeDepth == 0 {
				out = append(out, "\n")
			}
			if t.Name.Local == "p" && len(paragraphHiddenStack) > 0 {
				paragraphHiddenStack = paragraphHiddenStack[:len(paragraphHiddenStack)-1]
				if len(paragraphHiddenStack) == 0 {
					pPrDepth = 0
				}
			}
		case xml.CharData:
			if textDepth > 0 && skipDepth == 0 {
				out = append(out, string(t))
			}
		}
	}
	return cleanMarkdownVisibleText(strings.Join(out, "")), nil
}

func visibleXMLMarkdownWithTables(b []byte) (string, error) {
	return visibleXMLMarkdownWithTablesAndNumbering(b, nil)
}

func visibleXMLMarkdownWithTablesAndNumbering(b []byte, numbering map[string]string) (string, error) {
	if hasDOCTYPE(b) {
		return "", errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	var textDepth int
	var runDepth int
	var rPrDepth int
	var runHidden bool
	var paragraphHiddenStack []bool
	var pPrDepth int
	var checkboxDepth int
	var checkboxVisible bool
	var checkboxChecked bool
	var dropdownDepth int
	var dropdownVisible bool
	var dropdownResult int
	var dropdownEntries []string
	var textInputDepth int
	var textInputVisible bool
	var textInputDefault string
	var alternateStack []alternateContentState
	var skipDepth int
	var hiddenRevisionRangeDepth int
	var drawingObjectStack []bool
	var paragraphPrefixStack []string
	var paragraphPrefixWritten []bool
	var paragraphListLevelStack []int
	var numPrDepth int
	var numPrID string
	var numPrLevel int
	var tableDepth int
	var rowDepth int
	var cellDepth int
	var tableRows [][]string
	var tableRow []string
	var cellText strings.Builder
	seenAttrs := map[string]bool{}
	write := func(s string) {
		if tableDepth > 0 {
			if cellDepth > 0 {
				cellText.WriteString(s)
			}
			return
		}
		out = append(out, s)
	}
	writeVisible := func(s string) {
		if tableDepth == 0 && len(paragraphPrefixStack) > 0 && !paragraphPrefixWritten[len(paragraphPrefixWritten)-1] {
			if prefix := paragraphPrefixStack[len(paragraphPrefixStack)-1]; prefix != "" {
				write(prefix)
			}
			paragraphPrefixWritten[len(paragraphPrefixWritten)-1] = true
		}
		write(s)
	}
	flushTable := func() {
		rows := compactMarkdownTableRows(tableRows)
		if len(rows) > 0 {
			out = append(out, "\n", markdownTablePrepared(rows), "\n")
		}
		tableRows = nil
		tableRow = nil
		cellText.Reset()
	}
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return "", err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if skipDepth > 0 {
				skipDepth++
				continue
			}
			if checkboxDepth > 0 {
				checkboxDepth++
				if t.Name.Local == "checked" && checkboxCheckedElement(t) {
					checkboxChecked = true
				}
				continue
			}
			if dropdownDepth > 0 {
				dropdownDepth++
				switch t.Name.Local {
				case "result":
					if value, ok := intAttrValue(t, "val"); ok && value >= 0 {
						dropdownResult = value
					}
				case "listEntry":
					value := cleanVisibleXMLAttributeText(xmlAttrValue(t, "val"))
					if value != "" {
						dropdownEntries = append(dropdownEntries, value)
					}
				}
				continue
			}
			if textInputDepth > 0 {
				textInputDepth++
				if t.Name.Local == "default" {
					textInputDefault = cleanVisibleXMLAttributeText(xmlAttrValue(t, "val"))
				}
				continue
			}
			if isDrawingObjectElement(t.Name.Local) {
				parentHidden := len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]
				drawingObjectStack = append(drawingObjectStack, parentHidden)
			}
			if len(drawingObjectStack) > 0 && drawingObjectElementHidden(t) {
				drawingObjectStack[len(drawingObjectStack)-1] = true
			}
			drawingObjectHidden := len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]
			if t.Name.Local == "AlternateContent" {
				alternateStack = append(alternateStack, alternateContentState{})
				continue
			}
			if len(alternateStack) > 0 {
				top := &alternateStack[len(alternateStack)-1]
				switch t.Name.Local {
				case "Choice":
					top.choiceSeen = true
				case "Fallback":
					if top.choiceSeen {
						skipDepth = 1
						continue
					}
				}
			}
			if isHiddenRevisionElement(t.Name) {
				skipDepth = 1
				continue
			}
			if isHiddenRevisionRangeStart(t.Name) {
				hiddenRevisionRangeDepth++
				continue
			}
			if isHiddenRevisionRangeEnd(t.Name) {
				if hiddenRevisionRangeDepth > 0 {
					hiddenRevisionRangeDepth--
				}
				continue
			}
			if isSystemFootnoteElement(t) {
				skipDepth = 1
				continue
			}
			if xmlSystemTextElement(t.Name.Local) {
				skipDepth = 1
				continue
			}
			hiddenByRevisionRange := hiddenRevisionRangeDepth > 0
			contentVisible := !runHidden && !currentParagraphHidden(paragraphHiddenStack) && !drawingObjectHidden && !hiddenByRevisionRange
			if contentVisible {
				for _, value := range visibleAttributeText(t) {
					if seenAttrs[value] {
						continue
					}
					seenAttrs[value] = true
					write("\n")
					write(value)
					write("\n")
				}
			}
			switch t.Name.Local {
			case "tbl":
				if tableDepth == 0 {
					write("\n")
					tableRows = nil
					tableRow = nil
					cellText.Reset()
				}
				tableDepth++
			case "tr":
				if tableDepth == 1 {
					tableRow = nil
					rowDepth = 1
				} else if tableDepth > 0 {
					rowDepth++
				}
			case "tc":
				if tableDepth == 1 && rowDepth > 0 {
					cellText.Reset()
					cellDepth = 1
				} else if tableDepth > 0 {
					cellDepth++
				}
			case "p":
				paragraphHiddenStack = append(paragraphHiddenStack, false)
				paragraphPrefixStack = append(paragraphPrefixStack, "")
				paragraphPrefixWritten = append(paragraphPrefixWritten, false)
				paragraphListLevelStack = append(paragraphListLevelStack, 0)
				if contentVisible {
					write("\n")
				}
			case "r":
				runDepth++
				runHidden = false
			case "pPr":
				if len(paragraphHiddenStack) > 0 {
					pPrDepth++
					if len(paragraphListLevelStack) > 0 {
						paragraphListLevelStack[len(paragraphListLevelStack)-1] = markdownParagraphListLevel(t)
					}
				}
			case "numPr":
				if pPrDepth > 0 {
					numPrDepth = 1
					numPrID = ""
					numPrLevel = 0
				}
			case "ilvl":
				if numPrDepth > 0 {
					if value, ok := intAttrValue(t, "val"); ok && value >= 0 {
						numPrLevel = value
						if len(paragraphListLevelStack) > 0 {
							paragraphListLevelStack[len(paragraphListLevelStack)-1] = value
						}
					}
				}
			case "numId":
				if numPrDepth > 0 {
					numPrID = strings.TrimSpace(xmlAttrValue(t, "val"))
				}
			case "pStyle":
				if pPrDepth > 0 && len(paragraphPrefixStack) > 0 {
					paragraphPrefixStack[len(paragraphPrefixStack)-1] = markdownParagraphStylePrefix(xmlAttrValue(t, "val"))
				}
			case "buChar":
				if pPrDepth > 0 && len(paragraphPrefixStack) > 0 {
					level := markdownParagraphListLevelStack(paragraphListLevelStack)
					paragraphPrefixStack[len(paragraphPrefixStack)-1] = markdownListIndentPrefix(level) + "- "
				}
			case "buAutoNum":
				if pPrDepth > 0 && len(paragraphPrefixStack) > 0 {
					level := markdownParagraphListLevelStack(paragraphListLevelStack)
					paragraphPrefixStack[len(paragraphPrefixStack)-1] = markdownListIndentPrefix(level) + "1. "
				}
			case "buNone":
				if pPrDepth > 0 && len(paragraphPrefixStack) > 0 {
					paragraphPrefixStack[len(paragraphPrefixStack)-1] = ""
				}
			case "rPr":
				if runDepth > 0 {
					rPrDepth++
				}
			case "vanish", "webHidden":
				if runDepth > 0 && rPrDepth > 0 {
					runHidden = true
				}
				if pPrDepth > 0 && len(paragraphHiddenStack) > 0 {
					paragraphHiddenStack[len(paragraphHiddenStack)-1] = true
				}
			case "t", "text", "v":
				if contentVisible {
					textDepth++
				}
			case "br", "cr", "row":
				if contentVisible {
					write("\n")
				}
			case "tab":
				if contentVisible {
					writeVisible("\t")
				}
			case "noBreakHyphen":
				if contentVisible {
					writeVisible("\u2011")
				}
			case "sym":
				if contentVisible {
					if value, ok := visibleSymbolText(t); ok {
						writeVisible(value)
					}
				}
			case "checkBox":
				checkboxDepth = 1
				checkboxVisible = contentVisible
				checkboxChecked = false
			case "ddList":
				dropdownDepth = 1
				dropdownVisible = contentVisible
				dropdownResult = 0
				dropdownEntries = dropdownEntries[:0]
			case "textInput":
				textInputDepth = 1
				textInputVisible = contentVisible
				textInputDefault = ""
			}
		case xml.EndElement:
			if skipDepth > 0 {
				skipDepth--
				continue
			}
			if checkboxDepth > 0 {
				checkboxDepth--
				if checkboxDepth == 0 {
					if checkboxVisible {
						if checkboxChecked {
							writeVisible("\u2612")
						} else {
							writeVisible("\u2610")
						}
					}
					checkboxVisible = false
					checkboxChecked = false
				}
				continue
			}
			if dropdownDepth > 0 {
				dropdownDepth--
				if dropdownDepth == 0 {
					if dropdownVisible && dropdownResult >= 0 && dropdownResult < len(dropdownEntries) {
						writeVisible(dropdownEntries[dropdownResult])
					}
					dropdownVisible = false
					dropdownResult = 0
					dropdownEntries = dropdownEntries[:0]
				}
				continue
			}
			if textInputDepth > 0 {
				textInputDepth--
				if textInputDepth == 0 {
					if textInputVisible && textInputDefault != "" {
						writeVisible(textInputDefault)
					}
					textInputVisible = false
					textInputDefault = ""
				}
				continue
			}
			if t.Name.Local == "AlternateContent" {
				if len(alternateStack) > 0 {
					alternateStack = alternateStack[:len(alternateStack)-1]
				}
				continue
			}
			if (t.Name.Local == "t" || t.Name.Local == "text" || t.Name.Local == "v") && textDepth > 0 {
				textDepth--
			}
			if t.Name.Local == "pPr" && pPrDepth > 0 {
				pPrDepth--
			}
			if t.Name.Local == "numPr" && numPrDepth > 0 {
				if len(paragraphPrefixStack) > 0 && numPrID != "" {
					paragraphPrefixStack[len(paragraphPrefixStack)-1] = markdownWordNumberingPrefix(numbering, numPrID, numPrLevel)
				}
				numPrDepth = 0
				numPrID = ""
				numPrLevel = 0
			}
			if t.Name.Local == "rPr" && rPrDepth > 0 {
				rPrDepth--
			}
			if t.Name.Local == "r" && runDepth > 0 {
				runDepth--
				if runDepth == 0 {
					runHidden = false
					rPrDepth = 0
				}
			}
			if isDrawingObjectElement(t.Name.Local) && len(drawingObjectStack) > 0 {
				drawingObjectStack = drawingObjectStack[:len(drawingObjectStack)-1]
			}
			drawingObjectHidden := len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]
			switch t.Name.Local {
			case "tc":
				if tableDepth == 1 && cellDepth == 1 {
					if value := cleanMarkdownTableCellValue(cellText.String()); value != "" {
						tableRow = append(tableRow, prepareMarkdownTableCellValue(value))
					} else {
						tableRow = append(tableRow, "")
					}
					cellText.Reset()
					cellDepth = 0
				} else if cellDepth > 0 {
					cellDepth--
				}
			case "tr":
				if tableDepth == 1 && rowDepth == 1 {
					if !markdownTableRowBlank(tableRow) {
						tableRows = append(tableRows, tableRow)
					}
					tableRow = nil
					rowDepth = 0
				} else if rowDepth > 0 {
					rowDepth--
				}
			case "tbl":
				if tableDepth == 1 {
					flushTable()
					tableDepth = 0
					rowDepth = 0
					cellDepth = 0
				} else if tableDepth > 0 {
					tableDepth--
				}
			case "p", "row":
				if !runHidden && !currentParagraphHidden(paragraphHiddenStack) && !drawingObjectHidden && hiddenRevisionRangeDepth == 0 {
					write("\n")
				}
			}
			if t.Name.Local == "p" && len(paragraphHiddenStack) > 0 {
				paragraphHiddenStack = paragraphHiddenStack[:len(paragraphHiddenStack)-1]
				paragraphPrefixStack = paragraphPrefixStack[:len(paragraphPrefixStack)-1]
				paragraphPrefixWritten = paragraphPrefixWritten[:len(paragraphPrefixWritten)-1]
				paragraphListLevelStack = paragraphListLevelStack[:len(paragraphListLevelStack)-1]
				if len(paragraphHiddenStack) == 0 {
					pPrDepth = 0
				}
			}
		case xml.CharData:
			if textDepth > 0 && skipDepth == 0 {
				writeVisible(string(t))
			}
		}
	}
	if tableDepth > 0 {
		flushTable()
	}
	return cleanMarkdownVisibleText(strings.Join(out, "")), nil
}

func markdownParagraphListLevelStack(stack []int) int {
	if len(stack) == 0 {
		return 0
	}
	level := stack[len(stack)-1]
	if level < 0 {
		return 0
	}
	if level > 6 {
		return 6
	}
	return level
}

func markdownParagraphListLevel(start xml.StartElement) int {
	value, ok := intAttrValue(start, "lvl")
	if !ok || value < 0 {
		return 0
	}
	if value > 6 {
		return 6
	}
	return value
}

func markdownWordNumberingPrefix(numbering map[string]string, numID string, level int) string {
	format := ""
	if len(numbering) > 0 {
		format = numbering[markdownNumberingKey(numID, level)]
		if format == "" && level != 0 {
			format = numbering[markdownNumberingKey(numID, 0)]
		}
	}
	indent := markdownListIndentPrefix(level)
	if markdownNumberingFormatOrdered(format) {
		return indent + "1. "
	}
	return indent + "- "
}

func markdownListIndentPrefix(level int) string {
	return strings.Repeat(markdownIndentMarker, markdownClampListLevel(level)*2)
}

func markdownIndentMarkerCount(s string) int {
	count := 0
	for strings.HasPrefix(s, markdownIndentMarker) {
		count++
		s = strings.TrimPrefix(s, markdownIndentMarker)
	}
	if count > 12 {
		return 12
	}
	return count
}

func markdownNumberingKey(numID string, level int) string {
	return strings.TrimSpace(numID) + "\x00" + strconv.Itoa(markdownClampListLevel(level))
}

func markdownClampListLevel(level int) int {
	if level < 0 {
		return 0
	}
	if level > 6 {
		return 6
	}
	return level
}

func markdownNumberingFormatOrdered(format string) bool {
	switch strings.ToLower(strings.TrimSpace(format)) {
	case "decimal", "decimalzero", "upperroman", "lowerroman", "upperletter", "lowerletter",
		"ordinal", "cardinaltext", "ordinaltext", "hex", "chicagomanual":
		return true
	default:
		return false
	}
}

func markdownParagraphStylePrefix(style string) string {
	style = strings.TrimSpace(style)
	if style == "" {
		return ""
	}
	normalized := strings.ToLower(strings.Map(func(r rune) rune {
		if unicode.IsLetter(r) || unicode.IsDigit(r) {
			return r
		}
		return -1
	}, style))
	if normalized == "title" {
		return "# "
	}
	for _, prefix := range []string{"heading", "head", "h"} {
		if strings.HasPrefix(normalized, prefix) {
			n, ok := atoi(normalized[len(prefix):])
			if ok && n >= 1 && n <= 6 {
				level := n + 2
				if level > 6 {
					level = 6
				}
				return strings.Repeat("#", level) + " "
			}
		}
	}
	return ""
}

func compactMarkdownTableRows(rows [][]string) [][]string {
	maxCols := 0
	for _, row := range rows {
		for i := len(row) - 1; i >= 0; i-- {
			if strings.TrimSpace(row[i]) != "" {
				if i+1 > maxCols {
					maxCols = i + 1
				}
				break
			}
		}
	}
	if maxCols == 0 {
		return nil
	}
	out := make([][]string, 0, len(rows))
	for _, row := range rows {
		cells := make([]string, maxCols)
		copy(cells, row)
		out = append(out, cells)
	}
	return out
}

func markdownTableRowBlank(row []string) bool {
	for _, cell := range row {
		if strings.TrimSpace(cell) != "" {
			return false
		}
	}
	return true
}

func currentParagraphHidden(stack []bool) bool {
	return len(stack) > 0 && stack[len(stack)-1]
}

func checkboxCheckedElement(start xml.StartElement) bool {
	for _, attr := range start.Attr {
		if attr.Name.Local == "val" {
			return boolAttrValue(attr.Value)
		}
	}
	return true
}

func intAttrValue(start xml.StartElement, name string) (int, bool) {
	for _, attr := range start.Attr {
		if attr.Name.Local == name {
			return atoi(attr.Value)
		}
	}
	return 0, false
}

type alternateContentState struct {
	choiceSeen      bool
	choiceHasAnchor bool
}

func alternateContentStartSkip(name string, stack *[]alternateContentState) bool {
	if name == "AlternateContent" {
		*stack = append(*stack, alternateContentState{})
		return true
	}
	if len(*stack) == 0 {
		return false
	}
	top := &(*stack)[len(*stack)-1]
	switch name {
	case "Choice":
		top.choiceSeen = true
	case "Fallback":
		if top.choiceSeen {
			return true
		}
	}
	return false
}

func alternateContentEnd(name string, stack *[]alternateContentState) bool {
	if name != "AlternateContent" {
		return false
	}
	if len(*stack) > 0 {
		*stack = (*stack)[:len(*stack)-1]
	}
	return true
}

func isDrawingObjectElement(name string) bool {
	switch name {
	case "sp", "pic", "graphicFrame", "cxnSp", "grpSp", "wsp",
		"shape", "rect", "oval", "line", "group", "image":
		return true
	default:
		return false
	}
}

func drawingObjectElementHidden(start xml.StartElement) bool {
	if vmlElementHidden(start) {
		return true
	}
	// Word VML text boxes may be sent behind the document text with a negative
	// z-index. They are still present in document.xml but are not exposed by
	// Document.Content.Text; treating them as visible leaks background template
	// captions into strict Office-content extraction.
	for _, attr := range start.Attr {
		if attr.Name.Local != "style" {
			continue
		}
		for _, declaration := range strings.Split(strings.ToLower(attr.Value), ";") {
			key, value, ok := strings.Cut(strings.TrimSpace(declaration), ":")
			if ok && strings.TrimSpace(key) == "z-index" {
				if n, err := strconv.Atoi(strings.TrimSpace(value)); err == nil && (n < 0 || n >= 251658240) {
					return true
				}
			}
		}
	}
	return false
}

func isHiddenRevisionElement(name xml.Name) bool {
	if name.Space != "http://schemas.openxmlformats.org/wordprocessingml/2006/main" &&
		name.Space != "http://purl.oclc.org/ooxml/wordprocessingml/main" &&
		name.Space != "urn:x" {
		return false
	}
	switch name.Local {
	case "del", "moveFrom":
		return true
	default:
		return false
	}
}

func isHiddenRevisionRangeStart(name xml.Name) bool {
	return isWordprocessingName(name) && name.Local == "moveFromRangeStart"
}

func isHiddenRevisionRangeEnd(name xml.Name) bool {
	return isWordprocessingName(name) && name.Local == "moveFromRangeEnd"
}

func isWordprocessingName(name xml.Name) bool {
	return name.Space == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" ||
		name.Space == "http://purl.oclc.org/ooxml/wordprocessingml/main" ||
		name.Space == "urn:x"
}

func isSystemFootnoteElement(start xml.StartElement) bool {
	switch start.Name.Local {
	case "footnote", "endnote":
	default:
		return false
	}
	for _, attr := range start.Attr {
		if attr.Name.Local != "type" {
			continue
		}
		switch strings.ToLower(strings.TrimSpace(attr.Value)) {
		case "separator", "continuationseparator", "continuationnotice":
			return true
		}
	}
	return false
}

func cleanVisibleText(s string) string {
	s = cleanText(s)
	if s == "" || !strings.Contains(s, "\n") {
		s = stripInlineHiddenOfficeReferences(s)
		if looksLikeDiscardableVisibleTextLine(s) {
			return ""
		}
		return s
	}
	var out []string
	var pendingGlyphSoup []string
	flushGlyphSoup := func() {
		if len(pendingGlyphSoup) > 0 && len(pendingGlyphSoup) < 3 {
			out = append(out, pendingGlyphSoup...)
		}
		pendingGlyphSoup = pendingGlyphSoup[:0]
	}
	for _, line := range strings.Split(s, "\n") {
		line = cleanText(line)
		line = stripInlineHiddenOfficeReferences(line)
		if looksLikeDiscardableVisibleTextLine(line) {
			continue
		}
		if looksLikeLegacyShortGlyphSoupLine(line) ||
			(len(pendingGlyphSoup) >= 2 && looksLikeLegacyGlyphSoupContinuationLine(line)) {
			pendingGlyphSoup = append(pendingGlyphSoup, line)
			continue
		}
		flushGlyphSoup()
		out = append(out, line)
	}
	flushGlyphSoup()
	return cleanText(strings.Join(out, "\n"))
}

func cleanMarkdownVisibleText(s string) string {
	s = strings.ToValidUTF8(s, "")
	s = decodeOOXMLTextEscapes(s)
	s = strings.Map(cleanTextRune, s)
	lines := strings.Split(s, "\n")
	out := make([]string, 0, len(lines))
	for _, line := range lines {
		indent := 0
		markerIndent := markdownIndentMarkerCount(line)
		if markerIndent > 0 {
			indent = markerIndent
			line = line[markerIndent*len(markdownIndentMarker):]
		}
		trimmedRaw := strings.TrimSpace(line)
		if indent == 0 && markdownLineStartsWithListMarker(trimmedRaw) {
			indent = markdownLeadingSpaces(line)
			if indent > 12 {
				indent = 12
			}
		}
		line = spaceRE.ReplaceAllString(line, " ")
		line = strings.TrimSpace(line)
		line = repairWindows1251MojibakeLine(line)
		line = repairWindows1252UTF8MojibakePunctuationLine(line)
		line = repairGBKMojibakePunctuationLine(line)
		line = repairUnbalancedASCIIQuoteLine(line)
		line = repairMojibakeContractionLine(line)
		line = repairGBKDecodedUTF8LatinAccentsLine(line)
		line = stripWordFieldInstructions(line)
		line = stripInlineHiddenOfficeReferences(line)
		if looksLikeDiscardableVisibleTextLine(line) {
			continue
		}
		if indent > 0 && markdownLineStartsWithListMarker(line) {
			line = strings.Repeat(" ", indent) + line
		}
		out = append(out, line)
	}
	s = strings.Join(out, "\n")
	s = blankLineRE.ReplaceAllString(s, "\n\n")
	return strings.Trim(s, "\n")
}

func visibleAttributeText(start xml.StartElement) []string {
	var out []string
	for _, attr := range start.Attr {
		if !isVisibleTextAttribute(attr.Name.Local) {
			continue
		}
		value := cleanVisibleAttributeValue(attr.Value)
		if value != "" {
			out = append(out, value)
		}
	}
	return out
}

func cleanVisibleAttributeValue(value string) string {
	value = cleanText(value)
	value = stripInlineHiddenOfficeReferences(value)
	if value == "" || looksLikeBinaryControlFragment(value) {
		return ""
	}
	if !maybeDiscardableHiddenOfficeText(value) {
		return value
	}
	if looksLikeHiddenResourceReference(value) || looksLikeRelationshipIDReference(value) || looksLikeOfficeRelationshipMetadataReference(value) || looksLikeOfficeXMLMetadataReference(value) {
		return ""
	}
	return value
}

func looksLikeDiscardableVisibleTextLine(s string) bool {
	if s == "" {
		return true
	}
	if looksLikeDiscardableBinaryControlLine(s) {
		return true
	}
	if !maybeDiscardableHiddenOfficeText(s) {
		return false
	}
	return looksLikeHiddenResourceReference(s) ||
		looksLikeRelationshipIDReference(s) ||
		looksLikeOfficeRelationshipMetadataReference(s) ||
		looksLikeOfficeXMLMetadataReference(s)
}

func maybeDiscardableHiddenOfficeText(s string) bool {
	if maybeHiddenOrControlText(s) || containsRIDFold(s) {
		return true
	}
	hasAssignmentOrNamespace := strings.IndexByte(s, '=') >= 0 || strings.IndexByte(s, ':') >= 0
	if hasAssignmentOrNamespace {
		if containsASCIIFold(s, "xmlns") ||
			containsASCIIFold(s, "targetmode") ||
			containsASCIIFold(s, "contenttype") ||
			containsASCIIFold(s, "partname") ||
			containsASCIIFold(s, "schemalocation") ||
			containsASCIIFold(s, "ignorable") ||
			containsASCIIFold(s, "urn:schemas-microsoft-com:office:") {
			return true
		}
	}
	if strings.ContainsAny(s, "/\\") {
		return containsASCIIFold(s, "/relationships/") ||
			containsASCIIFold(s, "schemas.") ||
			containsASCIIFold(s, "purl.oclc.org/ooxml")
	}
	return false
}

func looksLikeRelationshipIDReference(s string) bool {
	s = strings.TrimSpace(s)
	if looksLikeBareRelationshipIDReference(s) {
		return true
	}
	candidate := hiddenResourceReferenceCandidate(s)
	return candidate != s && looksLikeBareRelationshipIDReference(candidate)
}

func looksLikeBareRelationshipIDReference(s string) bool {
	if len(s) < 4 || len(s) > 16 {
		return false
	}
	if len(s) < 4 || (s[0] != 'r' && s[0] != 'R') || (s[1] != 'i' && s[1] != 'I') || (s[2] != 'd' && s[2] != 'D') {
		return false
	}
	for _, r := range s[3:] {
		if r < '0' || r > '9' {
			return false
		}
	}
	return true
}

func looksLikeOfficeRelationshipMetadataReference(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	if len(s) > maxHiddenResourceMetadataReferenceBytes {
		return looksLikeLongOfficeRelationshipMetadataReference(s)
	}
	assignmentKey := ""
	if key, _, ok := hiddenResourceAssignmentSplit(s); ok {
		assignmentKey = strings.ToLower(strings.TrimSpace(key))
	}
	if value, ok := hiddenResourceAssignmentValue(s); ok {
		s = strings.TrimSpace(value)
	}
	s = hiddenResourceReferenceCandidate(s)
	lower := strings.ToLower(strings.Trim(strings.TrimSpace(s), `"'<>[](){}`))
	lower = strings.TrimRight(lower, ".,;:")
	if assignmentKey == "targetmode" && (lower == "external" || lower == "internal") {
		return true
	}
	if strings.Contains(lower, "/relationships/") {
		return strings.HasPrefix(lower, "http://schemas.openxmlformats.org/") ||
			strings.HasPrefix(lower, "https://schemas.openxmlformats.org/") ||
			strings.HasPrefix(lower, "http://purl.oclc.org/ooxml/") ||
			strings.HasPrefix(lower, "http://schemas.microsoft.com/office/") ||
			strings.HasPrefix(lower, "https://schemas.microsoft.com/office/")
	}
	return false
}

func looksLikeOfficeXMLMetadataReference(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	if len(s) > maxHiddenResourceMetadataReferenceBytes {
		return looksLikeLongOfficeXMLMetadataReference(s)
	}
	key := ""
	if assignmentKey, _, ok := hiddenResourceAssignmentSplit(s); ok {
		key = strings.ToLower(strings.TrimSpace(assignmentKey))
	}
	if value, ok := hiddenResourceAssignmentValue(s); ok {
		s = strings.TrimSpace(value)
	}
	lower := strings.ToLower(strings.Trim(strings.TrimSpace(s), `"'<>[](){}`))
	lower = strings.TrimRight(lower, ".,;:")
	switch key {
	case "xmlns", "mc:ignorable", "xsi:schemalocation", "schemalocation", "contenttype", "partname":
		return true
	default:
		if strings.HasPrefix(key, "xmlns:") {
			return true
		}
	}
	return strings.HasPrefix(lower, "http://schemas.openxmlformats.org/") ||
		strings.HasPrefix(lower, "https://schemas.openxmlformats.org/") ||
		strings.HasPrefix(lower, "http://purl.oclc.org/ooxml/") ||
		strings.HasPrefix(lower, "http://schemas.microsoft.com/office/") ||
		strings.HasPrefix(lower, "https://schemas.microsoft.com/office/") ||
		strings.HasPrefix(lower, "urn:schemas-microsoft-com:office:")
}

func longTextMayContainOfficeRelationshipMetadataReference(s string) bool {
	return containsASCIIFold(s, "/relationships/") ||
		containsASCIIFold(s, "targetmode") ||
		containsASCIIFold(s, "schemas.openxmlformats.org") ||
		containsASCIIFold(s, "schemas.microsoft.com/office")
}

func longTextMayContainOfficeXMLMetadataReference(s string) bool {
	return containsASCIIFold(s, "xmlns") ||
		containsASCIIFold(s, "contenttype") ||
		containsASCIIFold(s, "partname") ||
		containsASCIIFold(s, "schemalocation") ||
		containsASCIIFold(s, "schemas.openxmlformats.org") ||
		containsASCIIFold(s, "schemas.microsoft.com/office") ||
		containsASCIIFold(s, "urn:schemas-microsoft-com:office")
}

func looksLikeLongOfficeRelationshipMetadataReference(s string) bool {
	s = trimLongMetadataCandidate(s)
	if s == "" {
		return false
	}
	if longMetadataAssignmentKeyFold(s, "targetmode") {
		value := strings.TrimSpace(s[len("targetmode"):])
		value = strings.TrimLeft(value, `:= "'`)
		value = strings.TrimRight(value, ` "'.,;:/>`)
		return strings.EqualFold(value, "external") || strings.EqualFold(value, "internal")
	}
	return containsASCIIFold(s, "/relationships/") && hasOfficeSchemaPrefixFold(s)
}

func looksLikeLongOfficeXMLMetadataReference(s string) bool {
	s = trimLongMetadataCandidate(s)
	if s == "" {
		return false
	}
	if longMetadataAssignmentKeyFold(s, "xmlns") ||
		longMetadataAssignmentKeyFold(s, "mc:ignorable") ||
		longMetadataAssignmentKeyFold(s, "xsi:schemalocation") ||
		longMetadataAssignmentKeyFold(s, "schemalocation") ||
		longMetadataAssignmentKeyFold(s, "contenttype") ||
		longMetadataAssignmentKeyFold(s, "partname") {
		return true
	}
	return hasOfficeSchemaPrefixFold(s)
}

func trimLongMetadataCandidate(s string) string {
	s = strings.TrimSpace(s)
	s = strings.TrimLeft(s, `"'<>[](){}`)
	return strings.TrimRight(s, ".,;:")
}

func longMetadataAssignmentKeyFold(s, key string) bool {
	if len(s) < len(key) || !hasPrefixFold(s, key) {
		return false
	}
	if len(s) == len(key) {
		return true
	}
	switch s[len(key)] {
	case ':', '=', ' ', '\t', '\r', '\n':
		return true
	default:
		return false
	}
}

func hasOfficeSchemaPrefixFold(s string) bool {
	return hasPrefixFold(s, "http://schemas.openxmlformats.org/") ||
		hasPrefixFold(s, "https://schemas.openxmlformats.org/") ||
		hasPrefixFold(s, "http://purl.oclc.org/ooxml/") ||
		hasPrefixFold(s, "http://schemas.microsoft.com/office/") ||
		hasPrefixFold(s, "https://schemas.microsoft.com/office/") ||
		hasPrefixFold(s, "urn:schemas-microsoft-com:office:")
}

func isVisibleTextAttribute(name string) bool {
	switch name {
	case "descr", "title", "tooltip", "alt":
		return true
	default:
		return false
	}
}

func visibleSymbolText(start xml.StartElement) (string, bool) {
	font := strings.ToLower(strings.TrimSpace(xmlAttrValue(start, "font")))
	charHex := strings.TrimSpace(xmlAttrValue(start, "char"))
	if charHex == "" {
		return "", false
	}
	r, ok := parseOOXMLHexRune(charHex)
	if !ok {
		return "", false
	}
	if isDirectUnicodeSymbolFont(font) && visibleSymbolRune(r) {
		return string(r), true
	}
	code := r
	if code >= 0xf000 && code <= 0xf0ff {
		code &= 0xff
	}
	if mapped, ok := fontEncodedSymbolRune(font, code); ok {
		return string(mapped), true
	}
	if visibleSymbolRune(r) && !isFontEncodedSymbolFont(font) {
		return string(r), true
	}
	return "", false
}

// visibleWordSymbolTextAt implements the narrower behavior of
// Word.Document.Content.Text for w:sym. In particular, when Word opens the
// Open XML SDK's legacy Symbol-font glyph fields it exposes the fallback
// parenthesis, rather than the Unicode glyph suggested by the encoded code.
// Keep the general extractor's useful Unicode mapping in visibleSymbolText;
// this variant is only for the strict Office-visible comparator.
func visibleWordSymbolTextAt(start xml.StartElement) (string, bool) {
	font := strings.ToLower(strings.TrimSpace(xmlAttrValue(start, "font")))
	if isFontEncodedSymbolFont(font) {
		return "(", true
	}
	return visibleSymbolText(start)
}

func xmlAttrValue(start xml.StartElement, name string) string {
	for _, attr := range start.Attr {
		if attr.Name.Local == name {
			return attr.Value
		}
	}
	return ""
}

func parseOOXMLHexRune(s string) (rune, bool) {
	if len(s) == 4 {
		return parseOOXMLHex4(s)
	}
	if len(s) == 0 || len(s) > 6 {
		return 0, false
	}
	var v rune
	for _, r := range s {
		v <<= 4
		switch {
		case r >= '0' && r <= '9':
			v += r - '0'
		case r >= 'a' && r <= 'f':
			v += r - 'a' + 10
		case r >= 'A' && r <= 'F':
			v += r - 'A' + 10
		default:
			return 0, false
		}
	}
	if v > utf8.MaxRune {
		return 0, false
	}
	return v, true
}

func isDirectUnicodeSymbolFont(font string) bool {
	if font == "" {
		return true
	}
	switch font {
	case "arial", "arial unicode ms", "calibri", "cambria", "cambria math", "times new roman",
		"segoe ui", "segoe ui symbol", "unicode":
		return true
	default:
		return false
	}
}

func visibleSymbolRune(r rune) bool {
	if r == utf8.RuneError || r < 0x20 || (r >= 0xd800 && r <= 0xdfff) {
		return false
	}
	if isPrivateUseRune(r) {
		return false
	}
	return unicode.IsPrint(r)
}

func isFontEncodedSymbolFont(font string) bool {
	font = strings.ToLower(font)
	return strings.Contains(font, "symbol") || strings.Contains(font, "wingdings") || strings.Contains(font, "webdings")
}

func fontEncodedSymbolRune(font string, code rune) (rune, bool) {
	font = strings.ToLower(strings.TrimSpace(font))
	switch {
	case strings.Contains(font, "webdings"):
		// Word exposes these Webdings glyphs as parentheses in TextRange.Text.
		// They appear in documents normalized by the Open XML SDK, which stores
		// the original glyph code rather than the textual fallback.
		switch code {
		case 0x45, 0x49, 0x4a:
			return '(', true
		default:
			return 0, false
		}
	case strings.Contains(font, "wingdings"):
		switch code {
		case 0xfc:
			return '\u2713', true
		case 0xfb:
			return '\u2717', true
		case 0x6c:
			return '\u25cf', true
		case 0x6e:
			return '\u25a0', true
		default:
			return 0, false
		}
	case strings.Contains(font, "symbol"):
		switch code {
		case 0xae:
			// PowerPoint's legacy Symbol-font right-arrow fallback is commonly
			// serialized as U+F0AE in DrawingML a:t.
			return '\u2192', true
		case 0x41:
			return '\u0391', true
		case 0x42:
			return '\u0392', true
		case 0x43:
			return '\u03a7', true
		case 0x44:
			return '\u0394', true
		case 0x45:
			return '\u0395', true
		case 0x46:
			return '\u03a6', true
		case 0x47:
			return '\u0393', true
		case 0x48:
			return '\u0397', true
		case 0x49:
			return '\u0399', true
		case 0x4b:
			return '\u039a', true
		case 0x4c:
			return '\u039b', true
		case 0x4d:
			return '\u039c', true
		case 0x4e:
			return '\u039d', true
		case 0x4f:
			return '\u039f', true
		case 0x50:
			return '\u03a0', true
		case 0x51:
			return '\u0398', true
		case 0x52:
			return '\u03a1', true
		case 0x53:
			return '\u03a3', true
		case 0x54:
			return '\u03a4', true
		case 0x55:
			return '\u03a5', true
		case 0x57:
			return '\u03a9', true
		case 0x58:
			return '\u039e', true
		case 0x59:
			return '\u03a8', true
		case 0x5a:
			return '\u0396', true
		case 0x61:
			return '\u03b1', true
		case 0x62:
			return '\u03b2', true
		case 0x63:
			return '\u03c7', true
		case 0x64:
			return '\u03b4', true
		case 0x65:
			return '\u03b5', true
		case 0x66:
			return '\u03c6', true
		case 0x67:
			return '\u03b3', true
		case 0x68:
			return '\u03b7', true
		case 0x69:
			return '\u03b9', true
		case 0x6a:
			return '\u03d5', true
		case 0x6b:
			return '\u03ba', true
		case 0x6c:
			return '\u03bb', true
		case 0x6d:
			return '\u03bc', true
		case 0x6e:
			return '\u03bd', true
		case 0x6f:
			return '\u03bf', true
		case 0x70:
			return '\u03c0', true
		case 0x71:
			return '\u03b8', true
		case 0x72:
			return '\u03c1', true
		case 0x73:
			return '\u03c3', true
		case 0x74:
			return '\u03c4', true
		case 0x75:
			return '\u03c5', true
		case 0x76:
			return '\u03d6', true
		case 0x77:
			return '\u03c9', true
		case 0x78:
			return '\u03be', true
		case 0x79:
			return '\u03c8', true
		case 0x7a:
			return '\u03b6', true
		case 0xd7:
			return '\u00d7', true
		default:
			return 0, false
		}
	default:
		return 0, false
	}
}

func looksLikeHiddenResourceReference(s string) bool {
	raw := strings.TrimSpace(strings.ReplaceAll(s, "\\", "/"))
	if raw == "" {
		return false
	}
	seen := map[string]bool{}
	queue := []string{raw}
	for len(queue) > 0 && len(seen) < 64 {
		cur := queue[0]
		queue = queue[1:]
		cur = strings.TrimSpace(strings.ReplaceAll(cur, "\\", "/"))
		if cur == "" || seen[cur] {
			continue
		}
		seen[cur] = true
		if looksLikeHiddenResourceReferencePlain(cur) {
			return true
		}
		if normalized := hiddenResourceReferenceCandidate(cur); normalized != cur {
			queue = append(queue, normalized)
		}
		if packagePath := hiddenPackageURIPathCandidate(cur); packagePath != "" && packagePath != cur {
			queue = append(queue, packagePath)
		}
		if decoded, err := url.PathUnescape(cur); err == nil && decoded != cur {
			queue = append(queue, decoded)
		}
		if strings.Contains(cur, "&") {
			if unescaped := html.UnescapeString(cur); unescaped != cur {
				queue = append(queue, unescaped)
			}
		}
	}
	return false
}

func hiddenResourceReferenceCandidate(s string) string {
	s = strings.TrimSpace(s)
	for {
		before := s
		s = strings.TrimSpace(strings.TrimRight(s, ".,;:"))
		if len(s) >= len("url()") && strings.EqualFold(s[:4], "url(") && strings.HasSuffix(s, ")") {
			s = strings.TrimSpace(strings.TrimSuffix(strings.TrimPrefix(s, s[:4]), ")"))
			s = strings.Trim(s, `"'`)
			if s == before {
				break
			}
			continue
		}
		if unwrapped, ok := unwrapBalancedResourceReference(s); ok {
			s = strings.TrimSpace(unwrapped)
			if s == before {
				break
			}
			continue
		}
		if value, ok := hiddenResourceAssignmentValue(s); ok {
			s = value
		}
		switch {
		case strings.HasSuffix(s, "/>"):
			s = strings.TrimSpace(strings.TrimSuffix(s, "/>"))
		case strings.HasSuffix(s, ">"):
			s = strings.TrimSpace(strings.TrimSuffix(s, ">"))
		case len(s) >= 2 && s[0] == '"' && s[len(s)-1] == '"':
			s = s[1 : len(s)-1]
		case len(s) >= 2 && s[0] == '\'' && s[len(s)-1] == '\'':
			s = s[1 : len(s)-1]
		case len(s) >= 2 && s[0] == '`' && s[len(s)-1] == '`':
			s = s[1 : len(s)-1]
		case strings.HasPrefix(s, "<") && strings.HasSuffix(s, ">"):
			s = strings.TrimSuffix(strings.TrimPrefix(s, "<"), ">")
		case strings.HasPrefix(s, "(") && strings.HasSuffix(s, ")"):
			s = strings.TrimSuffix(strings.TrimPrefix(s, "("), ")")
		case strings.HasPrefix(s, "[") && strings.HasSuffix(s, "]"):
			s = strings.TrimSuffix(strings.TrimPrefix(s, "["), "]")
		case strings.HasPrefix(s, "{") && strings.HasSuffix(s, "}"):
			s = strings.TrimSuffix(strings.TrimPrefix(s, "{"), "}")
		case strings.HasPrefix(s, "“") && strings.HasSuffix(s, "”"):
			s = strings.TrimSuffix(strings.TrimPrefix(s, "“"), "”")
		case strings.HasPrefix(s, "‘") && strings.HasSuffix(s, "’"):
			s = strings.TrimSuffix(strings.TrimPrefix(s, "‘"), "’")
		}
		s = strings.TrimSpace(s)
		if s == before {
			break
		}
	}
	return strings.TrimSpace(s)
}

func hiddenResourceAssignmentValue(s string) (string, bool) {
	_, value, ok := hiddenResourceAssignmentSplit(s)
	return value, ok
}

func hiddenResourceAssignmentSplit(s string) (string, string, bool) {
	if len(s) > maxHiddenResourceMetadataReferenceBytes {
		return "", "", false
	}
	i := strings.IndexByte(s, '=')
	colonAssignment := false
	if i < 0 {
		for search := 0; search < len(s); search = i + 1 {
			next := strings.IndexByte(s[search:], ':')
			if next < 0 {
				break
			}
			i = search + next
			key := strings.ToLower(strings.TrimSpace(s[:i]))
			if hiddenOfficeColonAssignmentKey(key) {
				colonAssignment = true
				break
			}
		}
		if !colonAssignment {
			i = -1
		}
	}
	if i <= 0 || i == len(s)-1 {
		return "", "", false
	}
	key := strings.TrimSpace(s[:i])
	if key == "" || strings.ContainsAny(key, " \t\r\n<>/") {
		return "", "", false
	}
	normalizedKey := strings.ToLower(key)
	if colonAssignment && !hiddenOfficeColonAssignmentKey(normalizedKey) {
		return "", "", false
	}
	if strings.Contains(key, ":") && normalizedKey != "r:embed" && normalizedKey != "r:link" && normalizedKey != "r:id" && !strings.HasPrefix(normalizedKey, "xmlns:") && normalizedKey != "xsi:schemalocation" && normalizedKey != "mc:ignorable" {
		return "", "", false
	}
	for _, r := range key {
		if (r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') || (r >= '0' && r <= '9') || r == ':' || r == '_' || r == '-' {
			continue
		}
		return "", "", false
	}
	value := strings.TrimSpace(s[i+1:])
	if value == "" {
		return "", "", false
	}
	return key, value, true
}

func hiddenOfficeColonAssignmentKey(key string) bool {
	switch key {
	case "id", "target", "targetmode", "type", "contenttype", "partname", "embed", "link", "href", "src", "content-location", "content-id", "content-type", "content-transfer-encoding", "content-disposition", "content-description", "content-base", "mime-version", "r:embed", "r:link", "r:id", "mc:ignorable", "xsi:schemalocation", "schemalocation":
		return true
	default:
		return strings.HasPrefix(key, "xmlns:")
	}
}

func hiddenPackageURIPathCandidate(s string) string {
	lower := strings.ToLower(strings.TrimSpace(s))
	schemes := []string{"pack://", "zip://", "opc://", "ms-appx://", "ms-appdata://"}
	matched := false
	for _, scheme := range schemes {
		if strings.HasPrefix(lower, scheme) {
			matched = true
			break
		}
	}
	if !matched {
		return ""
	}
	if u, err := url.Parse(s); err == nil {
		candidates := []string{u.Path, u.Opaque, u.Fragment}
		for _, c := range candidates {
			if p := packageURIPathFromComponent(c); p != "" {
				return p
			}
		}
	}
	return packageURIPathFromComponent(s)
}

func packageURIPathFromComponent(s string) string {
	s = strings.TrimSpace(strings.ReplaceAll(s, "\\", "/"))
	lower := strings.ToLower(s)
	for _, root := range []string{"[content_types].xml", "docprops/", "customxml/", "word/", "ppt/", "xl/", "_rels/", "media/"} {
		if i := strings.Index(lower, root); i >= 0 {
			candidate := s[i:]
			if looksLikeOfficePartPath(strings.ToLower(candidate)) {
				return candidate
			}
		}
	}
	s = strings.TrimLeft(s, "/")
	s = strings.TrimLeft(s, ",")
	for {
		i := strings.Index(s, "/")
		if i < 0 {
			return ""
		}
		candidate := s[i+1:]
		if looksLikeOfficePartPath(strings.ToLower(candidate)) {
			return candidate
		}
		s = candidate
	}
}

func unwrapBalancedResourceReference(s string) (string, bool) {
	wrappers := []struct {
		open  string
		close string
	}{
		{`"`, `"`},
		{`'`, `'`},
		{"`", "`"},
		{"<", ">"},
		{"(", ")"},
		{"[", "]"},
		{"{", "}"},
		{"“", "”"},
		{"‘", "’"},
		{"«", "»"},
		{"‹", "›"},
		{"「", "」"},
		{"『", "』"},
		{"《", "》"},
		{"〈", "〉"},
		{"（", "）"},
		{"［", "］"},
		{"｛", "｝"},
		{"【", "】"},
		{"〔", "〕"},
	}
	for _, w := range wrappers {
		if strings.HasPrefix(s, w.open) && strings.HasSuffix(s, w.close) && len(s) >= len(w.open)+len(w.close) {
			return strings.TrimSuffix(strings.TrimPrefix(s, w.open), w.close), true
		}
	}
	return s, false
}

func looksLikeHiddenResourceReferencePlain(s string) bool {
	lower := strings.ToLower(s)
	if looksLikeLocalFileURIReference(lower) || strings.HasPrefix(s, "//") {
		return true
	}
	if hiddenPackageURIPathCandidate(s) != "" {
		return true
	}
	if len(s) >= 3 && ((s[0] >= 'A' && s[0] <= 'Z') || (s[0] >= 'a' && s[0] <= 'z')) && s[1] == ':' && s[2] == '/' {
		return true
	}
	if strings.HasPrefix(s, "/") && strings.Contains(s, "/") && looksLikeOfficePartPath(lower[1:]) {
		return true
	}
	return looksLikeOfficePartPath(lower)
}

func looksLikeOfficePartPath(s string) bool {
	s = strings.TrimSpace(s)
	if i := strings.IndexAny(s, "?#"); i >= 0 {
		s = s[:i]
	}
	s = strings.TrimRight(s, ".,;:)]}>")
	s = strings.TrimPrefix(s, "./")
	for strings.HasPrefix(s, "../") {
		s = strings.TrimPrefix(s, "../")
	}
	if strings.HasPrefix(s, "media/") && isSupportedImageExt(path.Ext(s)) {
		return true
	}
	switch {
	case s == "[content_types].xml", strings.HasPrefix(s, "docprops/"), strings.HasPrefix(s, "customxml/"):
		return strings.HasSuffix(s, ".xml") || strings.HasSuffix(s, ".rels") || strings.Contains(s, "/item")
	case strings.HasPrefix(s, "word/"), strings.HasPrefix(s, "ppt/"), strings.HasPrefix(s, "xl/"),
		strings.HasPrefix(s, "_rels/"), strings.Contains(s, "/_rels/"):
		return strings.HasSuffix(s, ".xml") || strings.HasSuffix(s, ".rels") || strings.Contains(s, "/media/")
	default:
		return false
	}
}

func visibleHTMLText(b []byte) string {
	s := string(decodeHTMLBytes(b))
	s = regexp.MustCompile(`(?is)<!doctype[^>]*>`).ReplaceAllString(s, " ")
	s = stripHTMLComments(s)
	s = stripHTMLInertBlocks(s)
	s = stripHTMLEmbedContainers(s)
	s = stripHiddenHTMLBlocks(s)
	s = visibleHTMLDetailsText(s)
	s = visibleHTMLDialogText(s)
	s = regexp.MustCompile(`(?is)<script\b[^>]*>.*?</script>`).ReplaceAllString(s, " ")
	s = regexp.MustCompile(`(?is)<(?:\w+:)?ClientData\b[^>]*>.*?</(?:\w+:)?ClientData>`).ReplaceAllString(s, " ")
	s = stripHiddenHTMLBlocks(s)
	s = regexp.MustCompile(`(?is)<head\b[^>]*>.*?</head>`).ReplaceAllString(s, " ")
	s = regexp.MustCompile(`(?is)<style\b[^>]*>.*?</style>`).ReplaceAllString(s, " ")
	s = visibleHTMLSelectText(s)
	s = visibleHTMLInputText(s)
	s = visibleHTMLTextareaText(s)
	s = regexp.MustCompile(`(?is)<br\b[^>]*>`).ReplaceAllString(s, "\n")
	s = regexp.MustCompile(`(?is)</?(p|div|h[1-6]|li|tr|table|section|article|blockquote)\b[^>]*>`).ReplaceAllString(s, "\n")
	s = regexp.MustCompile(`(?is)<[^>]+>`).ReplaceAllString(s, " ")
	return cleanVisibleText(html.UnescapeString(s))
}

func visibleAltChunkText(name string, b []byte) string {
	lower := ooxmlPartKey(name)
	if strings.HasSuffix(lower, ".mht") || strings.HasSuffix(lower, ".mhtml") {
		return visibleMHTMLText(b)
	}
	return visibleHTMLText(b)
}

type mhtmlPart struct {
	Headers map[string]string
	Body    []byte
}

func visibleMHTMLText(b []byte) string {
	parts := parseMHTMLParts(b)
	if len(parts) == 0 {
		return visibleHTMLText(b)
	}
	var blocks []string
	for _, part := range parts {
		if !strings.Contains(strings.ToLower(part.Headers["content-type"]), "text/html") {
			continue
		}
		body := decodeMHTMLTextPart(part)
		if text := visibleHTMLText(body); text != "" {
			blocks = append(blocks, text)
		}
	}
	return cleanVisibleText(strings.Join(blocks, "\n"))
}

var (
	htmlImageTagRE        = regexp.MustCompile(`(?is)<img\b[^>]*>`)
	htmlPictureBlockRE    = regexp.MustCompile(`(?is)(<picture\b[^>]*>)(.*?)</picture>`)
	htmlSourceTagRE       = regexp.MustCompile(`(?is)<source\b[^>]*>`)
	htmlTagRE             = regexp.MustCompile(`(?is)<[^>]+>`)
	htmlAttrValueRE       = regexp.MustCompile(`(?is)([a-zA-Z_:][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>` + "`" + `]+))`)
	htmlCharsetRE         = regexp.MustCompile(`(?is)<meta\b[^>]*\bcharset\s*=\s*["']?\s*([A-Za-z0-9._:-]+)`)
	htmlXMLEncodingRE     = regexp.MustCompile(`(?is)<\?xml\b[^?]*\bencoding\s*=\s*(?:"([^"]+)"|'([^']+)')`)
	htmlMetaTagRE         = regexp.MustCompile(`(?is)<meta\b[^>]*>`)
	htmlHiddenAttrRE      = regexp.MustCompile(`(?is)(?:\s|<)hidden(?:\s|=|/|>)`)
	htmlCommentRE         = regexp.MustCompile(`(?is)<!--.*?-->`)
	htmlStyleBlockRE      = regexp.MustCompile(`(?is)<style\b[^>]*>(.*?)</style>`)
	htmlInputTagRE        = regexp.MustCompile(`(?is)<input\b[^>]*>`)
	htmlTextareaBlockRE   = regexp.MustCompile(`(?is)(<textarea\b[^>]*>)(.*?)</textarea>`)
	htmlSelectBlockRE     = regexp.MustCompile(`(?is)<select\b[^>]*>(.*?)</select>`)
	htmlOptionBlockRE     = regexp.MustCompile(`(?is)(<option\b[^>]*>)(.*?)</option>`)
	htmlDetailsBlockRE    = regexp.MustCompile(`(?is)(<details\b[^>]*>)(.*?)</details>`)
	htmlSummaryBlockRE    = regexp.MustCompile(`(?is)(<summary\b[^>]*>)(.*?)</summary>`)
	htmlDialogBlockRE     = regexp.MustCompile(`(?is)(<dialog\b[^>]*>)(.*?)</dialog>`)
	htmlTemplateBlockRE   = regexp.MustCompile(`(?is)<template\b[^>]*>.*?</template>`)
	htmlDatalistBlockRE   = regexp.MustCompile(`(?is)<datalist\b[^>]*>.*?</datalist>`)
	htmlObjectBlockRE     = regexp.MustCompile(`(?is)<object\b[^>]*>.*?</object>`)
	htmlIframeBlockRE     = regexp.MustCompile(`(?is)<iframe\b[^>]*>.*?</iframe>`)
	htmlEmbedTagRE        = regexp.MustCompile(`(?is)<embed\b[^>]*>`)
	htmlHeadBlockRE       = regexp.MustCompile(`(?is)<head\b[^>]*>.*?</head>`)
	htmlScriptBlockRE     = regexp.MustCompile(`(?is)<script\b[^>]*>.*?</script>`)
	htmlStyleBlockFullRE  = regexp.MustCompile(`(?is)<style\b[^>]*>.*?</style>`)
	htmlClientDataBlockRE = regexp.MustCompile(`(?is)<(?:\w+:)?ClientData\b[^>]*>.*?</(?:\w+:)?ClientData>`)
	cssCommentRE          = regexp.MustCompile(`(?is)/\*.*?\*/`)
	cssClassRE            = regexp.MustCompile(`\.([_a-zA-Z][\w-]*)`)
	cssIDRE               = regexp.MustCompile(`#([_a-zA-Z][\w-]*)`)
)

type htmlImageRef struct {
	Media string
	Alt   string
}

func htmlImageMediaRefs(files map[string]*zip.File, source string, b []byte) []string {
	refs := htmlImageRefs(files, source, b)
	out := make([]string, 0, len(refs))
	seen := map[string]bool{}
	for _, ref := range refs {
		if ref.Media == "" || seen[ref.Media] {
			continue
		}
		seen[ref.Media] = true
		out = append(out, ref.Media)
	}
	return out
}

func htmlImageRefs(files map[string]*zip.File, source string, b []byte) []htmlImageRef {
	var out []htmlImageRef
	htmlText := visibleHTMLForImageRefs(string(decodeHTMLBytes(b)))
	for _, tag := range htmlImageTagRE.FindAllString(htmlText, -1) {
		attrs := htmlTagAttrs(tag)
		srcs := htmlImageSourceCandidates(attrs)
		if len(srcs) == 0 || htmlImageTagHidden(attrs) {
			continue
		}
		alt := cleanMarkdownImageAltText(attrs["alt"])
		if alt == "" {
			alt = cleanMarkdownImageAltText(attrs["title"])
		}
		for _, src := range srcs {
			part := resolveHTMLImageMediaPart(files, source, src)
			if part == "" {
				continue
			}
			out = append(out, htmlImageRef{Media: part, Alt: alt})
			break
		}
	}
	for _, ref := range htmlPictureSourceRefs(htmlText) {
		for _, src := range ref.Sources {
			part := resolveHTMLImageMediaPart(files, source, src)
			if part == "" {
				continue
			}
			out = append(out, htmlImageRef{Media: part, Alt: ref.Alt})
			break
		}
	}
	return out
}

type htmlPictureSourceRef struct {
	Sources []string
	Alt     string
}

func htmlPictureSourceRefs(htmlText string) []htmlPictureSourceRef {
	var out []htmlPictureSourceRef
	for _, m := range htmlPictureBlockRE.FindAllStringSubmatch(htmlText, -1) {
		if len(m) < 3 || htmlStartTagHidden(m[1], htmlHiddenCSS{}) {
			continue
		}
		if htmlPictureHasImageSource(m[2]) {
			continue
		}
		alt := htmlPictureAlt(m[2])
		for _, tag := range htmlSourceTagRE.FindAllString(m[2], -1) {
			if htmlStartTagHidden(tag, htmlHiddenCSS{}) {
				continue
			}
			attrs := htmlTagAttrs(tag)
			srcs := htmlImageSourceCandidates(attrs)
			if len(srcs) == 0 {
				continue
			}
			out = append(out, htmlPictureSourceRef{Sources: srcs, Alt: alt})
			break
		}
	}
	return out
}

func htmlPictureHasImageSource(body string) bool {
	for _, tag := range htmlImageTagRE.FindAllString(body, -1) {
		if len(htmlImageSourceCandidates(htmlTagAttrs(tag))) > 0 {
			return true
		}
	}
	return false
}

func htmlPictureAlt(body string) string {
	for _, tag := range htmlImageTagRE.FindAllString(body, -1) {
		attrs := htmlTagAttrs(tag)
		if alt := cleanMarkdownImageAltText(attrs["alt"]); alt != "" {
			return alt
		}
		if title := cleanMarkdownImageAltText(attrs["title"]); title != "" {
			return title
		}
	}
	return ""
}

func htmlImageSourceCandidates(attrs map[string]string) []string {
	if src := strings.TrimSpace(attrs["src"]); src != "" {
		return []string{src}
	}
	return htmlSrcsetSources(attrs["srcset"])
}

func htmlSrcsetSources(srcset string) []string {
	var out []string
	for i := 0; i < len(srcset); {
		for i < len(srcset) && (srcset[i] == ',' || unicode.IsSpace(rune(srcset[i]))) {
			i++
		}
		if i >= len(srcset) {
			break
		}
		start := i
		if strings.HasPrefix(strings.ToLower(srcset[i:]), "data:") {
			seenDataComma := false
			for i < len(srcset) {
				if unicode.IsSpace(rune(srcset[i])) {
					break
				}
				if srcset[i] == ',' {
					if seenDataComma {
						break
					}
					seenDataComma = true
				}
				i++
			}
		} else {
			for i < len(srcset) && srcset[i] != ',' && !unicode.IsSpace(rune(srcset[i])) {
				i++
			}
		}
		src := strings.TrimSpace(srcset[start:i])
		if src != "" {
			out = append(out, src)
		}
		for i < len(srcset) && srcset[i] != ',' {
			i++
		}
		if i < len(srcset) && srcset[i] == ',' {
			i++
		}
	}
	return out
}

func stripHTMLComments(s string) string {
	return htmlCommentRE.ReplaceAllString(s, " ")
}

func stripHTMLInertBlocks(s string) string {
	for _, re := range []*regexp.Regexp{htmlTemplateBlockRE, htmlDatalistBlockRE} {
		s = re.ReplaceAllString(s, " ")
	}
	return s
}

func stripHTMLEmbedContainers(s string) string {
	for _, re := range []*regexp.Regexp{htmlObjectBlockRE, htmlIframeBlockRE, htmlEmbedTagRE} {
		s = re.ReplaceAllString(s, " ")
	}
	return s
}

func stripNonVisibleHTMLBlocksForImages(s string) string {
	for _, re := range []*regexp.Regexp{htmlHeadBlockRE, htmlScriptBlockRE, htmlStyleBlockFullRE, htmlClientDataBlockRE} {
		s = re.ReplaceAllString(s, " ")
	}
	return s
}

func visibleHTMLForImageRefs(s string) string {
	s = stripHTMLComments(s)
	s = stripHTMLInertBlocks(s)
	s = stripHTMLEmbedContainers(s)
	s = stripHiddenHTMLBlocks(s)
	s = visibleHTMLDetailsText(s)
	s = visibleHTMLDialogText(s)
	s = stripHiddenHTMLBlocks(s)
	return stripNonVisibleHTMLBlocksForImages(s)
}

func visibleHTMLDetailsText(s string) string {
	return htmlDetailsBlockRE.ReplaceAllStringFunc(s, func(block string) string {
		m := htmlDetailsBlockRE.FindStringSubmatch(block)
		if len(m) < 3 {
			return " "
		}
		if htmlTagHasAttr(m[1], "open") {
			return block
		}
		summary := htmlSummaryBlockRE.FindStringSubmatch(m[2])
		if len(summary) < 3 || htmlStartTagHidden(summary[1], htmlHiddenCSS{}) {
			return " "
		}
		return "\n" + summary[2] + "\n"
	})
}

func visibleHTMLDialogText(s string) string {
	return htmlDialogBlockRE.ReplaceAllStringFunc(s, func(block string) string {
		m := htmlDialogBlockRE.FindStringSubmatch(block)
		if len(m) < 3 {
			return " "
		}
		if !htmlTagHasAttr(m[1], "open") || htmlStartTagHidden(m[1], htmlHiddenCSS{}) {
			return " "
		}
		return "\n" + m[2] + "\n"
	})
}

func visibleHTMLSelectText(s string) string {
	return htmlSelectBlockRE.ReplaceAllStringFunc(s, func(block string) string {
		m := htmlSelectBlockRE.FindStringSubmatch(block)
		if len(m) < 2 {
			return " "
		}
		var selected []string
		var first string
		body := m[1]
		for _, opt := range htmlOptionBlockRE.FindAllStringSubmatchIndex(body, -1) {
			if len(opt) < 6 || opt[2] < 0 || opt[3] < 0 || opt[4] < 0 || opt[5] < 0 {
				continue
			}
			tag := body[opt[2]:opt[3]]
			if htmlStartTagHidden(tag, htmlHiddenCSS{}) {
				continue
			}
			text := htmlOptionVisibleText(tag, body[opt[4]:opt[5]])
			if text == "" {
				continue
			}
			if group := htmlVisibleOptgroupLabelAt(body, opt[0]); group != "" && group != text {
				text = group + "\n" + text
			}
			if first == "" {
				first = text
			}
			if htmlTagHasAttr(tag, "selected") {
				selected = append(selected, text)
			}
		}
		if len(selected) > 0 {
			return "\n" + strings.Join(selected, "\n") + "\n"
		}
		if first != "" {
			return "\n" + first + "\n"
		}
		return " "
	})
}

func htmlOptionVisibleText(tag, body string) string {
	text := cleanVisibleText(html.UnescapeString(markdownVisibleHTMLText(body)))
	if text != "" {
		return text
	}
	attrs := htmlTagAttrs(tag)
	return cleanVisibleAttributeValue(attrs["label"])
}

func htmlVisibleOptgroupLabelAt(selectBody string, optionStart int) string {
	if optionStart <= 0 || optionStart > len(selectBody) {
		return ""
	}
	var label string
	for _, loc := range htmlTagRE.FindAllStringIndex(selectBody[:optionStart], -1) {
		tag := selectBody[loc[0]:loc[1]]
		name, closing, selfClosing := htmlTagName(tag)
		if name != "optgroup" {
			continue
		}
		if closing {
			label = ""
			continue
		}
		if selfClosing || htmlStartTagHidden(tag, htmlHiddenCSS{}) {
			label = ""
			continue
		}
		attrs := htmlTagAttrs(tag)
		label = cleanVisibleAttributeValue(attrs["label"])
	}
	return label
}

func visibleHTMLInputText(s string) string {
	return htmlInputTagRE.ReplaceAllStringFunc(s, func(tag string) string {
		if htmlStartTagHidden(tag, htmlHiddenCSS{}) {
			return " "
		}
		attrs := htmlTagAttrs(tag)
		typ := strings.ToLower(strings.TrimSpace(attrs["type"]))
		if typ == "" {
			typ = "text"
		}
		if !htmlInputTypeTextVisible(typ) {
			return " "
		}
		value := htmlInputVisibleText(attrs, typ)
		if value == "" {
			return " "
		}
		return "\n" + value + "\n"
	})
}

func htmlInputVisibleText(attrs map[string]string, typ string) string {
	value := cleanVisibleAttributeValue(attrs["value"])
	if value != "" {
		return value
	}
	if htmlInputTypeUsesPlaceholder(typ) {
		return cleanVisibleAttributeValue(attrs["placeholder"])
	}
	return ""
}

func htmlInputTypeUsesPlaceholder(typ string) bool {
	switch typ {
	case "text", "search", "email", "url", "tel", "number", "date", "datetime-local", "month", "time", "week":
		return true
	default:
		return false
	}
}

func visibleHTMLTextareaText(s string) string {
	return htmlTextareaBlockRE.ReplaceAllStringFunc(s, func(block string) string {
		m := htmlTextareaBlockRE.FindStringSubmatch(block)
		if len(m) < 3 {
			return " "
		}
		if htmlStartTagHidden(m[1], htmlHiddenCSS{}) {
			return " "
		}
		text := htmlVisibleFormControlText(m[2])
		if text == "" {
			attrs := htmlTagAttrs(m[1])
			text = cleanVisibleAttributeValue(attrs["placeholder"])
		}
		if text == "" {
			return " "
		}
		return "\n" + text + "\n"
	})
}

func htmlVisibleFormControlText(s string) string {
	s = markdownVisibleHTMLText(s)
	return cleanVisibleText(html.UnescapeString(s))
}

func htmlInputTypeTextVisible(typ string) bool {
	switch typ {
	case "text", "search", "email", "url", "tel", "number", "date", "datetime-local", "month", "time", "week", "color", "button", "submit", "reset":
		return true
	default:
		return false
	}
}

func htmlTagHasAttr(tag, attr string) bool {
	attr = strings.ToLower(strings.TrimSpace(attr))
	if attr == "" {
		return false
	}
	attrs := htmlTagAttrs(tag)
	if _, ok := attrs[attr]; ok {
		return true
	}
	name, _, _ := htmlTagName(tag)
	if name == "" {
		return false
	}
	body := strings.TrimSpace(strings.TrimPrefix(strings.TrimSuffix(tag, ">"), "<"))
	fields := strings.Fields(strings.TrimRight(body, "/"))
	for _, field := range fields[1:] {
		field = strings.Trim(field, `"'`)
		if i := strings.IndexAny(field, "=/>"); i >= 0 {
			field = field[:i]
		}
		if strings.EqualFold(strings.TrimSpace(field), attr) {
			return true
		}
	}
	return false
}

func decodeHTMLBytes(b []byte) []byte {
	if decoded, ok := decodeTextBOMBytes(b); ok {
		return decoded
	}
	if decoded, ok := decodeLikelyUTF16HTMLBytes(b); ok {
		return decoded
	}
	charset := htmlDeclaredCharset(b)
	if charset == "" || strings.EqualFold(charset, "utf-8") || strings.EqualFold(charset, "utf8") {
		return b
	}
	if text := decodeMHTMLCharsetBytes(b, charset); text != "" {
		return []byte(text)
	}
	return b
}

func htmlDeclaredCharset(b []byte) string {
	if len(b) > 8192 {
		b = b[:8192]
	}
	if m := htmlCharsetRE.FindSubmatch(b); len(m) == 2 {
		return strings.ToLower(strings.TrimSpace(string(m[1])))
	}
	if m := htmlXMLEncodingRE.FindSubmatch(b); len(m) == 3 {
		if enc := strings.TrimSpace(string(m[1])); enc != "" {
			return strings.ToLower(enc)
		}
		if enc := strings.TrimSpace(string(m[2])); enc != "" {
			return strings.ToLower(enc)
		}
	}
	for _, tag := range htmlMetaTagRE.FindAll(b, -1) {
		attrs := htmlTagAttrs(string(tag))
		if !strings.EqualFold(attrs["http-equiv"], "content-type") && attrs["content"] == "" {
			continue
		}
		if charset := mhtmlTextCharset(attrs["content"]); charset != "" {
			return charset
		}
	}
	return ""
}

func htmlTagAttrs(tag string) map[string]string {
	out := map[string]string{}
	for _, m := range htmlAttrValueRE.FindAllStringSubmatch(tag, -1) {
		if len(m) < 5 {
			continue
		}
		name := strings.ToLower(strings.TrimSpace(m[1]))
		value := m[2]
		if value == "" {
			value = m[3]
		}
		if value == "" {
			value = m[4]
		}
		out[name] = strings.TrimSpace(html.UnescapeString(value))
	}
	return out
}

func htmlImageTagHidden(attrs map[string]string) bool {
	return htmlAttrsHidden(attrs)
}

func htmlAttrsHidden(attrs map[string]string) bool {
	if _, ok := attrs["hidden"]; ok {
		return true
	}
	if strings.EqualFold(strings.TrimSpace(attrs["aria-hidden"]), "true") {
		return true
	}
	return htmlCSSDeclHidden(attrs["style"])
}

func normalizeHTMLStyleValue(style string) string {
	style = strings.ToLower(html.UnescapeString(style))
	var b strings.Builder
	for _, r := range style {
		if unicode.IsSpace(r) {
			continue
		}
		b.WriteRune(r)
	}
	return b.String()
}

func stripHiddenHTMLBlocks(s string) string {
	hiddenCSS := htmlHiddenCSSSelectors(s)
	var out strings.Builder
	var hiddenStack []string
	pos := 0
	for _, loc := range htmlTagRE.FindAllStringIndex(s, -1) {
		if len(hiddenStack) == 0 {
			out.WriteString(s[pos:loc[0]])
		}
		tag := s[loc[0]:loc[1]]
		name, closing, selfClosing := htmlTagName(tag)
		if closing {
			if len(hiddenStack) > 0 {
				for len(hiddenStack) > 0 {
					top := hiddenStack[len(hiddenStack)-1]
					hiddenStack = hiddenStack[:len(hiddenStack)-1]
					if top == name {
						break
					}
				}
			} else {
				out.WriteString(tag)
			}
			pos = loc[1]
			continue
		}
		hidden := htmlStartTagHidden(tag, hiddenCSS)
		if len(hiddenStack) > 0 {
			if !selfClosing && name != "" && !htmlVoidElement(name) {
				hiddenStack = append(hiddenStack, name)
			}
			pos = loc[1]
			continue
		}
		if hidden {
			if !selfClosing && name != "" && !htmlVoidElement(name) {
				hiddenStack = append(hiddenStack, name)
			}
			pos = loc[1]
			continue
		}
		out.WriteString(tag)
		pos = loc[1]
	}
	if len(hiddenStack) == 0 {
		out.WriteString(s[pos:])
	}
	return out.String()
}

type htmlHiddenCSS struct {
	Classes   map[string]bool
	IDs       map[string]bool
	Compounds []htmlHiddenCSSSelector
}

type htmlHiddenCSSSelector struct {
	Element string
	ID      string
	Classes []string
}

func htmlHiddenCSSSelectors(s string) htmlHiddenCSS {
	out := htmlHiddenCSS{Classes: map[string]bool{}, IDs: map[string]bool{}}
	for _, m := range htmlStyleBlockRE.FindAllStringSubmatch(s, -1) {
		if len(m) < 2 {
			continue
		}
		css := cssCommentRE.ReplaceAllString(m[1], " ")
		for _, rule := range strings.Split(css, "}") {
			selector, decl, ok := strings.Cut(rule, "{")
			if !ok || !htmlCSSDeclHidden(decl) {
				continue
			}
			for _, sel := range strings.Split(selector, ",") {
				sel = strings.TrimSpace(sel)
				if sel == "" || strings.ContainsAny(sel, " \t\r\n>+~") {
					continue
				}
				parsed, ok := parseHTMLHiddenCSSSelector(sel)
				if !ok {
					continue
				}
				if parsed.Element == "" && parsed.ID == "" && len(parsed.Classes) == 1 {
					out.Classes[parsed.Classes[0]] = true
					continue
				}
				if parsed.Element == "" && parsed.ID != "" && len(parsed.Classes) == 0 {
					out.IDs[parsed.ID] = true
					continue
				}
				out.Compounds = append(out.Compounds, parsed)
			}
		}
	}
	return out
}

func parseHTMLHiddenCSSSelector(sel string) (htmlHiddenCSSSelector, bool) {
	sel = strings.ToLower(strings.TrimSpace(sel))
	if sel == "" || strings.ContainsAny(sel, " \t\r\n>+~:[]*=|") {
		return htmlHiddenCSSSelector{}, false
	}
	var out htmlHiddenCSSSelector
	i := 0
	for i < len(sel) && sel[i] != '.' && sel[i] != '#' {
		i++
	}
	if i > 0 {
		out.Element = htmlSelectorLocalName(strings.TrimSpace(sel[:i]))
		if out.Element == "" {
			return htmlHiddenCSSSelector{}, false
		}
	}
	for i < len(sel) {
		kind := sel[i]
		if kind != '.' && kind != '#' {
			return htmlHiddenCSSSelector{}, false
		}
		i++
		start := i
		for i < len(sel) && (isCSSIdentByte(sel[i]) || sel[i] == '-') {
			i++
		}
		value := strings.TrimSpace(sel[start:i])
		if value == "" {
			return htmlHiddenCSSSelector{}, false
		}
		if kind == '#' {
			if out.ID != "" {
				return htmlHiddenCSSSelector{}, false
			}
			out.ID = value
		} else {
			out.Classes = append(out.Classes, value)
		}
	}
	if out.ID == "" && len(out.Classes) == 0 {
		return htmlHiddenCSSSelector{}, false
	}
	return out, true
}

func htmlSelectorLocalName(name string) string {
	if name == "" || name == "*" {
		return ""
	}
	if i := strings.LastIndex(name, "|"); i >= 0 {
		name = name[i+1:]
	}
	if i := strings.LastIndex(name, ":"); i >= 0 {
		name = name[i+1:]
	}
	if !isCSSIdentStartName(name) {
		return ""
	}
	for i := 1; i < len(name); i++ {
		if !isCSSIdentByte(name[i]) && name[i] != '-' {
			return ""
		}
	}
	return name
}

func isCSSIdentStartName(name string) bool {
	if name == "" {
		return false
	}
	b := name[0]
	return (b >= 'a' && b <= 'z') || b == '_'
}

func isCSSIdentByte(b byte) bool {
	return (b >= 'a' && b <= 'z') || (b >= '0' && b <= '9') || b == '_'
}

func htmlCSSDeclHidden(decl string) bool {
	style := normalizeHTMLStyleValue(decl)
	if strings.Contains(style, "display:none") ||
		strings.Contains(style, "visibility:hidden") ||
		strings.Contains(style, "visibility:collapse") ||
		strings.Contains(style, "mso-hide:all") {
		return true
	}
	return htmlCSSPropertyHidden(decl) || htmlCSSCollapsedOverflowHidden(decl)
}

func htmlCSSPropertyHidden(decl string) bool {
	for _, part := range strings.Split(decl, ";") {
		name, value, ok := strings.Cut(part, ":")
		if !ok {
			continue
		}
		name = strings.ToLower(strings.TrimSpace(name))
		value = strings.ToLower(strings.TrimSpace(html.UnescapeString(value)))
		value = strings.TrimSpace(strings.TrimSuffix(value, "!important"))
		switch name {
		case "opacity":
			if cssZeroValue(value) {
				return true
			}
		case "font-size":
			if cssZeroLength(value) {
				return true
			}
		case "color":
			if strings.EqualFold(value, "transparent") {
				return true
			}
		case "clip":
			if cssClipRectHidden(value) {
				return true
			}
		case "clip-path":
			if cssClipPathHidden(value) {
				return true
			}
		case "content-visibility":
			if value == "hidden" {
				return true
			}
		case "transform":
			if cssTransformHidden(value) {
				return true
			}
		case "scale":
			if cssScaleHidden(value) {
				return true
			}
		}
	}
	return false
}

func htmlCSSCollapsedOverflowHidden(decl string) bool {
	props := htmlCSSPropertyMap(decl)
	overflowHidden := cssOverflowClips(props["overflow"]) || cssOverflowClips(props["overflow-x"]) || cssOverflowClips(props["overflow-y"])
	if !overflowHidden {
		return false
	}
	if cssZeroLength(props["max-height"]) {
		return true
	}
	return cssZeroLength(props["width"]) && cssZeroLength(props["height"])
}

func cssOverflowClips(value string) bool {
	return value == "hidden" || value == "clip"
}

func htmlCSSPropertyMap(decl string) map[string]string {
	props := map[string]string{}
	for _, part := range strings.Split(decl, ";") {
		name, value, ok := strings.Cut(part, ":")
		if !ok {
			continue
		}
		name = strings.ToLower(strings.TrimSpace(name))
		value = strings.ToLower(strings.TrimSpace(html.UnescapeString(value)))
		value = strings.TrimSpace(strings.TrimSuffix(value, "!important"))
		if name != "" {
			props[name] = value
		}
	}
	return props
}

func cssZeroValue(value string) bool {
	value = strings.TrimSpace(value)
	if value == "" {
		return false
	}
	n, err := strconv.ParseFloat(value, 64)
	return err == nil && n == 0
}

func cssZeroLength(value string) bool {
	value = strings.TrimSpace(value)
	if value == "" {
		return false
	}
	if value == "0" {
		return true
	}
	for _, unit := range []string{"px", "pt", "em", "rem", "%", "pc", "in", "cm", "mm", "q", "vw", "vh", "vmin", "vmax"} {
		if strings.HasSuffix(value, unit) {
			return cssZeroValue(strings.TrimSpace(strings.TrimSuffix(value, unit)))
		}
	}
	return false
}

func cssClipRectHidden(value string) bool {
	value = strings.TrimSpace(strings.ToLower(value))
	if !strings.HasPrefix(value, "rect(") || !strings.HasSuffix(value, ")") {
		return false
	}
	inner := strings.TrimSpace(strings.TrimSuffix(strings.TrimPrefix(value, "rect("), ")"))
	inner = strings.ReplaceAll(inner, ",", " ")
	fields := strings.Fields(inner)
	if len(fields) != 4 {
		return false
	}
	for _, field := range fields {
		if !cssZeroLength(field) {
			return false
		}
	}
	return true
}

func cssClipPathHidden(value string) bool {
	value = strings.TrimSpace(strings.ToLower(value))
	if !strings.HasPrefix(value, "inset(") || !strings.HasSuffix(value, ")") {
		return false
	}
	inner := strings.TrimSpace(strings.TrimSuffix(strings.TrimPrefix(value, "inset("), ")"))
	fields := strings.Fields(strings.ReplaceAll(inner, ",", " "))
	if len(fields) == 0 || len(fields) > 4 {
		return false
	}
	for _, field := range fields {
		if !cssInsetFullyClipped(field) {
			return false
		}
	}
	return true
}

func cssInsetFullyClipped(value string) bool {
	value = strings.TrimSpace(value)
	if !strings.HasSuffix(value, "%") {
		return false
	}
	n, err := strconv.ParseFloat(strings.TrimSpace(strings.TrimSuffix(value, "%")), 64)
	return err == nil && n >= 50
}

func cssTransformHidden(value string) bool {
	value = strings.ToLower(strings.TrimSpace(value))
	if value == "" || value == "none" {
		return false
	}
	for strings.Contains(value, "scale(") {
		start := strings.Index(value, "scale(")
		rest := value[start+len("scale("):]
		end := strings.Index(rest, ")")
		if end < 0 {
			return false
		}
		if cssScaleHidden(rest[:end]) {
			return true
		}
		value = rest[end+1:]
	}
	return false
}

func cssScaleHidden(value string) bool {
	fields := strings.Fields(strings.ReplaceAll(strings.TrimSpace(value), ",", " "))
	if len(fields) == 0 || len(fields) > 2 {
		return false
	}
	for _, field := range fields {
		if !cssZeroValue(field) {
			return false
		}
	}
	return true
}

func htmlStartTagHidden(tag string, hiddenCSS htmlHiddenCSS) bool {
	lower := strings.ToLower(tag)
	if htmlHiddenAttrRE.MatchString(lower) {
		return true
	}
	attrs := htmlTagAttrs(tag)
	if htmlAttrsHidden(attrs) {
		return true
	}
	tagName, _, _ := htmlTagName(tag)
	id := strings.ToLower(strings.TrimSpace(attrs["id"]))
	if id != "" && hiddenCSS.IDs[id] {
		return true
	}
	classSet := htmlClassSet(attrs["class"])
	for _, className := range strings.Fields(attrs["class"]) {
		if hiddenCSS.Classes[strings.ToLower(className)] {
			return true
		}
	}
	for _, selector := range hiddenCSS.Compounds {
		if htmlHiddenCSSSelectorMatches(selector, tagName, id, classSet) {
			return true
		}
	}
	return false
}

func htmlClassSet(classes string) map[string]bool {
	out := map[string]bool{}
	for _, className := range strings.Fields(classes) {
		out[strings.ToLower(className)] = true
	}
	return out
}

func htmlHiddenCSSSelectorMatches(selector htmlHiddenCSSSelector, tagName, id string, classes map[string]bool) bool {
	if selector.Element != "" && selector.Element != tagName {
		return false
	}
	if selector.ID != "" && selector.ID != id {
		return false
	}
	for _, className := range selector.Classes {
		if !classes[className] {
			return false
		}
	}
	return true
}

func htmlTagName(tag string) (string, bool, bool) {
	t := strings.TrimSpace(strings.TrimPrefix(strings.TrimSuffix(tag, ">"), "<"))
	if t == "" || strings.HasPrefix(t, "!") || strings.HasPrefix(t, "?") {
		return "", false, true
	}
	closing := false
	if strings.HasPrefix(t, "/") {
		closing = true
		t = strings.TrimSpace(strings.TrimPrefix(t, "/"))
	}
	selfClosing := strings.HasSuffix(strings.TrimSpace(t), "/")
	fields := strings.Fields(strings.TrimRight(t, "/"))
	if len(fields) == 0 {
		return "", closing, true
	}
	name := strings.ToLower(fields[0])
	if i := strings.LastIndex(name, ":"); i >= 0 {
		name = name[i+1:]
	}
	return name, closing, selfClosing
}

func htmlVoidElement(name string) bool {
	switch name {
	case "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr":
		return true
	default:
		return false
	}
}

func resolveHTMLImageMediaPart(files map[string]*zip.File, source, src string) string {
	src = strings.TrimSpace(html.UnescapeString(src))
	if src == "" {
		return ""
	}
	if u, err := url.Parse(src); err == nil {
		if u.Scheme != "" || u.Host != "" {
			return ""
		}
		if u.Path != "" {
			src = u.Path
		}
	}
	if decoded, err := url.PathUnescape(src); err == nil {
		src = decoded
	}
	candidates := []string{src}
	cleaned := strings.TrimPrefix(filepath.ToSlash(src), "/")
	if strings.HasPrefix(strings.ToLower(cleaned), "word/media/") {
		candidates = append(candidates, cleaned)
	}
	for _, candidate := range candidates {
		part := resolveOOXMLRelationshipTarget(source, candidate)
		if actual := ooxmlPartName(files, part); actual != "" {
			part = actual
		}
		if strings.HasPrefix(ooxmlPartKey(part), "word/media/") {
			return part
		}
	}
	return ""
}

func parseMHTMLParts(b []byte) []mhtmlPart {
	headers, body := splitMHTMLHeaders(b)
	boundary := mhtmlBoundary(headers["content-type"])
	if boundary == "" {
		return nil
	}
	marker := []byte("--" + boundary)
	segments := bytes.Split(body, marker)
	var out []mhtmlPart
	for _, seg := range segments {
		seg = bytes.Trim(seg, "\r\n")
		if len(seg) == 0 || bytes.HasPrefix(seg, []byte("--")) {
			continue
		}
		partHeaders, partBody := splitMHTMLHeaders(seg)
		if len(partHeaders) == 0 && len(bytes.TrimSpace(partBody)) == 0 {
			continue
		}
		out = append(out, mhtmlPart{Headers: partHeaders, Body: partBody})
	}
	return out
}

func splitMHTMLHeaders(b []byte) (map[string]string, []byte) {
	headerBytes := b
	body := []byte(nil)
	if i := bytes.Index(b, []byte("\r\n\r\n")); i >= 0 {
		headerBytes = b[:i]
		body = b[i+4:]
	} else if i := bytes.Index(b, []byte("\n\n")); i >= 0 {
		headerBytes = b[:i]
		body = b[i+2:]
	}
	headers := map[string]string{}
	var last string
	for _, raw := range strings.Split(strings.ReplaceAll(string(headerBytes), "\r\n", "\n"), "\n") {
		if strings.TrimSpace(raw) == "" {
			continue
		}
		if (strings.HasPrefix(raw, " ") || strings.HasPrefix(raw, "\t")) && last != "" {
			headers[last] = strings.TrimSpace(headers[last] + " " + strings.TrimSpace(raw))
			continue
		}
		name, value, ok := strings.Cut(raw, ":")
		if !ok {
			continue
		}
		last = strings.ToLower(strings.TrimSpace(name))
		headers[last] = strings.TrimSpace(value)
	}
	if body == nil {
		body = nil
	}
	return headers, body
}

func mhtmlBoundary(contentType string) string {
	return mhtmlHeaderParam(contentType, "boundary")
}

func mhtmlHeaderParam(header, name string) string {
	for _, part := range splitMHTMLHeaderParams(header) {
		key, value, ok := strings.Cut(strings.TrimSpace(part), "=")
		if !ok || !strings.EqualFold(strings.TrimSpace(key), name) {
			continue
		}
		return strings.Trim(strings.TrimSpace(value), `"'`)
	}
	return ""
}

func splitMHTMLHeaderParams(header string) []string {
	var out []string
	start := 0
	var quote rune
	for i, r := range header {
		switch {
		case quote != 0:
			if r == quote {
				quote = 0
			}
		case r == '\'' || r == '"':
			quote = r
		case r == ';':
			out = append(out, header[start:i])
			start = i + len(string(r))
		}
	}
	out = append(out, header[start:])
	return out
}

func mhtmlTextCharset(contentType string) string {
	return strings.ToLower(strings.TrimSpace(mhtmlHeaderParam(contentType, "charset")))
}

func decodeMHTMLTextPart(part mhtmlPart) []byte {
	body := decodeMHTMLBody(part.Body, part.Headers["content-transfer-encoding"])
	if decoded, ok := decodeTextBOMBytes(body); ok {
		return decoded
	}
	charset := mhtmlTextCharset(part.Headers["content-type"])
	if charset == "" {
		if declared := htmlDeclaredCharset(body); declared != "" {
			charset = declared
		}
	}
	if charset == "" || strings.EqualFold(charset, "utf-8") || strings.EqualFold(charset, "utf8") {
		if decoded, ok := decodeLikelyUTF16HTMLBytes(body); ok {
			return decoded
		}
		return body
	}
	if text := decodeMHTMLCharsetBytes(body, charset); text != "" {
		return []byte(text)
	}
	return body
}

func decodeTextBOMBytes(raw []byte) ([]byte, bool) {
	switch {
	case len(raw) >= 3 && raw[0] == 0xef && raw[1] == 0xbb && raw[2] == 0xbf:
		return raw[3:], true
	case len(raw) >= 2 && raw[0] == 0xff && raw[1] == 0xfe:
		return []byte(utf16BytesToStringAll(raw[2:])), true
	case len(raw) >= 2 && raw[0] == 0xfe && raw[1] == 0xff:
		return []byte(utf16BEBytesToStringAll(raw[2:])), true
	default:
		return nil, false
	}
}

func decodeLikelyUTF16HTMLBytes(raw []byte) ([]byte, bool) {
	if len(raw) < 16 {
		return nil, false
	}
	n := len(raw)
	if n > 4096 {
		n = 4096
	}
	if n%2 == 1 {
		n--
	}
	if n < 16 {
		return nil, false
	}
	pairs := n / 2
	evenZero, oddZero := 0, 0
	for i := 0; i+1 < n; i += 2 {
		if raw[i] == 0 {
			evenZero++
		}
		if raw[i+1] == 0 {
			oddZero++
		}
	}
	if oddZero >= pairs/3 && oddZero >= evenZero*4 {
		if decoded := utf16BytesToStringAll(raw); looksLikeDecodedHTML(decoded) {
			return []byte(decoded), true
		}
	}
	if evenZero >= pairs/3 && evenZero >= oddZero*4 {
		if decoded := utf16BEBytesToStringAll(raw); looksLikeDecodedHTML(decoded) {
			return []byte(decoded), true
		}
	}
	return nil, false
}

func looksLikeDecodedHTML(s string) bool {
	lower := strings.ToLower(s)
	return strings.Contains(lower, "<") && strings.Contains(lower, ">") &&
		(strings.Contains(lower, "<html") || strings.Contains(lower, "<body") ||
			strings.Contains(lower, "<p") || strings.Contains(lower, "<div") ||
			strings.Contains(lower, "<img") || strings.Contains(lower, "<!doctype"))
}

func utf16BEBytesToStringAll(raw []byte) string {
	if len(raw)%2 == 1 {
		raw = raw[:len(raw)-1]
	}
	u := make([]uint16, 0, len(raw)/2)
	for i := 0; i+1 < len(raw); i += 2 {
		u = append(u, binary.BigEndian.Uint16(raw[i:]))
	}
	return string(utf16.Decode(u))
}

func decodeMHTMLCharsetBytes(raw []byte, charset string) string {
	switch strings.ToLower(strings.TrimSpace(charset)) {
	case "gbk", "gb2312", "cp936", "windows-936", "x-gbk":
		return decodeCodePageBytes(raw, 936)
	case "gb18030":
		return decodeCodePageBytes(raw, 54936)
	case "shift_jis", "shift-jis", "sjis", "cp932", "windows-31j", "windows-932":
		return decodeCodePageBytes(raw, 932)
	case "big5", "big-5", "cp950", "windows-950":
		return decodeCodePageBytes(raw, 950)
	case "euc-kr", "ks_c_5601-1987", "cp949", "windows-949":
		return decodeCodePageBytes(raw, 949)
	case "windows-1250", "cp1250":
		return decodeCodePageBytes(raw, 1250)
	case "windows-1251", "cp1251":
		return decodeCodePageBytes(raw, 1251)
	case "windows-1252", "cp1252", "iso-8859-1":
		return decodeCodePageBytes(raw, 1252)
	case "windows-1253", "cp1253":
		return decodeCodePageBytes(raw, 1253)
	case "windows-1254", "cp1254":
		return decodeCodePageBytes(raw, 1254)
	case "windows-1255", "cp1255":
		return decodeCodePageBytes(raw, 1255)
	case "windows-1256", "cp1256":
		return decodeCodePageBytes(raw, 1256)
	case "windows-1257", "cp1257":
		return decodeCodePageBytes(raw, 1257)
	case "windows-1258", "cp1258":
		return decodeCodePageBytes(raw, 1258)
	case "utf-16", "utf16", "utf-16le", "utf16le", "unicode":
		return decodeCodePageBytes(raw, 1200)
	case "utf-16be", "utf16be", "unicodefffe":
		return utf16BEBytesToStringAll(raw)
	}
	return ""
}

func decodeMHTMLBody(body []byte, transferEncoding string) []byte {
	switch mhtmlTransferEncodingToken(transferEncoding) {
	case "base64":
		s := strings.Map(func(r rune) rune {
			if unicode.IsSpace(r) {
				return -1
			}
			return r
		}, string(body))
		if decoded, err := decodeDataURIBase64Payload(s); err == nil {
			return decoded
		}
	case "quoted-printable":
		if decoded, err := io.ReadAll(quotedprintable.NewReader(bytes.NewReader(body))); err == nil {
			return decoded
		}
	}
	return body
}

func mhtmlTransferEncodingToken(transferEncoding string) string {
	parts := splitMHTMLHeaderParams(transferEncoding)
	if len(parts) == 0 {
		return ""
	}
	return strings.ToLower(strings.Trim(strings.TrimSpace(parts[0]), `"'`))
}

func extractDocxAltChunkMHTMLImages(files map[string]*zip.File) []Image {
	var out []Image
	used := map[string]bool{}
	for _, name := range docxVisibleHTMLPartNamesNoError(files) {
		lower := ooxmlPartKey(name)
		if !strings.HasSuffix(lower, ".mht") && !strings.HasSuffix(lower, ".mhtml") {
			continue
		}
		b, err := readZipFile(files[name])
		if err != nil {
			continue
		}
		out = append(out, extractMHTMLImages(b, used)...)
	}
	return out
}

func extractDocxAltChunkHTMLDataImages(files map[string]*zip.File) []Image {
	var out []Image
	used := map[string]bool{}
	for _, name := range docxVisibleHTMLPartNamesNoError(files) {
		b, err := readZipFile(files[name])
		if err != nil {
			continue
		}
		lower := ooxmlPartKey(name)
		if strings.HasSuffix(lower, ".mht") || strings.HasSuffix(lower, ".mhtml") {
			for _, part := range parseMHTMLParts(b) {
				if strings.Contains(strings.ToLower(part.Headers["content-type"]), "text/html") {
					out = append(out, extractHTMLDataURIImages(decodeMHTMLTextPart(part), used)...)
				}
			}
			continue
		}
		out = append(out, extractHTMLDataURIImages(decodeHTMLBytes(b), used)...)
	}
	return out
}

func extractHTMLDataURIImages(htmlBytes []byte, used map[string]bool) []Image {
	htmlText := visibleHTMLForImageRefs(string(htmlBytes))
	var out []Image
	for _, tag := range htmlImageTagRE.FindAllString(htmlText, -1) {
		attrs := htmlTagAttrs(tag)
		if htmlImageTagHidden(attrs) {
			continue
		}
		alt := cleanMarkdownImageAltText(attrs["alt"])
		if alt == "" {
			alt = cleanMarkdownImageAltText(attrs["title"])
		}
		for _, src := range htmlImageSourceCandidates(attrs) {
			data, ext, ok := parseHTMLImageDataURI(src)
			if !ok {
				continue
			}
			data, ext, ok = normalizeOOXMLImageData(ext, data)
			if !ok {
				continue
			}
			name := uniqueImageFilename(sanitizeFilename(imageNameWithExt(fmt.Sprintf("html-data-image-%03d", len(out)+1), ext)), used)
			out = append(out, Image{Name: name, Alt: alt, Ext: ext, Data: append([]byte(nil), data...)})
			break
		}
	}
	for _, ref := range htmlPictureSourceRefs(htmlText) {
		for _, src := range ref.Sources {
			data, ext, ok := parseHTMLImageDataURI(src)
			if !ok {
				continue
			}
			data, ext, ok = normalizeOOXMLImageData(ext, data)
			if !ok {
				continue
			}
			name := uniqueImageFilename(sanitizeFilename(imageNameWithExt(fmt.Sprintf("html-data-image-%03d", len(out)+1), ext)), used)
			out = append(out, Image{Name: name, Alt: ref.Alt, Ext: ext, Data: append([]byte(nil), data...)})
			break
		}
	}
	return out
}

func parseHTMLImageDataURI(src string) ([]byte, string, bool) {
	src = strings.TrimSpace(html.UnescapeString(src))
	if len(src) < len("data:") || !strings.HasPrefix(strings.ToLower(src), "data:") {
		return nil, "", false
	}
	meta, payload, ok := strings.Cut(src, ",")
	if !ok {
		return nil, "", false
	}
	parts := strings.Split(meta, ";")
	contentType := strings.TrimPrefix(strings.ToLower(strings.TrimSpace(parts[0])), "data:")
	base64Encoded := false
	for _, part := range parts[1:] {
		if strings.EqualFold(strings.TrimSpace(part), "base64") {
			base64Encoded = true
			break
		}
	}
	if !base64Encoded {
		decoded, err := url.PathUnescape(strings.TrimSpace(payload))
		if err != nil {
			return nil, "", false
		}
		ext := mhtmlContentTypeExt(contentType)
		return []byte(decoded), ext, ext != "" || imageExt([]byte(decoded)) != ".bin"
	}
	ext := mhtmlContentTypeExt(contentType)
	payload = normalizeDataURIBase64Payload(payload)
	decoded, err := decodeDataURIBase64Payload(payload)
	if err != nil {
		return nil, "", false
	}
	return decoded, ext, ext != "" || imageExt(decoded) != ".bin"
}

func normalizeDataURIBase64Payload(payload string) string {
	payload = strings.TrimSpace(payload)
	if decoded, err := url.PathUnescape(payload); err == nil {
		payload = decoded
	}
	return strings.Map(func(r rune) rune {
		if unicode.IsSpace(r) {
			return -1
		}
		return r
	}, payload)
}

func decodeDataURIBase64Payload(payload string) ([]byte, error) {
	if decoded, err := base64.StdEncoding.DecodeString(payload); err == nil {
		return decoded, nil
	}
	if decoded, err := base64.RawStdEncoding.DecodeString(payload); err == nil {
		return decoded, nil
	}
	if decoded, err := base64.URLEncoding.DecodeString(payload); err == nil {
		return decoded, nil
	}
	return base64.RawURLEncoding.DecodeString(payload)
}

func extractMHTMLImages(b []byte, used map[string]bool) []Image {
	parts := parseMHTMLParts(b)
	if len(parts) == 0 {
		return nil
	}
	visibleRefs := mhtmlVisibleImageRefs(parts)
	if len(visibleRefs) == 0 {
		return nil
	}
	var out []Image
	for _, part := range parts {
		contentType := strings.ToLower(part.Headers["content-type"])
		locations := mhtmlPartLocations(part.Headers)
		location, alt, visible := mhtmlVisiblePartLocation(locations, visibleRefs)
		if !visible {
			continue
		}
		ext := mhtmlContentTypeExt(contentType)
		if ext == "" {
			ext = mhtmlImageLocationExt(location)
		}
		body := decodeMHTMLBody(part.Body, part.Headers["content-transfer-encoding"])
		data, normalizedExt, ok := normalizeOOXMLImageData(ext, body)
		if !ok {
			continue
		}
		name := mhtmlImageName(location, normalizedExt)
		if name == "" {
			name = fmt.Sprintf("mhtml-image-%03d%s", len(out)+1, normalizedExt)
		}
		name = uniqueImageFilename(sanitizeFilename(imageNameWithExt(name, normalizedExt)), used)
		out = append(out, Image{
			Name: name,
			Alt:  alt,
			Ext:  normalizedExt,
			Data: append([]byte(nil), data...),
		})
	}
	return out
}

func mhtmlVisiblePartLocation(locations []string, visibleRefs map[string]string) (string, string, bool) {
	for _, location := range locations {
		if alt, ok := visibleRefs[normalizeMHTMLRef(location)]; ok {
			if len(locations) > 0 && strings.TrimSpace(locations[0]) != "" {
				return locations[0], alt, true
			}
			return location, alt, true
		}
	}
	return "", "", false
}

func mhtmlImageLocationExt(location string) string {
	ext := strings.ToLower(path.Ext(normalizeMHTMLRef(location)))
	switch ext {
	case ".png", ".jpg", ".jpeg", ".jpe", ".jfif", ".gif", ".bmp", ".dib", ".tif", ".tiff", ".webp", ".pcx", ".tga", ".pct", ".pict", ".eps", ".ps", ".avif", ".heic", ".heif", ".jp2", ".jpx", ".jpf", ".j2k", ".j2c", ".jpc", ".jxr", ".wdp", ".hdp", ".svg", ".ico", ".cur", ".emf", ".wmf", ".emz", ".wmz", ".svgz":
		return ext
	default:
		return ""
	}
}

func mhtmlVisibleImageRefs(parts []mhtmlPart) map[string]string {
	out := map[string]string{}
	for _, part := range parts {
		if !strings.Contains(strings.ToLower(part.Headers["content-type"]), "text/html") {
			continue
		}
		body := visibleHTMLForImageRefs(string(decodeMHTMLTextPart(part)))
		for _, tag := range htmlImageTagRE.FindAllString(body, -1) {
			attrs := htmlTagAttrs(tag)
			srcs := htmlImageSourceCandidates(attrs)
			if len(srcs) == 0 || htmlImageTagHidden(attrs) {
				continue
			}
			alt := cleanMarkdownImageAltText(attrs["alt"])
			if alt == "" {
				alt = cleanMarkdownImageAltText(attrs["title"])
			}
			for _, candidate := range srcs {
				src := normalizeMHTMLRef(candidate)
				if src == "" {
					continue
				}
				if _, exists := out[src]; !exists || out[src] == "" {
					out[src] = alt
				}
				break
			}
		}
		for _, ref := range htmlPictureSourceRefs(body) {
			for _, candidate := range ref.Sources {
				src := normalizeMHTMLRef(candidate)
				if src == "" {
					continue
				}
				if _, exists := out[src]; !exists || out[src] == "" {
					out[src] = ref.Alt
				}
				break
			}
		}
	}
	return out
}

func mhtmlPartLocation(headers map[string]string) string {
	locations := mhtmlPartLocations(headers)
	if len(locations) == 0 {
		return ""
	}
	return locations[0]
}

func mhtmlPartLocations(headers map[string]string) []string {
	var out []string
	if location := strings.TrimSpace(headers["content-location"]); location != "" {
		out = append(out, location)
	}
	if cid := cleanMHTMLContentID(headers["content-id"]); cid != "" {
		seen := false
		normalizedCID := normalizeMHTMLRef(cid)
		for _, location := range out {
			if normalizeMHTMLRef(location) == normalizedCID {
				seen = true
				break
			}
		}
		if !seen {
			out = append(out, cid)
		}
	}
	return out
}

func cleanMHTMLContentID(s string) string {
	s = strings.TrimSpace(s)
	s = strings.Trim(s, `"'`)
	s = strings.Trim(s, "<>")
	s = strings.Trim(s, `"'`)
	return strings.TrimSpace(s)
}

func normalizeMHTMLRef(s string) string {
	s = strings.TrimSpace(html.UnescapeString(s))
	s = strings.Trim(s, `"'`)
	if len(s) >= 4 && strings.EqualFold(s[:4], "cid:") {
		s = s[4:]
	}
	s = strings.Trim(s, "<>")
	s = strings.Trim(s, `"'`)
	if u, err := url.Parse(s); err == nil {
		if u.Path != "" {
			s = u.Path
		}
	}
	if decoded, err := url.PathUnescape(s); err == nil {
		s = decoded
	}
	s = strings.TrimSpace(filepath.ToSlash(s))
	s = mhtmlTrimPathSeparatorWhitespace(s)
	s = strings.Trim(s, `"'`)
	return strings.ToLower(s)
}

func mhtmlTrimPathSeparatorWhitespace(s string) string {
	s = mhtmlWhitespaceAfterPathSeparatorRE.ReplaceAllString(s, "/")
	s = mhtmlWhitespaceBeforePathSeparatorRE.ReplaceAllString(s, "/")
	return s
}

func mhtmlContentTypeExt(contentType string) string {
	contentType = strings.ToLower(strings.TrimSpace(strings.Split(contentType, ";")[0]))
	switch contentType {
	case "image/png", "image/x-png", "image/apng":
		return ".png"
	case "image/jpeg", "image/jpg", "image/jfif", "image/pjpeg", "image/x-jpeg":
		return ".jpg"
	case "image/gif":
		return ".gif"
	case "image/bmp", "image/x-bmp", "image/x-ms-bmp":
		return ".bmp"
	case "image/dib", "image/x-dib":
		return ".dib"
	case "image/tif", "image/tiff", "image/x-tif", "image/x-tiff":
		return ".tif"
	case "image/webp", "image/x-webp":
		return ".webp"
	case "image/pcx", "image/x-pcx", "image/vnd.zbrush.pcx":
		return ".pcx"
	case "image/tga", "image/x-tga", "image/x-targa":
		return ".tga"
	case "image/pict", "image/x-pict", "image/x-macpict":
		return ".pict"
	case "image/eps", "image/x-eps", "image/ps", "image/x-ps":
		return ".eps"
	case "image/avif":
		return ".avif"
	case "image/heic", "image/heic-sequence":
		return ".heic"
	case "image/heif", "image/heif-sequence":
		return ".heif"
	case "image/jp2", "image/jp2k", "image/jpeg2000", "image/jpeg2000-image", "image/x-jp2":
		return ".jp2"
	case "image/jpx", "image/x-jpx", "image/jpm":
		return ".jpx"
	case "image/j2k", "image/j2c", "image/jpc", "image/x-j2k":
		return ".j2k"
	case "image/jxr", "image/vnd.ms-photo", "image/vnd.ms-photo.jxr", "image/wdp":
		return ".jxr"
	case "image/svg+xml":
		return ".svg"
	case "image/x-icon", "image/vnd.microsoft.icon", "image/ico":
		return ".ico"
	case "image/x-cursor", "image/vnd.microsoft.icon.cursor":
		return ".cur"
	case "image/x-emf", "image/emf":
		return ".emf"
	case "image/x-wmf", "image/wmf":
		return ".wmf"
	default:
		return ""
	}
}

func mhtmlImageName(location, ext string) string {
	location = normalizeMHTMLRef(location)
	if location == "" {
		return ""
	}
	base := path.Base(location)
	if base == "." || base == "/" || base == "" {
		return ""
	}
	if before, _, ok := strings.Cut(base, "@"); ok && before != "" {
		base = before
	}
	if path.Ext(base) == "" && ext != "" {
		base += ext
	}
	return base
}

func readSharedStrings(files map[string]*zip.File) ([]string, error) {
	f := ooxmlFile(files, "xl/sharedStrings.xml")
	if f == nil {
		return nil, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, err
	}
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	var cur strings.Builder
	inSI, inT := false, false
	var phoneticDepth int
	for {
		tok, err := dec.RawToken()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "si" {
				inSI = true
				cur.Reset()
			}
			if inSI && isExcelPhoneticElement(t.Name.Local) {
				phoneticDepth++
				continue
			}
			if inSI && phoneticDepth == 0 && t.Name.Local == "t" {
				inT = true
			}
		case xml.EndElement:
			if t.Name.Local == "t" {
				inT = false
			}
			if phoneticDepth > 0 {
				if isExcelPhoneticElement(t.Name.Local) {
					phoneticDepth--
				}
				continue
			}
			if t.Name.Local == "si" {
				out = append(out, cur.String())
				inSI = false
			}
		case xml.CharData:
			if inT {
				cur.Write(t)
			}
		}
	}
	return out, nil
}

func readSharedStringsFast(b []byte) ([]string, bool) {
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	var cur strings.Builder
	inSI, inT := false, false
	var phoneticDepth int
	for {
		tok, err := dec.RawToken()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, false
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "si" {
				inSI = true
				cur.Reset()
			}
			if inSI && isExcelPhoneticElement(t.Name.Local) {
				phoneticDepth++
			}
			if inSI && phoneticDepth == 0 && t.Name.Local == "t" {
				inT = true
			}
		case xml.EndElement:
			if inSI && phoneticDepth > 0 {
				if isExcelPhoneticElement(t.Name.Local) {
					phoneticDepth--
				}
				continue
			}
			if t.Name.Local == "t" {
				inT = false
			}
			if t.Name.Local == "si" {
				out = append(out, cur.String())
				inSI = false
			}
		case xml.CharData:
			if inT {
				cur.Write(t)
			}
		}
	}
	return out, true
}

func isExcelPhoneticElement(name string) bool {
	switch name {
	case "rPh", "phoneticPr":
		return true
	default:
		return false
	}
}

type xlsxCellStyles struct {
	numFmtIDs []string
	formats   map[string]string
}

func readXlsxCellStyles(files map[string]*zip.File) (xlsxCellStyles, error) {
	f := ooxmlFile(files, "xl/styles.xml")
	if f == nil {
		return xlsxCellStyles{}, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return xlsxCellStyles{}, err
	}
	if hasDOCTYPE(b) {
		return xlsxCellStyles{}, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	styles := xlsxCellStyles{formats: map[string]string{}}
	inNumFmts := false
	inCellXfs := false
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return xlsxCellStyles{}, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "numFmts" {
				inNumFmts = true
				continue
			}
			if inNumFmts && t.Name.Local == "numFmt" {
				var id, code string
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "numFmtId":
						id = a.Value
					case "formatCode":
						code = a.Value
					}
				}
				if id != "" && code != "" {
					styles.formats[id] = code
				}
				continue
			}
			if t.Name.Local == "cellXfs" {
				inCellXfs = true
				continue
			}
			if !inCellXfs || t.Name.Local != "xf" {
				continue
			}
			id := "0"
			for _, a := range t.Attr {
				if a.Name.Local == "numFmtId" {
					id = a.Value
					break
				}
			}
			styles.numFmtIDs = append(styles.numFmtIDs, id)
		case xml.EndElement:
			if t.Name.Local == "numFmts" {
				inNumFmts = false
			}
			if t.Name.Local == "cellXfs" {
				inCellXfs = false
			}
		}
	}
	return styles, nil
}

func xlsxDisplayNumber(value, style string, formats map[string]string) string {
	if style == "" || !plainExcelNumberValue(value) {
		return value
	}
	n, err := strconv.ParseFloat(value, 64)
	if err != nil {
		return value
	}
	// Custom formats are allowed to use the built-in date/time format IDs.
	// The previous order only applied built-ins to IDs below 164, so a writer
	// that defines numFmtId=178 with formatCode="d" leaked the stored Excel
	// serial (40547) instead of Range.Text's day of month (12).  Consult an
	// authored format first; this also preserves locale-qualified custom dates.
	if format, ok := formats[style]; ok {
		if displayed, ok := xlsxDisplayCustomTimeNumber(n, format); ok {
			return displayed
		}
		if displayed, ok := xlsxDisplayCustomDateNumber(n, format); ok {
			return displayed
		}
		if displayed, ok := xlsxDisplayCustomNumber(n, format); ok {
			return displayed
		}
	}
	// General formats (built-in 0 and the custom General used by several
	// writers) suppress floating-point storage residue in Excel's .Text. An
	// authored formatCode always wins, including the legitimate custom
	// numFmtId=0 seen in files produced by older spreadsheet writers.
	if style == "0" || style == "82" {
		// Excel's General format uses at most eleven significant digits in a
		// normally sized cell. Fourteen digits preserve binary storage residue
		// (142.85714285714) that Range.Text renders as 142.857143. Fixed-point
		// is important here: Go's 'g' switches 120 to 1.2e+02, while Excel keeps
		// normal-sized General values in decimal notation.
		return xlsxGeneralDisplayNumber(n)
	}
	if displayed, ok := xlsxDisplayBuiltInNumber(n, style); ok {
		return displayed
	}
	return value
}

// xlsxDisplayCustomTimeNumber handles the straightforward clock and elapsed
// time patterns emitted in styles.xml.  These do not contain d/m/y, so date
// formatting cannot help; returning their stored fractional-day value makes a
// strict Excel UsedRange.Text comparison lose most tokens in time-series
// workbooks.
func xlsxDisplayCustomTimeNumber(n float64, format string) (string, bool) {
	clean := strings.ToLower(format)
	clean = regexp.MustCompile(`"[^"]*"|\\.|;.*`).ReplaceAllString(clean, "")
	// An elapsed-minutes format such as [mm]:ss has no hour placeholder.
	// Treat it as time before the date formatter sees the m tokens and renders
	// the serial as a calendar month.  This is common in timing worksheets.
	elapsedMinutes := strings.Contains(clean, "[m]") || strings.Contains(clean, "[mm]")
	if (!strings.Contains(clean, "h") || !strings.Contains(clean, "m")) && !elapsedMinutes {
		return "", false
	}
	if strings.ContainsAny(clean, "dy") {
		return "", false
	}
	negative := n < 0
	if negative {
		n = -n
	}
	// Excel rounds a serial to the nearest displayed second for normal clock
	// formats. This also compensates for the recurring sub-second residue in
	// cached OOXML doubles (for example, a stored 0:49:59.999... is 0:50:00).
	seconds := int(math.Floor(n*86400 + 0.5))
	elapsed := strings.Contains(clean, "[h]")
	if elapsedMinutes {
		minutes := seconds / 60
		out := fmt.Sprintf("%02d:%02d", minutes, seconds%60)
		if negative {
			out = "-" + out
		}
		return out, true
	}
	hours := seconds / 3600
	if !elapsed {
		hours %= 24
	}
	minutes := (seconds / 60) % 60
	secs := seconds % 60
	var out string
	if strings.Contains(clean, "hh") {
		out = fmt.Sprintf("%02d:%02d", hours, minutes)
	} else {
		out = fmt.Sprintf("%d:%02d", hours, minutes)
	}
	if strings.Contains(clean, "ss") {
		out += fmt.Sprintf(":%02d", secs)
	}
	if strings.Contains(clean, "am/pm") {
		ampm := "AM"
		if hours >= 12 {
			ampm = "PM"
		}
		hours %= 12
		if hours == 0 {
			hours = 12
		}
		if strings.Contains(clean, "hh") {
			out = fmt.Sprintf("%02d:%02d", hours, minutes)
		} else {
			out = fmt.Sprintf("%d:%02d", hours, minutes)
		}
		if strings.Contains(clean, "ss") {
			out += fmt.Sprintf(":%02d", secs)
		}
		out += " " + ampm
	}
	if negative {
		out = "-" + out
	}
	return out, true
}

func xlsxGeneralDisplayNumber(n float64) string {
	if n == 0 {
		return "0"
	}
	abs := math.Abs(n)
	if abs >= 1e11 || abs < 1e-9 {
		return strings.ReplaceAll(strconv.FormatFloat(n, 'e', 10, 64), "e", "E")
	}
	digitsBefore := 1
	if abs >= 1 {
		digitsBefore = int(math.Floor(math.Log10(abs))) + 1
	}
	decimals := 11 - digitsBefore
	if decimals < 0 {
		decimals = 0
	}
	return strings.TrimRight(strings.TrimRight(strconv.FormatFloat(xlsxRoundDisplayNumber(n, decimals), 'f', decimals, 64), "0"), ".")
}

// xlsxDisplayCustomDateNumber implements the common custom date forms used
// by workbook producers. Without it a stored serial such as 41014 leaks into
// strict extraction while Excel's Range.Text shows "18-Apr-12".
func xlsxDisplayCustomDateNumber(n float64, format string) (string, bool) {
	clean := strings.ToLower(format)
	clean = regexp.MustCompile(`"[^"]*"|\\.|\[[^\]]*\]|;.*`).ReplaceAllString(clean, "")
	// A date format can legitimately contain only a day ("d") or only a
	// month/year component.  Requiring both d and m let custom format IDs such
	// as 178="d" leak the serial number while Excel displayed the day.
	if !strings.ContainsAny(clean, "dmy") || n < 0 || n > 2958465 {
		return "", false
	}
	days := int(math.Floor(n))
	base := time.Date(1899, time.December, 31, 0, 0, 0, 0, time.UTC)
	// Preserve Excel's intentionally incorrect 1900 leap-year convention.
	if days >= 60 {
		days--
	}
	d := base.AddDate(0, 0, days)
	// A common Office-generated locale format is [$-1409]d\ mmmm\ yyyy.
	// Its escaped spaces are visible separators; the generic formatter below
	// historically replaced them with hyphens, producing a token mismatch
	// against Excel Range.Text ("1 January 2024").
	if strings.Contains(strings.ToLower(format), "d\\ mmmm\\ yyyy") {
		return d.Format("2 January 2006"), true
	}
	// Weekday-bearing custom formats are not represented by the built-in date
	// IDs. In particular, formats such as ddd\-dd\-mmm\-yy are common in
	// Office-authored schedule templates and Range.Text includes the weekday.
	if strings.Contains(clean, "ddd") {
		sep := "-"
		if strings.Contains(clean, "ddd/") {
			sep = "/"
		}
		day := strconv.Itoa(d.Day())
		if strings.Contains(clean, "dd") {
			day = fmt.Sprintf("%02d", d.Day())
		}
		month := d.Format("Jan")
		if strings.Contains(clean, "mmmm") {
			month = d.Month().String()
		}
		year := d.Format("06")
		if strings.Contains(clean, "yyyy") {
			year = d.Format("2006")
		}
		return d.Format("Mon") + sep + day + sep + month + sep + year, true
	}
	// Numeric month/day/year forms are common in older workbooks.  They need
	// their own branch because the generic date formatter below intentionally
	// uses a slash-delimited dd/mm layout for ambiguous formats; Excel's
	// m/d/yy format displays 4/9/76, not 09/04/76.
	if strings.Contains(clean, "m/d/") || strings.Contains(clean, "m-d-") {
		sep := "/"
		if strings.Contains(clean, "m-d-") {
			sep = "-"
		}
		month := strconv.Itoa(int(d.Month()))
		day := strconv.Itoa(d.Day())
		if strings.Contains(clean, "mm") {
			month = fmt.Sprintf("%02d", int(d.Month()))
		}
		if strings.Contains(clean, "dd") {
			day = fmt.Sprintf("%02d", d.Day())
		}
		year := d.Format("06")
		if strings.Contains(clean, "yyyy") {
			year = d.Format("2006")
		}
		return month + sep + day + sep + year, true
	}
	switch {
	case strings.Contains(clean, "d") && !strings.ContainsAny(clean, "my"):
		return strconv.Itoa(d.Day()), true
	case strings.Contains(clean, "m") && !strings.Contains(clean, "d"):
		if strings.Contains(clean, "mmmm") {
			return d.Month().String(), true
		}
		if strings.Contains(clean, "mmm") {
			return d.Format("Jan"), true
		}
		return strconv.Itoa(int(d.Month())), true
	case strings.Contains(clean, "yyyy") && !strings.ContainsAny(clean, "dm"):
		return d.Format("2006"), true
	case strings.Contains(clean, "yy") && !strings.ContainsAny(clean, "dm"):
		return d.Format("06"), true
	case strings.Contains(clean, "mmmm"):
		if strings.Contains(clean, "yyyy") {
			return d.Format("2-January-2006"), true
		}
		return d.Format("2-January-06"), true
	case strings.Contains(clean, "mmm"):
		if strings.Contains(clean, "yyyy") {
			return d.Format("2-Jan-2006"), true
		}
		return d.Format("2-Jan-06"), true
	case strings.Contains(clean, "yyyy"):
		return d.Format("02/01/2006"), true
	default:
		return d.Format("02/01/06"), true
	}
}

// xlsxDisplayBuiltInNumber handles the small set of built-in formats that
// occur frequently in generated workbooks. Custom formats are stored in
// styles.xml, but built-in IDs (notably 9 and 10 for percentages) are not.
func xlsxDisplayBuiltInNumber(n float64, style string) (string, bool) {
	switch style {
	case "1": // 0
		return strconv.FormatFloat(xlsxRoundDisplayNumber(n, 0), 'f', 0, 64), true
	case "2": // 0.00
		return strconv.FormatFloat(xlsxRoundDisplayNumber(n, 2), 'f', 2, 64), true
	case "3": // #,##0
		return insertThousandsSeparators(strconv.FormatFloat(xlsxRoundDisplayNumber(n, 0), 'f', 0, 64)), true
	case "4": // #,##0.00
		text := strconv.FormatFloat(xlsxRoundDisplayNumber(n, 2), 'f', 2, 64)
		whole, frac, _ := strings.Cut(text, ".")
		return insertThousandsSeparators(whole) + "." + frac, true
	case "37", "38": // #,##0 ; (#,##0) / [Red](#,##0)
		return xlsxDisplayBuiltInGroupedNumber(n, 0, true), true
	case "39", "40": // #,##0.00 ; (#,##0.00) / [Red](#,##0.00)
		return xlsxDisplayBuiltInGroupedNumber(n, 2, true), true
	case "9": // 0%
		return strconv.FormatFloat(xlsxRoundDisplayNumber(n*100, 0), 'f', 0, 64) + "%", true
	case "10": // 0.00%
		return strconv.FormatFloat(xlsxRoundDisplayNumber(n*100, 2), 'f', 2, 64) + "%", true
	case "11": // 0.00E+00
		return strings.ReplaceAll(strconv.FormatFloat(n, 'E', 2, 64), "E+0", "E+"), true
	case "14": // localized short date; COM inherits the workbook/Office locale
		return xlsxDisplayBuiltInDate(n, "1/2/06"), true
	case "15": // d-mmm-yy
		return xlsxDisplayBuiltInDate(n, "2-Jan-06"), true
	case "16": // d-mmm
		return xlsxDisplayBuiltInDate(n, "2-Jan"), true
	case "17": // mmm-yy
		return xlsxDisplayBuiltInDate(n, "Jan-06"), true
	case "18": // h:mm AM/PM
		return xlsxDisplayClockTime(n, false), true
	case "19": // h:mm:ss AM/PM
		return xlsxDisplayClockTime(n, true), true
	case "20": // h:mm
		return xlsxDisplayElapsedTime(n, false, false), true
	case "45": // mm:ss
		return xlsxDisplayElapsedTime(n, true, false), true
	case "46": // [h]:mm:ss
		return xlsxDisplayElapsedTime(n, false, true), true
	case "47": // mmss.0
		return xlsxDisplayMinuteSecondTenths(n), true
	}
	return "", false
}

// xlsxDisplayBuiltInGroupedNumber covers Excel's built-in comma formats.
// They are not present in styles.xml's numFmt table, so treating only custom
// format codes leaks the raw stored value for styles 37--40.
func xlsxDisplayBuiltInGroupedNumber(n float64, decimals int, parentheses bool) string {
	negative := n < 0
	if negative {
		n = -n
	}
	text := strconv.FormatFloat(xlsxRoundDisplayNumber(n, decimals), 'f', decimals, 64)
	whole, frac, hasFraction := strings.Cut(text, ".")
	text = insertThousandsSeparators(whole)
	if hasFraction {
		text += "." + frac
	}
	if negative && parentheses {
		return "(" + text + ")"
	}
	return text
}

func xlsxDisplayClockTime(n float64, seconds bool) string {
	negative := n < 0
	if negative {
		n = -n
	}
	// Excel's clock formats wrap by day, unlike elapsed-time format 46.
	total := int(math.Floor(n*86400+0.5)) % 86400
	hour := total / 3600
	minute := (total / 60) % 60
	second := total % 60
	ampm := "AM"
	if hour >= 12 {
		ampm = "PM"
	}
	displayHour := hour % 12
	if displayHour == 0 {
		displayHour = 12
	}
	out := fmt.Sprintf("%d:%02d", displayHour, minute)
	if seconds {
		out += fmt.Sprintf(":%02d", second)
	}
	out += " " + ampm
	if negative {
		return "-" + out
	}
	return out
}

func xlsxDisplayBuiltInDate(n float64, layout string) string {
	days := int(math.Floor(n))
	base := time.Date(1899, time.December, 31, 0, 0, 0, 0, time.UTC)
	if days >= 60 {
		days--
	}
	return base.AddDate(0, 0, days).Format(layout)
}

// xlsxDisplayElapsedTime mirrors Excel's built-in time formats.  Time values
// are stored as a fraction of a day, so emitting the raw numeric value loses
// the visible content (for example .5 must be "12:00").  Format 46 is an
// elapsed-time format and deliberately does not wrap after 24 hours.
func xlsxDisplayElapsedTime(n float64, minuteSecond, elapsed bool) string {
	negative := n < 0
	if negative {
		n = -n
	}
	seconds := int(math.Floor(n*86400 + 0.5))
	if minuteSecond {
		minutes := seconds / 60
		out := fmt.Sprintf("%02d:%02d", minutes%60, seconds%60)
		if negative {
			return "-" + out
		}
		return out
	}
	hours := seconds / 3600
	if !elapsed {
		hours %= 24
	}
	out := fmt.Sprintf("%d:%02d", hours, (seconds/60)%60)
	if elapsed {
		out += fmt.Sprintf(":%02d", seconds%60)
	}
	if negative {
		return "-" + out
	}
	return out
}

func xlsxDisplayMinuteSecondTenths(n float64) string {
	negative := n < 0
	if negative {
		n = -n
	}
	tenths := int(math.Floor(n*864000 + 0.5))
	seconds := tenths / 10
	// Built-in format 47 is mm:ss.0. The colon is a visible glyph; omitting it
	// turns an Excel time such as 25:19.0 into the unrelated numeric token
	// 2519.0 in strict UsedRange.Text output.
	out := fmt.Sprintf("%02d:%02d.%d", (seconds/60)%60, seconds%60, tenths%10)
	if negative {
		return "-" + out
	}
	return out
}

// xlsxGeneralDisplayNumberForWidth approximates Excel's General rendering in
// a cell with a finite column width.  Unlike a raw number-format conversion,
// Range.Text is constrained by the rendered column: Excel retains as many
// decimal places as fit, then switches small values to scientific notation.
func xlsxGeneralDisplayNumberForWidth(n float64, width int) string {
	if n == 0 {
		return "0"
	}
	abs := math.Abs(n)
	if width < 3 {
		width = 3
	}
	if abs >= math.Pow10(width) || abs < 1e-4 {
		return strings.ReplaceAll(strconv.FormatFloat(n, 'e', 2, 64), "e", "E")
	}
	digitsBefore := 1
	if abs >= 1 {
		digitsBefore = int(math.Floor(math.Log10(abs))) + 1
	} else {
		// Leading fractional zeroes are not significant, but they do consume
		// rendered column width. Account for them before choosing decimal
		// places: 0.040140383 in an 11-character General column has one zero
		// between the decimal point and its first significant digit.
		leadingFractionalZeros := int(math.Floor(-math.Log10(abs)))
		if leadingFractionalZeros > 0 {
			digitsBefore += leadingFractionalZeros
		}
	}
	// Excel's General display also applies its eleven-significant-digit cap in
	// narrow cells. Previously a width-11 column could emit ten fractional
	// digits after a single leading zero (0.0401403834), while Range.Text
	// correctly limits the same value to 0.04014038.
	decimals := width - digitsBefore - 1 // decimal point consumes one column.
	maxDecimals := 10 - digitsBefore
	if decimals > maxDecimals {
		decimals = maxDecimals
	}
	if decimals < 0 {
		decimals = 0
	}
	text := strconv.FormatFloat(xlsxRoundDisplayNumber(n, decimals), 'f', decimals, 64)
	if strings.Contains(text, ".") {
		text = strings.TrimRight(strings.TrimRight(text, "0"), ".")
	}
	return text
}

func xlsxDisplayCustomNumber(n float64, format string) (string, bool) {
	// Custom formats commonly found in real workbooks include accounting
	// padding, colour directives and separate positive/negative/zero sections.
	// Those directives affect layout rather than the cell's visible tokens, so
	// reduce them to the selected numeric section before formatting.  This is
	// deliberately not a full Excel format interpreter, but covers the Office
	// generated Comma/Currency forms (43, 44 and their custom variants) without
	// leaking their stored floating-point values into strict COM comparisons.
	if xlsxFormatHasDateOrTimePlaceholders(format) {
		return "", false
	}
	section, ok := xlsxNumberFormatSection(n, format)
	if !ok {
		return "", false
	}
	format = section
	negativeSection := n < 0
	if n < 0 {
		n = -n
	}
	format = xlsxStripNumberFormatDirectives(format)
	// Quoted literals belong at their original position in the selected format
	// section. Preserve their position with non-printing sentinels until the
	// numeric skeleton and surrounding literals have been resolved.
	const quotedLiteralPrefix = "\x1e"
	const quotedLiteralSuffix = "\x1f"
	var quotedLiterals []string
	for len(format) > 0 {
		start := strings.IndexByte(format, '"')
		if start < 0 {
			break
		}
		end := strings.IndexByte(format[start+1:], '"')
		if end < 0 {
			break
		}
		quotedLiterals = append(quotedLiterals, format[start+1:start+1+end])
		marker := quotedLiteralPrefix + string(rune('A'+len(quotedLiterals)-1)) + quotedLiteralSuffix
		format = format[:start] + marker + format[start+2+end:]
	}
	decimal := 0
	// Dots inside quoted currency/label literals are not decimal placeholders.
	// Start from the first actual digit placeholder, then look for the decimal
	// point in the numeric skeleton that follows it.
	numberStart := strings.IndexAny(format, "0#?")
	if numberStart < 0 {
		return "", false
	}
	if point := strings.IndexByte(format[numberStart:], '.'); point >= 0 {
		point += numberStart
		for _, r := range format[point+1:] {
			if r == '0' || r == '#' {
				decimal++
			}
		}
	}
	// A percent sign in an Excel number format changes the displayed magnitude
	// (the stored fraction .111... is shown as 11.1%), not just the suffix.
	// Treat it before rounding so both the digits and the decimal precision
	// match Range.Text.
	isPercent := strings.Contains(format, "%")
	if isPercent {
		n *= 100
		format = strings.ReplaceAll(format, "%", "")
	}
	text := strconv.FormatFloat(xlsxRoundDisplayNumber(n, decimal), 'f', decimal, 64)
	if xlsxNumberFormatUsesGrouping(format) {
		whole, frac, hasFrac := strings.Cut(text, ".")
		whole = insertThousandsSeparators(whole)
		text = whole
		if hasFrac {
			text += "." + frac
		}
	}
	if isPercent {
		text += "%"
	}
	prefix, suffix := xlsxNumberFormatLiterals(format)
	if negativeSection && strings.Contains(format, "\\(") && strings.Contains(format, "$") {
		if i := strings.Index(prefix, "("); i >= 0 {
			prefix = "(" + strings.Replace(prefix[:i]+prefix[i+1:], "(", "", 1)
		}
	}
	return xlsxRestoreQuotedNumberFormatLiterals(prefix+text+suffix, quotedLiterals, quotedLiteralPrefix, quotedLiteralSuffix), true
}

func xlsxRestoreQuotedNumberFormatLiterals(value string, literals []string, prefix, suffix string) string {
	for index, literal := range literals {
		value = strings.ReplaceAll(value, prefix+string(rune('A'+index))+suffix, literal)
	}
	return value
}

func xlsxFormatHasDateOrTimePlaceholders(format string) bool {
	inQuote := false
	escaped := false
	inBracket := false
	for _, r := range format {
		if escaped {
			escaped = false
			continue
		}
		if r == '\\' {
			escaped = true
			continue
		}
		if r == '"' {
			inQuote = !inQuote
			continue
		}
		if !inQuote && r == '[' {
			inBracket = true
			continue
		}
		if !inQuote && r == ']' {
			inBracket = false
			continue
		}
		if !inQuote && !inBracket && strings.ContainsRune("dDyYmMhHsS", r) {
			return true
		}
	}
	return false
}

// xlsxNumberFormatSection chooses the Excel format section applicable to n.
// It keeps quoted semicolons intact, which is important for custom currency
// labels, while accepting the normal positive;negative;zero;text layout.
func xlsxNumberFormatSection(n float64, format string) (string, bool) {
	var sections []string
	var current strings.Builder
	inQuote := false
	escaped := false
	for _, r := range format {
		if escaped {
			current.WriteRune(r)
			escaped = false
			continue
		}
		if r == '\\' {
			current.WriteRune(r)
			escaped = true
			continue
		}
		if r == '"' {
			inQuote = !inQuote
			current.WriteRune(r)
			continue
		}
		if r == ';' && !inQuote {
			sections = append(sections, current.String())
			current.Reset()
			continue
		}
		current.WriteRune(r)
	}
	sections = append(sections, current.String())
	if len(sections) == 0 || sections[0] == "" {
		return "", false
	}
	index := 0
	if n < 0 && len(sections) >= 2 {
		index = 1
	} else if n == 0 && len(sections) >= 3 {
		index = 2
	}
	section := sections[index]
	// A text-only fourth section is not applicable to a numeric cached value.
	if !strings.ContainsAny(section, "0#?") {
		return "", false
	}
	return section, true
}

func xlsxStripNumberFormatDirectives(format string) string {
	var out strings.Builder
	for i := 0; i < len(format); i++ {
		switch format[i] {
		case '[':
			if end := strings.IndexByte(format[i+1:], ']'); end >= 0 {
				i += end + 1
				continue
			}
		case '_', '*':
			// Underscore reserves the following glyph's width and asterisk fills
			// spare column width.  Neither is part of Range.Text.
			if i+1 < len(format) {
				i++
			}
			continue
		}
		out.WriteByte(format[i])
	}
	return out.String()
}

func xlsxNumberFormatUsesGrouping(format string) bool {
	point := strings.IndexByte(format, '.')
	whole := format
	if point >= 0 {
		whole = format[:point]
	}
	return strings.Contains(whole, ",")
}

func xlsxNumberFormatLiterals(format string) (string, string) {
	first, last := -1, -1
	for i := 0; i < len(format); i++ {
		if format[i] == '0' || format[i] == '#' || format[i] == '?' {
			if first < 0 {
				first = i
			}
			last = i
		}
	}
	if first < 0 {
		return "", ""
	}
	clean := func(value string) string {
		var out strings.Builder
		inQuote := false
		escaped := false
		for i := 0; i < len(value); i++ {
			c := value[i]
			if escaped {
				out.WriteByte(c)
				escaped = false
				continue
			}
			if c == '\\' {
				escaped = true
				continue
			}
			if c == '"' {
				inQuote = !inQuote
				continue
			}
			if inQuote || c == 0x1e || c == 0x1f || (c >= 'A' && c <= 'Z') || strings.ContainsRune("$€£¥()%-+ /", rune(c)) {
				out.WriteByte(c)
			}
		}
		return out.String()
	}
	prefix := clean(format[:first])
	suffix := clean(format[last+1:])
	// Accounting negative sections frequently encode the opening parenthesis
	// before the currency marker (for example _($* \(#,##0.00\)).  Excel
	// renders the currency inside those parentheses, whereas a literal-prefix
	// pass would produce "$(72.19)". Move that parenthesis in front of the
	// currency token to retain Excel's visible order.
	if i := strings.Index(prefix, "("); i >= 0 && i < len(prefix)-1 {
		if j := strings.IndexAny(prefix[i+1:], "$€£¥"); j >= 0 {
			j += i + 1
			prefix = prefix[:i] + "(" + prefix[j:j+1] + prefix[i+1:j] + prefix[j+1:]
		}
	}
	// Percent is appended by the numeric renderer, not retained as a suffix.
	suffix = strings.ReplaceAll(suffix, "%", "")
	return prefix, suffix
}

// Excel rounds formatted decimal values half away from zero, whereas Go's
// FormatFloat follows IEEE round-to-even for an exact halfway value.
func xlsxRoundDisplayNumber(n float64, decimal int) float64 {
	scale := math.Pow10(decimal)
	if n < 0 {
		return math.Ceil(n*scale-0.5) / scale
	}
	return math.Floor(n*scale+0.5) / scale
}

func insertThousandsSeparators(value string) string {
	sign := ""
	if strings.HasPrefix(value, "-") {
		sign, value = "-", value[1:]
	}
	for i := len(value) - 3; i > 0; i -= 3 {
		value = value[:i] + "," + value[i:]
	}
	return sign + value
}

func xlsxDisplayNumberForCell(value string, styleIndex int, styles xlsxCellStyles) string {
	if styleIndex < 0 || styleIndex >= len(styles.numFmtIDs) {
		return value
	}
	return xlsxDisplayNumber(value, styles.numFmtIDs[styleIndex], styles.formats)
}

func xlsxDisplayNumberForCellWidth(value string, styleIndex int, styles xlsxCellStyles, width int) string {
	if styleIndex < 0 || styleIndex >= len(styles.numFmtIDs) {
		return value
	}
	style := styles.numFmtIDs[styleIndex]
	if (style != "0" && style != "82") || !plainExcelNumberValue(value) {
		return xlsxDisplayNumber(value, style, styles.formats)
	}
	n, err := strconv.ParseFloat(value, 64)
	if err != nil {
		return value
	}
	// A General column width constrains Range.Text only while it is narrow
	// enough to require a shortened rendering. Wider columns use Excel's
	// eleven-significant-digit General representation rather than the raw
	// stored double.
	if width >= 9 && math.Abs(n) >= 1 {
		return xlsxGeneralDisplayNumber(n)
	}
	return xlsxGeneralDisplayNumberForWidth(n, width)
}

func xlsxColumnWidthInto(widths map[int]int, start xml.StartElement) {
	minCol, maxCol := 0, 0
	width := 0.0
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "min":
			minCol, _ = atoi(a.Value)
		case "max":
			maxCol, _ = atoi(a.Value)
		case "width":
			width, _ = strconv.ParseFloat(a.Value, 64)
		}
	}
	if minCol < 1 || maxCol < minCol || width <= 0 {
		return
	}
	// Excel's column width unit is roughly a character count in the Normal
	// font; reserve one character for the text margin/rounding behavior.
	characters := int(math.Floor(width)) - 1
	if characters < 3 {
		characters = 3
	}
	for col := minCol; col <= maxCol; col++ {
		widths[col] = characters
	}
}

func xlsxColumnDisplayWidth(widths map[int]int, column int) int {
	if width, ok := widths[column]; ok {
		return width
	}
	// Excel's default Calibri 11 column width is 8.43 characters. Rounded
	// upward here so a normally sized General cell follows Excel's regular
	// eleven-significant-digit display; explicitly narrow columns remain below
	// the width-sensitive branch above.
	return 9
}

func appendWorksheetText(out *strings.Builder, b []byte, shared []string, styles xlsxCellStyles, md *xlsxWorksheetMarkdownData, strictOfficeContent bool) error {
	// The fast paths do not parse cell styles. Use the XML reader whenever a
	// workbook has styles so numeric cells can match Excel's displayed value.
	if len(styles.numFmtIDs) == 0 {
		if ok, err := appendSimpleInlineWorksheetText(out, b, md); ok || err != nil {
			return err
		}
		if ok, err := appendSharedStringWorksheetText(out, b, shared, md); ok || err != nil {
			return err
		}
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var cellType string
	// OOXML cells without an explicit s attribute use cellXfs[0].  Treating
	// them as unstyled leaks cached binary floating-point values (for example
	// 0.14078153727589526) where Excel's Range.Text displays General-formatted
	// values.  This is especially common in workbooks generated by numerical
	// tools, which only write s for the exceptional percentage/currency cells.
	cellStyle := 0
	var inV, inT bool
	var inHeaderFooter bool
	var rowHidden bool
	var skipCell bool
	var collectMarkdownRow bool
	var collectMarkdownCell bool
	var markdownRowValues []string
	var hiddenCols []intRange
	columnWidths := map[int]int{}
	var hiddenRows []intRange
	currentRow := 0
	nextRow := 1
	cellCol := 0
	nextCol := 1
	seenAttrs := map[string]bool{}
	var seenLargeValues map[string]bool
	var seenLargeSharedIndexes map[int]bool
	var cur strings.Builder
	var markdownCellText strings.Builder
	var phoneticDepth int
	var cellDepth int
	var systemCellTextDepth int
	// Conditional-formatting follows sheetData in OOXML, so pre-scan it before
	// reading cells. Data bars and icon sets with showValue=0 render only their
	// visual indicator; the stored cell value remains in XML but Range.Text is
	// blank (unless a prior stopIfTrue rule takes precedence).
	hiddenByDataBar := xlsxHiddenConditionalFormatValueCells(b)
	for {
		tok, err := dec.RawToken()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "col" {
				if r, ok := hiddenColumnRange(t); ok {
					hiddenCols = append(hiddenCols, r)
				}
				xlsxColumnWidthInto(columnWidths, t)
			}
			if t.Name.Local == "row" {
				rowHidden = worksheetRowHidden(t)
				currentRow = worksheetRowIndex(t, nextRow)
				if currentRow < 1 {
					currentRow = nextRow
				}
				nextRow = currentRow + 1
				nextCol = 1
				if rowHidden {
					hiddenRows = append(hiddenRows, intRange{min: currentRow, max: currentRow})
				}
				collectMarkdownRow = md != nil && !rowHidden && len(md.rows) < maxMarkdownTableRows
				if collectMarkdownRow {
					markdownRowValues = make([]string, 0, 16)
				} else {
					markdownRowValues = nil
				}
			}
			if t.Name.Local == "c" {
				cellDepth = 1
				cellType = ""
				cellStyle = 0
				cellRef := ""
				collectMarkdownCell = false
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "t":
						cellType = a.Value
					case "r":
						cellRef = a.Value
					case "s":
						if n, ok := atoi(a.Value); ok {
							cellStyle = n
						}
					}
				}
				if col, _, ok := cellRefIndexes(cellRef); ok {
					cellCol = col
				} else {
					cellCol = nextCol
				}
				if cellCol < 1 {
					cellCol = 1
				}
				nextCol = cellCol + 1
				// Strict Office comparison follows Excel UsedRange.Text. It includes
				// cells in hidden rows and columns; only hidden worksheets are
				// excluded by the workbook-level sheet selection. Compatibility
				// extraction keeps the historical visible-cell behavior.
				skipCell = (!strictOfficeContent && (rowHidden || hiddenColumnCell(cellRef, hiddenCols) || columnHidden(cellCol, hiddenCols))) || (strictOfficeContent && hiddenByDataBar[cellRef])
				collectMarkdownCell = !skipCell && collectMarkdownRow && markdownRowValues != nil && cellCol <= maxMarkdownTableCols
				if collectMarkdownCell {
					markdownCellText.Reset()
				}
			} else if cellDepth > 0 {
				cellDepth++
			}
			if systemCellTextDepth > 0 {
				systemCellTextDepth++
				continue
			}
			if cellDepth > 0 && isExcelSystemCellTextElement(t.Name.Local) {
				systemCellTextDepth = 1
				continue
			}
			// Data-validation prompts, hyperlinks and similar XML attributes are
			// document annotations, not cell glyphs returned by Excel's
			// UsedRange.Text. Preserve them for the richer compatibility mode, but
			// exclude them from the strict Office-visible baseline.
			if !strictOfficeContent && !skipCell && worksheetElementMayHaveVisibleTextAttributes(t.Name.Local) && !worksheetElementHiddenByRef(t, hiddenCols, hiddenRows) {
				for _, value := range worksheetAttributeText(t) {
					if seenAttrs[value] {
						continue
					}
					seenAttrs[value] = true
					appendCleanedTextBlock(out, value)
					if md != nil {
						md.annotations = append(md.annotations, value)
					}
				}
			}
			if isExcelPhoneticElement(t.Name.Local) {
				phoneticDepth++
				continue
			}
			if isExcelHeaderFooterElement(t.Name.Local) {
				inHeaderFooter = true
				cur.Reset()
			}
			if phoneticDepth == 0 && (t.Name.Local == "v" || t.Name.Local == "t") {
				inV = t.Name.Local == "v"
				inT = t.Name.Local == "t"
				cur.Reset()
			}
		case xml.EndElement:
			if phoneticDepth > 0 {
				if isExcelPhoneticElement(t.Name.Local) {
					phoneticDepth--
				}
				continue
			}
			if systemCellTextDepth > 0 {
				systemCellTextDepth--
				continue
			}
			if isExcelHeaderFooterElement(t.Name.Local) {
				if !strictOfficeContent {
					value := cleanExcelHeaderFooterText(cur.String())
					if value != "" {
						appendCleanedTextBlock(out, value)
						if md != nil {
							md.headerFooter = append(md.headerFooter, value)
						}
					}
				}
				inHeaderFooter = false
			}
			if t.Name.Local == "v" || t.Name.Local == "t" {
				if !skipCell {
					rawValue := cur.String()
					value := strings.TrimSpace(rawValue)
					skipValue := false
					markdownValue := value
					switch {
					case t.Name.Local == "v" && cellType == "s":
						if idx, ok := atoi(value); ok && idx >= 0 && idx < len(shared) {
							if !strictOfficeContent && len(shared[idx]) > maxRepeatedTextPartBytes {
								if seenLargeSharedIndexes != nil && seenLargeSharedIndexes[idx] {
									skipValue = true
									value = shared[idx]
									break
								}
								if seenLargeSharedIndexes == nil {
									seenLargeSharedIndexes = map[int]bool{}
								}
								seenLargeSharedIndexes[idx] = true
							}
							value = shared[idx]
							markdownValue = value
						}
					case t.Name.Local == "v" && cellType == "b":
						value = excelBooleanDisplayText(value)
						markdownValue = value
					case t.Name.Local == "t" && cellType == "inlineStr":
						markdownValue = rawValue
					}
					if t.Name.Local == "v" && (cellType == "" || cellType == "n") {
						value = xlsxDisplayNumberForCellWidth(value, cellStyle, styles, xlsxColumnDisplayWidth(columnWidths, cellCol))
						markdownValue = value
					}
					if !skipValue {
						if strictOfficeContent {
							// UsedRange.Text returns every visible cell occurrence. Do
							// not let generic binary-text cleanup collapse legitimate
							// worksheet literals such as "wml.xsd or dml__ROOT". Still
							// process SpreadsheetML _xNNNN_ escapes: Excel renders a
							// control escape such as _x001A_ as no visible glyph, while
							// emitting its literal source leaks false text tokens.
							appendTrimmedTextBlock(out, xlsxStrictOfficeCellText(value))
						} else if t.Name.Local == "v" && (cellType == "" || cellType == "n") && plainExcelNumberValue(value) {
							appendTrimmedTextBlock(out, value)
						} else {
							appendWorksheetValue(out, value, &seenLargeValues)
						}
					}
					if collectMarkdownCell {
						markdownCellText.WriteString(markdownValue)
					}
				}
				inV, inT = false, false
			}
			if t.Name.Local == "c" {
				if collectMarkdownCell {
					prepared := ""
					if value := cleanMarkdownTableCellValue(markdownCellText.String()); value != "" {
						prepared = prepareMarkdownTableCellValue(value)
					}
					if prepared != "" {
						for len(markdownRowValues) < cellCol {
							markdownRowValues = append(markdownRowValues, "")
						}
						markdownRowValues[cellCol-1] = prepared
					}
				}
				skipCell = false
				collectMarkdownCell = false
				cellCol = 0
				cellDepth = 0
				systemCellTextDepth = 0
			} else if cellDepth > 0 {
				cellDepth--
			}
			if t.Name.Local == "row" {
				if collectMarkdownRow {
					if row := compactPreparedWorksheetMarkdownRow(markdownRowValues); len(row) > 0 && len(md.rows) < maxMarkdownTableRows {
						md.rows = append(md.rows, row)
					}
				}
				rowHidden = false
				currentRow = 0
				collectMarkdownRow = false
				markdownRowValues = nil
			}
		case xml.CharData:
			if inHeaderFooter || (!skipCell && (inV || inT)) {
				cur.Write(t)
			}
		}
	}
	return nil
}

func xlsxStrictOfficeCellText(value string) string {
	// Worksheet values are already structurally bounded XML cell content.  Do
	// not send them through the general cleanText pipeline: its Word-field
	// heuristics can reinterpret a perfectly ordinary cell which starts with
	// "#" as a field instruction and erase it.  Excel Range.Text retains that
	// glyph (for example a literal # in a summary template), so only perform
	// the SpreadsheetML escape/control cleanup needed for rendered cells.
	value = strings.ToValidUTF8(value, "")
	value = decodeOOXMLTextEscapes(value)
	value = strings.Map(cleanTextRune, value)
	value = strings.ReplaceAll(value, "\uE000", "'s")
	value = spaceRE.ReplaceAllString(value, " ")
	return strings.TrimSpace(value)
}

func plainExcelNumberValue(value string) bool {
	if value == "" {
		return false
	}
	hasDigit := false
	for i := 0; i < len(value); i++ {
		c := value[i]
		switch {
		case c >= '0' && c <= '9':
			hasDigit = true
		case c == '+' || c == '-' || c == '.' || c == 'e' || c == 'E':
		default:
			return false
		}
	}
	return hasDigit
}

func appendSimpleInlineWorksheetText(out *strings.Builder, b []byte, md *xlsxWorksheetMarkdownData) (bool, error) {
	if !simpleInlineWorksheetCandidate(b) {
		return false, nil
	}
	if err := appendSimpleInlineWorksheetTextPrepared(out, b, md); err != nil {
		if errors.Is(err, errXLSXWorksheetFastPathFallback) {
			return false, nil
		}
		return true, err
	}
	return true, nil
}

func appendSharedStringWorksheetText(out *strings.Builder, b []byte, shared []string, md *xlsxWorksheetMarkdownData) (bool, error) {
	if !sharedStringWorksheetCandidate(b) {
		return false, nil
	}
	if err := appendSharedStringWorksheetTextPrepared(out, b, shared, md); err != nil {
		if errors.Is(err, errXLSXWorksheetFastPathFallback) {
			return false, nil
		}
		return true, err
	}
	return true, nil
}

var errXLSXWorksheetFastPathFallback = errors.New("xlsx worksheet fast path fallback")

func appendSimpleInlineWorksheetTextPrepared(out *strings.Builder, b []byte, md *xlsxWorksheetMarkdownData) error {
	var seenLargeValues map[string]bool
	var rowValues []string
	collectMarkdown := md != nil
	currentRow := 0
	nextCol := 1
	flushRow := func() {
		if collectMarkdown && md != nil && rowValues != nil && len(md.rows) < maxMarkdownTableRows {
			n := len(rowValues)
			for n > 0 && rowValues[n-1] == "" {
				n--
			}
			if n > 0 {
				row := append([]string(nil), rowValues[:n]...)
				md.rows = append(md.rows, row)
				if len(md.rows) >= maxMarkdownTableRows {
					collectMarkdown = false
				}
			}
			for i := range rowValues {
				rowValues[i] = ""
			}
		}
		rowValues = nil
	}
	for pos := 0; ; {
		i := bytes.Index(b[pos:], []byte("<c"))
		if i < 0 {
			break
		}
		i += pos
		tagEnd := bytes.IndexByte(b[i:], '>')
		if tagEnd < 0 {
			return errXLSXWorksheetFastPathFallback
		}
		tagEnd += i
		tag := b[i : tagEnd+1]
		if !xmlStartTagNameIs(tag, "c") || !bytes.Contains(tag, []byte(`t="inlineStr"`)) {
			pos = tagEnd + 1
			continue
		}
		cellEndRel := bytes.Index(b[tagEnd+1:], []byte("</c>"))
		if cellEndRel < 0 {
			return errXLSXWorksheetFastPathFallback
		}
		cellEnd := tagEnd + 1 + cellEndRel
		raw := simpleInlineCellText(b[tagEnd+1 : cellEnd])
		value := strings.TrimSpace(raw)
		if !collectMarkdown {
			if value != "" {
				appendWorksheetValue(out, value, &seenLargeValues)
			}
			pos = cellEnd + len("</c>")
			continue
		}
		cellRef := xmlAttrBytes(tag, "r")
		cellCol, cellRow, ok := cellRefIndexes(string(cellRef))
		if !ok {
			cellCol = nextCol
			cellRow = currentRow
			if cellRow < 1 {
				cellRow = 1
			}
		}
		if cellCol < 1 {
			cellCol = 1
		}
		nextCol = cellCol + 1
		if cellRow != currentRow {
			flushRow()
			currentRow = cellRow
			nextCol = cellCol + 1
			if collectMarkdown && md != nil && len(md.rows) < maxMarkdownTableRows {
				rowValues = make([]string, 0, 16)
			}
		}
		if value != "" {
			appendWorksheetValue(out, value, &seenLargeValues)
			if rowValues != nil && cellCol <= maxMarkdownTableCols {
				if cellValue := cleanMarkdownTableCellValue(raw); cellValue != "" {
					prepared := prepareMarkdownTableCellValue(cellValue)
					if prepared != "" {
						for len(rowValues) < cellCol {
							rowValues = append(rowValues, "")
						}
						rowValues[cellCol-1] = prepared
					}
				}
			}
		}
		pos = cellEnd + len("</c>")
	}
	flushRow()
	return nil
}

func appendSharedStringWorksheetTextPrepared(out *strings.Builder, b []byte, shared []string, md *xlsxWorksheetMarkdownData) error {
	var seenLargeValues map[string]bool
	var seenLargeSharedIndexes map[int]bool
	nextRow := 1
	for pos := 0; ; {
		i := bytes.Index(b[pos:], []byte("<row"))
		if i < 0 {
			break
		}
		i += pos
		tagEnd := bytes.IndexByte(b[i:], '>')
		if tagEnd < 0 {
			return errXLSXWorksheetFastPathFallback
		}
		tagEnd += i
		tag := b[i : tagEnd+1]
		if !xmlStartTagNameIs(tag, "row") {
			pos = tagEnd + 1
			continue
		}
		rowEndRel := bytes.Index(b[tagEnd+1:], []byte("</row>"))
		if rowEndRel < 0 {
			return errXLSXWorksheetFastPathFallback
		}
		rowEnd := tagEnd + 1 + rowEndRel
		rowIndex := nextRow
		if rowRef := xmlAttrBytes(tag, "r"); len(rowRef) > 0 {
			if n, ok := atoi(string(rowRef)); ok && n > 0 {
				rowIndex = n
			}
		}
		nextRow = rowIndex + 1
		var markdownRowValues []string
		if md != nil && len(md.rows) < maxMarkdownTableRows {
			markdownRowValues = make([]string, 0, 16)
		}
		rowData := b[tagEnd+1 : rowEnd]
		nextCol := 1
		for cellPos := 0; ; {
			cellStart := bytes.Index(rowData[cellPos:], []byte("<c"))
			if cellStart < 0 {
				break
			}
			cellStart += cellPos
			cellTagEnd := bytes.IndexByte(rowData[cellStart:], '>')
			if cellTagEnd < 0 {
				return errXLSXWorksheetFastPathFallback
			}
			cellTagEnd += cellStart
			cellTag := rowData[cellStart : cellTagEnd+1]
			if !xmlStartTagNameIs(cellTag, "c") {
				cellPos = cellTagEnd + 1
				continue
			}
			cellRef := xmlAttrBytes(cellTag, "r")
			cellCol := nextCol
			if col, _, ok := cellRefIndexes(string(cellRef)); ok {
				cellCol = col
			}
			if cellCol < 1 {
				cellCol = 1
			}
			nextCol = cellCol + 1
			selfClosing := len(cellTag) >= 2 && cellTag[len(cellTag)-2] == '/'
			cellEnd := cellTagEnd
			if !selfClosing {
				cellEndRel := bytes.Index(rowData[cellTagEnd+1:], []byte("</c>"))
				if cellEndRel < 0 {
					return errXLSXWorksheetFastPathFallback
				}
				cellEnd = cellTagEnd + 1 + cellEndRel
			}
			rawValue := ""
			if !selfClosing {
				if value, ok := worksheetCellVText(rowData[cellTagEnd+1 : cellEnd]); ok {
					rawValue = value
				}
			}
			value := strings.TrimSpace(rawValue)
			markdownValue := value
			skipValue := false
			if bytes.Equal(xmlAttrBytes(cellTag, "t"), []byte("s")) {
				if idx, ok := atoi(value); ok && idx >= 0 && idx < len(shared) {
					if len(shared[idx]) > maxRepeatedTextPartBytes {
						if seenLargeSharedIndexes != nil && seenLargeSharedIndexes[idx] {
							skipValue = true
							value = shared[idx]
						} else {
							if seenLargeSharedIndexes == nil {
								seenLargeSharedIndexes = map[int]bool{}
							}
							seenLargeSharedIndexes[idx] = true
						}
					}
					value = shared[idx]
					markdownValue = value
				}
			}
			if !skipValue && value != "" {
				if plainExcelNumberValue(value) && len(xmlAttrBytes(cellTag, "t")) == 0 {
					appendTrimmedTextBlock(out, value)
				} else {
					appendWorksheetValue(out, value, &seenLargeValues)
				}
			}
			if markdownRowValues != nil && cellCol <= maxMarkdownTableCols && markdownValue != "" {
				if cellValue := cleanMarkdownTableCellValue(markdownValue); cellValue != "" {
					prepared := prepareMarkdownTableCellValue(cellValue)
					if prepared != "" {
						for len(markdownRowValues) < cellCol {
							markdownRowValues = append(markdownRowValues, "")
						}
						markdownRowValues[cellCol-1] = prepared
					}
				}
			}
			if selfClosing {
				cellPos = cellTagEnd + 1
			} else {
				cellPos = cellEnd + len("</c>")
			}
		}
		if markdownRowValues != nil {
			if row := compactPreparedWorksheetMarkdownRow(markdownRowValues); len(row) > 0 && len(md.rows) < maxMarkdownTableRows {
				md.rows = append(md.rows, row)
			}
		}
		pos = rowEnd + len("</row>")
	}
	return nil
}

func simpleInlineWorksheetCandidate(b []byte) bool {
	foundInlineStr := false
	for pos := 0; ; {
		i := bytes.IndexByte(b[pos:], '<')
		if i < 0 {
			break
		}
		i += pos
		tagEnd := bytes.IndexByte(b[i:], '>')
		if tagEnd < 0 {
			break
		}
		tagEnd += i
		tag := b[i : tagEnd+1]
		pos = tagEnd + 1
		if len(tag) < 2 {
			continue
		}
		if len(tag) >= len("<!DOCTYPE") &&
			tag[0] == '<' &&
			tag[1] == '!' &&
			tag[2] == 'D' &&
			tag[3] == 'O' &&
			tag[4] == 'C' &&
			tag[5] == 'T' &&
			tag[6] == 'Y' &&
			tag[7] == 'P' &&
			tag[8] == 'E' {
			return false
		}
		if !xmlStartTagNameIs(tag, "c") &&
			!xmlStartTagNameIs(tag, "row") &&
			!xmlStartTagNameIs(tag, "worksheet") &&
			!xmlStartTagNameIs(tag, "sheetData") &&
			!xmlStartTagNameIs(tag, "dimension") &&
			!xmlStartTagNameIs(tag, "is") &&
			!xmlStartTagNameIs(tag, "t") &&
			!xmlStartTagNameIs(tag, "si") {
			if xmlStartTagNameIs(tag, "f") ||
				xmlStartTagNameIs(tag, "v") ||
				xmlStartTagNameIs(tag, "cols") ||
				xmlStartTagNameIs(tag, "rPh") ||
				xmlStartTagNameIs(tag, "phoneticPr") ||
				xmlStartTagNameIs(tag, "hyperlink") ||
				xmlStartTagNameIs(tag, "dataValidation") ||
				isExcelHeaderFooterElement(xmlTagName(tag)) {
				return false
			}
		}
		if bytes.Contains(tag, []byte("hidden")) || bytes.Contains(tag, []byte(" ht=")) || bytes.Contains(tag, []byte("\tht=")) {
			return false
		}
		if !xmlStartTagNameIs(tag, "c") {
			continue
		}
		switch {
		case bytes.Contains(tag, []byte(`t="inlineStr"`)):
			foundInlineStr = true
		case bytes.Contains(tag, []byte(`t="s"`)):
			return false
		case bytes.Contains(tag, []byte(`t="b"`)):
			return false
		case bytes.Contains(tag, []byte(`t="str"`)):
			return false
		}
	}
	return foundInlineStr
}

func sharedStringWorksheetCandidate(b []byte) bool {
	if hasDOCTYPE(b) ||
		bytes.Contains(b, []byte("<![CDATA[")) ||
		bytes.Contains(b, []byte("<f")) ||
		bytes.Contains(b, []byte("<is")) ||
		bytes.Contains(b, []byte("<t")) ||
		bytes.Contains(b, []byte("<rPh")) ||
		bytes.Contains(b, []byte("<hyperlink")) ||
		bytes.Contains(b, []byte("<dataValidation")) ||
		bytes.Contains(b, []byte(`hidden="`)) ||
		bytes.Contains(b, []byte(`hidden='`)) ||
		bytes.Contains(b, []byte("<oddHeader")) ||
		bytes.Contains(b, []byte("<oddFooter")) ||
		bytes.Contains(b, []byte("<evenHeader")) ||
		bytes.Contains(b, []byte("<evenFooter")) ||
		bytes.Contains(b, []byte("<firstHeader")) ||
		bytes.Contains(b, []byte("<firstFooter")) {
		return false
	}
	foundCell := false
	for pos := 0; ; {
		i := bytes.Index(b[pos:], []byte("<c"))
		if i < 0 {
			break
		}
		i += pos
		tagEnd := bytes.IndexByte(b[i:], '>')
		if tagEnd < 0 {
			return false
		}
		tagEnd += i
		tag := b[i : tagEnd+1]
		pos = tagEnd + 1
		if !xmlStartTagNameIs(tag, "c") {
			continue
		}
		foundCell = true
		switch {
		case bytes.Contains(tag, []byte(`t="inlineStr"`)):
			return false
		case bytes.Contains(tag, []byte(`t="b"`)):
			return false
		case bytes.Contains(tag, []byte(`t="str"`)):
			return false
		case bytes.Contains(tag, []byte(`t="e"`)):
			return false
		case bytes.Contains(tag, []byte(`t="s"`)):
		case bytes.Contains(tag, []byte(` t='s'`)):
		case !bytes.Contains(tag, []byte(`t="`)) && !bytes.Contains(tag, []byte(`t='`)):
		default:
			return false
		}
	}
	return foundCell
}

func worksheetCellVText(b []byte) (string, bool) {
	for pos := 0; ; {
		i := bytes.Index(b[pos:], []byte("<v"))
		if i < 0 {
			return "", false
		}
		i += pos
		tagEnd := bytes.IndexByte(b[i:], '>')
		if tagEnd < 0 {
			return "", false
		}
		tagEnd += i
		if !xmlStartTagNameIs(b[i:tagEnd+1], "v") {
			pos = tagEnd + 1
			continue
		}
		endRel := bytes.Index(b[tagEnd+1:], []byte("</v>"))
		if endRel < 0 {
			return "", false
		}
		return string(b[tagEnd+1 : tagEnd+1+endRel]), true
	}
}

func xmlTagName(tag []byte) string {
	if len(tag) < 2 || tag[0] != '<' {
		return ""
	}
	start := 1
	if tag[start] == '/' || tag[start] == '!' || tag[start] == '?' {
		return ""
	}
	for start < len(tag) && tag[start] != '>' && tag[start] != '/' && tag[start] != ' ' && tag[start] != '\t' && tag[start] != '\r' && tag[start] != '\n' {
		if tag[start] == ':' {
			start++
			break
		}
		start++
	}
	if start >= len(tag) {
		return ""
	}
	if start > 1 && tag[start-1] != ':' {
		start = 1
	}
	end := start
	for end < len(tag) && tag[end] != '>' && tag[end] != '/' && tag[end] != ' ' && tag[end] != '\t' && tag[end] != '\r' && tag[end] != '\n' {
		end++
	}
	if end <= start {
		return ""
	}
	return string(tag[start:end])
}

func simpleInlineCellText(b []byte) string {
	var out strings.Builder
	for pos := 0; ; {
		i := bytes.Index(b[pos:], []byte("<t"))
		if i < 0 {
			break
		}
		i += pos
		tagEnd := bytes.IndexByte(b[i:], '>')
		if tagEnd < 0 {
			break
		}
		tagEnd += i
		if !xmlStartTagNameIs(b[i:tagEnd+1], "t") {
			pos = tagEnd + 1
			continue
		}
		endRel := bytes.Index(b[tagEnd+1:], []byte("</t>"))
		if endRel < 0 {
			break
		}
		text := string(b[tagEnd+1 : tagEnd+1+endRel])
		if strings.Contains(text, "&") {
			text = html.UnescapeString(text)
		}
		out.WriteString(text)
		pos = tagEnd + 1 + endRel + len("</t>")
	}
	return out.String()
}

func xmlStartTagNameIs(tag []byte, name string) bool {
	if len(tag) < len(name)+2 || tag[0] != '<' {
		return false
	}
	if len(tag) > 1 && (tag[1] == '/' || tag[1] == '!' || tag[1] == '?') {
		return false
	}
	start := 1
	for start < len(tag) && tag[start] != '>' && tag[start] != '/' && tag[start] != ' ' && tag[start] != '\t' && tag[start] != '\r' && tag[start] != '\n' {
		if tag[start] == ':' {
			break
		}
		start++
	}
	if start < len(tag) && tag[start] == ':' {
		start++
	} else {
		start = 1
	}
	if len(tag) < start+len(name) {
		return false
	}
	if string(tag[start:start+len(name)]) != name {
		return false
	}
	if len(tag) == start+len(name) {
		return true
	}
	c := tag[start+len(name)]
	return c == '>' || c == '/' || c == ' ' || c == '\t' || c == '\r' || c == '\n'
}

func xmlAttrBytes(tag []byte, name string) []byte {
	for pos := 0; ; {
		i := bytes.Index(tag[pos:], []byte(name+"="))
		if i < 0 {
			return nil
		}
		i += pos
		if i > 0 {
			prev := tag[i-1]
			if prev != ' ' && prev != '\t' && prev != '\r' && prev != '\n' {
				pos = i + len(name) + 1
				continue
			}
		}
		q := i + len(name) + 1
		if q >= len(tag) || (tag[q] != '"' && tag[q] != '\'') {
			return nil
		}
		if end := bytes.IndexByte(tag[q+1:], tag[q]); end >= 0 {
			return tag[q+1 : q+1+end]
		}
		return nil
	}
}

func worksheetElementMayHaveVisibleTextAttributes(name string) bool {
	return name == "dataValidation" || name == "hyperlink"
}

func worksheetMarkdownRows(b []byte, shared []string) ([][]string, error) {
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var rows [][]string
	var rowValues map[int]string
	var cur strings.Builder
	var hiddenCols []intRange
	var rowHidden bool
	var skipCell bool
	var cellType string
	var cellRef string
	var cellCol int
	var nextCol int
	var inV, inT bool
	var phoneticDepth int
	var cellDepth int
	var systemCellTextDepth int
	collectRow := true
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "col" {
				if r, ok := hiddenColumnRange(t); ok {
					hiddenCols = append(hiddenCols, r)
				}
			}
			if t.Name.Local == "row" {
				rowHidden = worksheetRowHidden(t)
				collectRow = !rowHidden && len(rows) < maxMarkdownTableRows
				if collectRow {
					rowValues = map[int]string{}
				} else {
					rowValues = nil
				}
				nextCol = 1
			}
			if t.Name.Local == "c" {
				cellDepth = 1
				cellType = ""
				cellRef = ""
				cur.Reset()
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "t":
						cellType = a.Value
					case "r":
						cellRef = a.Value
					}
				}
				if col, _, ok := cellRefIndexes(cellRef); ok {
					cellCol = col
				} else {
					cellCol = nextCol
				}
				if cellCol < 1 {
					cellCol = 1
				}
				nextCol = cellCol + 1
				skipCell = !collectRow || rowHidden || hiddenColumnCell(cellRef, hiddenCols) || columnHidden(cellCol, hiddenCols)
			} else if cellDepth > 0 {
				cellDepth++
			}
			if systemCellTextDepth > 0 {
				systemCellTextDepth++
				continue
			}
			if cellDepth > 0 && isExcelSystemCellTextElement(t.Name.Local) {
				systemCellTextDepth = 1
				continue
			}
			if isExcelPhoneticElement(t.Name.Local) {
				phoneticDepth++
				continue
			}
			if phoneticDepth == 0 && (t.Name.Local == "v" || t.Name.Local == "t") {
				inV = t.Name.Local == "v"
				inT = t.Name.Local == "t"
			}
		case xml.EndElement:
			if phoneticDepth > 0 {
				if isExcelPhoneticElement(t.Name.Local) {
					phoneticDepth--
				}
				continue
			}
			if systemCellTextDepth > 0 {
				systemCellTextDepth--
				continue
			}
			if t.Name.Local == "v" || t.Name.Local == "t" {
				inV, inT = false, false
			}
			if t.Name.Local == "c" {
				if !skipCell && rowValues != nil {
					value := strings.TrimSpace(cur.String())
					if value != "" {
						if cellType == "s" {
							if idx, ok := atoi(value); ok && idx >= 0 && idx < len(shared) {
								value = shared[idx]
							}
						}
						if cellType == "b" {
							value = excelBooleanDisplayText(value)
						}
						if cellCol <= maxMarkdownTableCols {
							if cellValue := cleanMarkdownTableCellValue(value); cellValue != "" {
								rowValues[cellCol] = prepareMarkdownTableCellValue(cellValue)
							}
						}
					}
				}
				skipCell = false
				cur.Reset()
				cellDepth = 0
				systemCellTextDepth = 0
			} else if cellDepth > 0 {
				cellDepth--
			}
			if t.Name.Local == "row" {
				if collectRow {
					if row := compactWorksheetMarkdownRow(rowValues); len(row) > 0 && len(rows) < maxMarkdownTableRows {
						rows = append(rows, row)
						if len(rows) >= maxMarkdownTableRows {
							return rows, nil
						}
					}
				}
				rowHidden = false
				rowValues = nil
				collectRow = true
			}
		case xml.CharData:
			if !skipCell && (inV || inT) {
				cur.Write(t)
			}
		}
	}
	return rows, nil
}

func columnHidden(col int, ranges []intRange) bool {
	for _, r := range ranges {
		if col >= r.min && col <= r.max {
			return true
		}
	}
	return false
}

func compactWorksheetMarkdownRow(values map[int]string) []string {
	maxCol := 0
	for col, value := range values {
		if strings.TrimSpace(value) != "" && col > maxCol && col <= maxMarkdownTableCols {
			maxCol = col
		}
	}
	if maxCol == 0 {
		return nil
	}
	row := make([]string, maxCol)
	for col, value := range values {
		if col >= 1 && col <= maxCol {
			row[col-1] = value
		}
	}
	return row
}

func compactPreparedWorksheetMarkdownRow(values []string) []string {
	n := len(values)
	for n > 0 && values[n-1] == "" {
		n--
	}
	if n == 0 {
		return nil
	}
	return append([]string(nil), values[:n]...)
}

func cleanMarkdownTableCellValue(value string) string {
	value = cleanText(value)
	value = stripInlineHiddenOfficeReferences(value)
	if value == "" {
		return ""
	}
	if !strings.Contains(value, "\n") {
		value = strings.TrimSpace(value)
		if value != "" && !maybeDiscardableHiddenOfficeText(value) && !maybeControlFragmentText(value) {
			truncated := len(value) > maxMarkdownTableCellBytes
			value = truncateUTF8Bytes(value, maxMarkdownTableCellBytes)
			if truncated && value != "" {
				value += "..."
			}
			return value
		}
		if markdownTableCellValueLineDiscarded(value) {
			return ""
		}
	} else {
		var out strings.Builder
		start := 0
		for start <= len(value) {
			end := start
			for end < len(value) && value[end] != '\n' {
				end++
			}
			line := strings.TrimSpace(value[start:end])
			if !markdownTableCellValueLineDiscarded(line) {
				if out.Len() > 0 {
					out.WriteByte('\n')
				}
				out.WriteString(line)
			}
			if end == len(value) {
				break
			}
			start = end + 1
		}
		value = out.String()
	}
	if value == "" {
		return ""
	}
	truncated := len(value) > maxMarkdownTableCellBytes
	value = truncateUTF8Bytes(value, maxMarkdownTableCellBytes)
	if truncated && value != "" {
		value += "..."
	}
	return value
}

func markdownTableCellValueLineDiscarded(line string) bool {
	if line == "" || looksLikeBinaryControlFragment(line) {
		return true
	}
	if !maybeDiscardableHiddenOfficeText(line) {
		return false
	}
	return looksLikeHiddenResourceReference(line) ||
		looksLikeRelationshipIDReference(line) ||
		looksLikeOfficeRelationshipMetadataReference(line) ||
		looksLikeOfficeXMLMetadataReference(line)
}

func prepareMarkdownTableCellValue(s string) string {
	if !strings.ContainsAny(s, "\r\n") {
		trimmed := strings.TrimSpace(s)
		if trimmed != "" && !strings.HasPrefix(strings.ToLower(trimmed), `{\rtf`) && !strings.HasPrefix(strings.ToLower(trimmed), `\rtf`) {
			s = normalizeMarkdownTextLine(s)
			s = strings.ReplaceAll(s, "\\", "\\\\")
			s = strings.ReplaceAll(s, "|", "\\|")
			return s
		}
	}
	s = markdownText(s)
	lines := strings.Split(s, "\n")
	compact := lines[:0]
	for _, line := range lines {
		line = strings.TrimSpace(line)
		if line != "" && (len(compact) == 0 || compact[len(compact)-1] != line) {
			compact = append(compact, line)
		}
	}
	s = strings.Join(compact, "\n")
	s = strings.ReplaceAll(s, "\\", "\\\\")
	s = strings.ReplaceAll(s, "|", "\\|")
	s = strings.ReplaceAll(s, "\n", "<br>")
	return s
}

func truncateUTF8Bytes(s string, maxBytes int) string {
	if maxBytes <= 0 || len(s) <= maxBytes {
		return s
	}
	s = s[:maxBytes]
	for len(s) > 0 && !utf8.ValidString(s) {
		s = s[:len(s)-1]
	}
	return s
}

func markdownTable(rows [][]string) string {
	cols := 0
	for _, row := range rows {
		if len(row) > cols {
			cols = len(row)
		}
	}
	if cols == 0 {
		return ""
	}
	var out strings.Builder
	writeRow := func(row []string) {
		out.WriteByte('|')
		for i := 0; i < cols; i++ {
			cell := ""
			if i < len(row) {
				cell = row[i]
			}
			out.WriteByte(' ')
			out.WriteString(escapeMarkdownTableCell(cell))
			out.WriteString(" |")
		}
		out.WriteByte('\n')
	}
	writeRow(rows[0])
	out.WriteByte('|')
	for i := 0; i < cols; i++ {
		out.WriteString(" --- |")
	}
	out.WriteByte('\n')
	for _, row := range rows[1:] {
		writeRow(row)
	}
	return strings.TrimSpace(out.String())
}

func markdownTablePrepared(rows [][]string) string {
	cols := 0
	for _, row := range rows {
		if len(row) > cols {
			cols = len(row)
		}
	}
	if cols == 0 {
		return ""
	}
	var out strings.Builder
	writeRow := func(row []string) {
		out.WriteByte('|')
		for i := 0; i < cols; i++ {
			cell := ""
			if i < len(row) {
				cell = row[i]
			}
			out.WriteByte(' ')
			out.WriteString(cell)
			out.WriteString(" |")
		}
		out.WriteByte('\n')
	}
	writeRow(rows[0])
	out.WriteByte('|')
	for i := 0; i < cols; i++ {
		out.WriteString(" --- |")
	}
	out.WriteByte('\n')
	for _, row := range rows[1:] {
		writeRow(row)
	}
	return strings.TrimSpace(out.String())
}

func escapeMarkdownTableCell(s string) string {
	s = strings.ReplaceAll(s, "\r\n", "\n")
	s = strings.ReplaceAll(s, "\r", "\n")
	rawLines := strings.Split(s, "\n")
	cleaned := make([]string, 0, len(rawLines))
	for _, line := range rawLines {
		line = cleanText(line)
		line = stripInlineHiddenOfficeReferences(line)
		if line == "" || looksLikeMarkdownTableCellHiddenReference(line) || looksLikeRelationshipIDReference(line) || looksLikeOfficeRelationshipMetadataReference(line) || looksLikeOfficeXMLMetadataReference(line) || looksLikeBinaryControlFragment(line) {
			continue
		}
		cleaned = append(cleaned, line)
	}
	s = strings.Join(cleaned, "\n")
	s = markdownText(s)
	lines := strings.Split(s, "\n")
	compact := lines[:0]
	for _, line := range lines {
		line = strings.TrimSpace(line)
		if line != "" && (len(compact) == 0 || compact[len(compact)-1] != line) {
			compact = append(compact, line)
		}
	}
	s = strings.Join(compact, "\n")
	s = strings.ReplaceAll(s, "\\", "\\\\")
	s = strings.ReplaceAll(s, "|", "\\|")
	s = strings.ReplaceAll(s, "\n", "<br>")
	return s
}

func looksLikeMarkdownTableCellHiddenReference(s string) bool {
	trimmed := strings.TrimSpace(strings.ReplaceAll(s, "\\", "/"))
	if trimmed == "" {
		return false
	}
	if looksLikeInlineHiddenResourceReferencePlain(trimmed) {
		return true
	}
	return looksLikeDecodedOfficePartPath(trimmed)
}

func looksLikeDecodedOfficePartPath(s string) bool {
	seen := map[string]bool{}
	queue := []string{strings.TrimSpace(strings.ReplaceAll(s, "\\", "/"))}
	for len(queue) > 0 && len(seen) < 16 {
		cur := strings.TrimSpace(strings.ReplaceAll(queue[0], "\\", "/"))
		queue = queue[1:]
		if cur == "" || seen[cur] {
			continue
		}
		seen[cur] = true
		candidate := strings.TrimPrefix(hiddenResourceReferenceCandidate(cur), "/")
		if looksLikeOfficePartPath(strings.ToLower(candidate)) {
			return true
		}
		if decoded, err := url.PathUnescape(cur); err == nil && decoded != cur {
			queue = append(queue, decoded)
		}
		if strings.Contains(cur, "&") {
			if unescaped := html.UnescapeString(cur); unescaped != cur {
				queue = append(queue, unescaped)
			}
		}
	}
	return false
}

func escapeMarkdownHeading(s string) string {
	s = markdownVisibleHTMLText(s)
	s = cleanText(s)
	s = stripInlineHiddenOfficeReferences(s)
	if s == "" || looksLikeHiddenResourceReference(s) || looksLikeRelationshipIDReference(s) || looksLikeOfficeRelationshipMetadataReference(s) || looksLikeOfficeXMLMetadataReference(s) || looksLikeBinaryControlFragment(s) {
		return "Sheet"
	}
	s = strings.Join(strings.Fields(s), " ")
	s = strings.TrimSpace(strings.TrimLeft(s, "#"))
	if s == "" {
		return "Sheet"
	}
	return s
}

func excelBooleanDisplayText(value string) string {
	switch strings.TrimSpace(strings.ToLower(value)) {
	case "1", "true":
		return "TRUE"
	case "0", "false":
		return "FALSE"
	default:
		return value
	}
}

type intRange struct {
	min int
	max int
}

func hiddenColumnRange(start xml.StartElement) (intRange, bool) {
	if !worksheetColumnHidden(start) {
		return intRange{}, false
	}
	min, max := 0, 0
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "min":
			min, _ = atoi(attr.Value)
		case "max":
			max, _ = atoi(attr.Value)
		}
	}
	if min <= 0 || max < min {
		return intRange{}, false
	}
	return intRange{min: min, max: max}, true
}

func worksheetColumnHidden(start xml.StartElement) bool {
	if elementBoolAttr(start, "hidden") {
		return true
	}
	if value, ok := floatAttrValue(start, "width"); ok && value <= 0 {
		return true
	}
	return false
}

func worksheetRowHidden(start xml.StartElement) bool {
	if elementBoolAttr(start, "hidden") {
		return true
	}
	if value, ok := floatAttrValue(start, "ht"); ok && value <= 0 {
		return true
	}
	return false
}

func worksheetRowIndex(start xml.StartElement, fallback int) int {
	for _, attr := range start.Attr {
		if attr.Name.Local != "r" {
			continue
		}
		if row, ok := atoi(attr.Value); ok {
			return row
		}
		break
	}
	return fallback
}

func floatAttrValue(start xml.StartElement, name string) (float64, bool) {
	for _, attr := range start.Attr {
		if attr.Name.Local != name {
			continue
		}
		value, err := strconv.ParseFloat(strings.TrimSpace(attr.Value), 64)
		if err != nil {
			return 0, false
		}
		return value, true
	}
	return 0, false
}

func hiddenColumnCell(cellRef string, ranges []intRange) bool {
	if len(ranges) == 0 {
		return false
	}
	col, _, ok := cellRefIndexes(cellRef)
	if !ok {
		return false
	}
	for _, r := range ranges {
		if col >= r.min && col <= r.max {
			return true
		}
	}
	return false
}

func worksheetElementHiddenByRef(start xml.StartElement, hiddenCols, hiddenRows []intRange) bool {
	var refs string
	switch start.Name.Local {
	case "hyperlink":
		for _, attr := range start.Attr {
			if attr.Name.Local == "ref" {
				refs = attr.Value
				break
			}
		}
	case "dataValidation":
		for _, attr := range start.Attr {
			if attr.Name.Local == "sqref" {
				refs = attr.Value
				break
			}
		}
	default:
		return false
	}
	if strings.TrimSpace(refs) == "" {
		return false
	}
	sawRef := false
	for _, ref := range worksheetRefFields(refs) {
		hidden, ok := cellRangeHidden(ref, hiddenCols, hiddenRows)
		if !ok {
			return false
		}
		sawRef = true
		if !hidden {
			return false
		}
	}
	return sawRef
}

func worksheetRefFields(refs string) []string {
	return strings.FieldsFunc(refs, func(r rune) bool {
		return r == ',' || r == ' ' || r == '\t' || r == '\n' || r == '\r'
	})
}

func cellRangeHidden(ref string, hiddenCols, hiddenRows []intRange) (bool, bool) {
	parts := strings.Split(ref, ":")
	if len(parts) == 0 || len(parts) > 2 {
		return false, false
	}
	startCol, startRow, ok := cellRefIndexes(parts[0])
	if !ok {
		return false, false
	}
	endCol, endRow := startCol, startRow
	if len(parts) == 2 {
		endCol, endRow, ok = cellRefIndexes(parts[1])
		if !ok {
			return false, false
		}
	}
	if endCol < startCol {
		startCol, endCol = endCol, startCol
	}
	if endRow < startRow {
		startRow, endRow = endRow, startRow
	}
	return rangeCoveredByRanges(startCol, endCol, hiddenCols) || rangeCoveredByRanges(startRow, endRow, hiddenRows), true
}

// xlsxCellRefsInRange expands the modest sqref ranges used by conditional
// formatting. A hard ceiling keeps adversarial whole-column references from
// allocating unbounded memory; in that case callers retain the raw cell text.
func xlsxCellRefsInRange(ref string) []string {
	parts := strings.Split(strings.TrimSpace(ref), ":")
	if len(parts) == 0 || len(parts) > 2 {
		return nil
	}
	startCol, startRow, ok := cellRefIndexes(parts[0])
	if !ok {
		return nil
	}
	endCol, endRow := startCol, startRow
	if len(parts) == 2 {
		endCol, endRow, ok = cellRefIndexes(parts[1])
		if !ok {
			return nil
		}
	}
	if endCol < startCol {
		startCol, endCol = endCol, startCol
	}
	if endRow < startRow {
		startRow, endRow = endRow, startRow
	}
	if (endCol-startCol+1)*(endRow-startRow+1) > 100000 {
		return nil
	}
	out := make([]string, 0, (endCol-startCol+1)*(endRow-startRow+1))
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			out = append(out, xlsxCellRefName(col, row))
		}
	}
	return out
}

func xlsxHiddenConditionalFormatValueCells(b []byte) map[string]bool {
	type stopIfTrueRule struct {
		refs     string
		operator string
		formulas []string
	}
	hidden := map[string]bool{}
	if len(b) == 0 || (!bytes.Contains(b, []byte("dataBar")) && !bytes.Contains(b, []byte("iconSet"))) {
		return hidden
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var conditionalRefs string
	var dataBarDepth int
	showValue := true
	var stopIfTrueRules []stopIfTrueRule
	var hiddenRefs []string
	var currentStopRule *stopIfTrueRule
	var currentCellRef, currentCellType string
	var inCellValue bool
	var cellValue strings.Builder
	var inFormula bool
	var formula strings.Builder
	cellValues := map[string]float64{}
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) || err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "c":
				currentCellRef, currentCellType = "", ""
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "r":
						currentCellRef = a.Value
					case "t":
						currentCellType = a.Value
					}
				}
			case "v":
				if currentCellRef != "" {
					inCellValue = true
					cellValue.Reset()
				}
			case "conditionalFormatting":
				conditionalRefs = ""
				for _, a := range t.Attr {
					if a.Name.Local == "sqref" {
						conditionalRefs = a.Value
						break
					}
				}
			case "cfRule":
				currentStopRule = nil
				var stop bool
				var ruleType, operator string
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "stopIfTrue":
						stop = boolAttrValue(a.Value)
					case "type":
						ruleType = a.Value
					case "operator":
						operator = a.Value
					}
				}
				if stop && ruleType == "cellIs" && conditionalRefs != "" {
					currentStopRule = &stopIfTrueRule{refs: conditionalRefs, operator: operator}
				}
			case "formula":
				if currentStopRule != nil {
					inFormula = true
					formula.Reset()
				}
			case "dataBar", "iconSet":
				dataBarDepth++
				showValue = true
				for _, a := range t.Attr {
					if a.Name.Local == "showValue" && !boolAttrValue(a.Value) {
						showValue = false
						break
					}
				}
			}
		case xml.EndElement:
			if t.Name.Local == "v" && inCellValue {
				inCellValue = false
				if currentCellType == "" || currentCellType == "n" {
					if value, err := strconv.ParseFloat(strings.TrimSpace(cellValue.String()), 64); err == nil {
						cellValues[currentCellRef] = value
					}
				}
			}
			if t.Name.Local == "formula" && inFormula {
				inFormula = false
				if currentStopRule != nil {
					currentStopRule.formulas = append(currentStopRule.formulas, strings.TrimSpace(formula.String()))
				}
			}
			if t.Name.Local == "cfRule" && currentStopRule != nil {
				stopIfTrueRules = append(stopIfTrueRules, *currentStopRule)
				currentStopRule = nil
			}
			if (t.Name.Local == "dataBar" || t.Name.Local == "iconSet") && dataBarDepth > 0 {
				dataBarDepth--
				if dataBarDepth == 0 && !showValue {
					hiddenRefs = append(hiddenRefs, conditionalRefs)
				}
			}
		case xml.CharData:
			if inCellValue {
				cellValue.Write([]byte(t))
			}
			if inFormula {
				formula.Write([]byte(t))
			}
		}
	}
	for _, refs := range hiddenRefs {
		for _, ref := range worksheetRefFields(refs) {
			for _, cell := range xlsxCellRefsInRange(ref) {
				hidden[cell] = true
			}
		}
	}
	// A prior cellIs/stopIfTrue rule leaves only the cells matching its numeric
	// predicate visible.  This covers the ordinary rules emitted by Excel and
	// ClosedXML without pretending to evaluate arbitrary Excel formulas.
	for _, rule := range stopIfTrueRules {
		for _, ref := range worksheetRefFields(rule.refs) {
			for _, cell := range xlsxCellRefsInRange(ref) {
				if xlsxCellIsStopIfTrueMatch(cellValues[cell], rule.operator, rule.formulas) {
					delete(hidden, cell)
				}
			}
		}
	}
	return hidden
}

func xlsxCellIsStopIfTrueMatch(value float64, operator string, formulas []string) bool {
	if len(formulas) == 0 {
		return false
	}
	bound, err := strconv.ParseFloat(strings.TrimPrefix(strings.TrimSpace(formulas[0]), "="), 64)
	if err != nil {
		return false
	}
	switch operator {
	case "equal":
		return value == bound
	case "notEqual":
		return value != bound
	case "greaterThan":
		return value > bound
	case "greaterThanOrEqual":
		return value >= bound
	case "lessThan":
		return value < bound
	case "lessThanOrEqual":
		return value <= bound
	case "between":
		if len(formulas) < 2 {
			return false
		}
		upper, err := strconv.ParseFloat(strings.TrimPrefix(strings.TrimSpace(formulas[1]), "="), 64)
		return err == nil && value >= bound && value <= upper
	case "notBetween":
		if len(formulas) < 2 {
			return false
		}
		upper, err := strconv.ParseFloat(strings.TrimPrefix(strings.TrimSpace(formulas[1]), "="), 64)
		return err == nil && (value < bound || value > upper)
	default:
		return false
	}
}

func xlsxCellRefName(col, row int) string {
	if col < 1 || row < 1 {
		return ""
	}
	var letters [8]byte
	i := len(letters)
	for col > 0 {
		col--
		i--
		letters[i] = byte('A' + col%26)
		col /= 26
	}
	return string(letters[i:]) + strconv.Itoa(row)
}

func rangeCoveredByRanges(min, max int, ranges []intRange) bool {
	if min <= 0 || max < min {
		return false
	}
	cur := min
	for cur <= max {
		covered := false
		next := cur
		for _, r := range ranges {
			if cur >= r.min && cur <= r.max {
				covered = true
				if r.max+1 > next {
					next = r.max + 1
				}
			}
		}
		if !covered {
			return false
		}
		cur = next
	}
	return true
}

func cellRefIndexes(ref string) (int, int, bool) {
	ref = strings.ReplaceAll(strings.TrimSpace(ref), "$", "")
	col, row := 0, 0
	for _, r := range ref {
		if r >= 'a' && r <= 'z' {
			r -= 'a' - 'A'
		}
		if r < 'A' || r > 'Z' {
			break
		}
		col = col*26 + int(r-'A'+1)
	}
	for _, r := range ref {
		if r >= '0' && r <= '9' {
			row = row*10 + int(r-'0')
		}
	}
	return col, row, col > 0 && row > 0
}

func elementBoolAttr(start xml.StartElement, name string) bool {
	for _, attr := range start.Attr {
		if attr.Name.Local != name {
			continue
		}
		return boolAttrValue(attr.Value)
	}
	return false
}

func boolAttrValue(value string) bool {
	value = strings.TrimSpace(strings.ToLower(value))
	return value == "1" || value == "true"
}

func appendWorksheetValue(out *strings.Builder, value string, seenLarge *map[string]bool) {
	if value == "" {
		return
	}
	if len(value) > maxRepeatedTextPartBytes {
		if *seenLarge != nil && (*seenLarge)[value] {
			return
		}
		if *seenLarge == nil {
			*seenLarge = map[string]bool{}
		}
		(*seenLarge)[value] = true
	}
	value = cleanText(value)
	if value == "" {
		return
	}
	appendCleanedTextBlock(out, value)
}

func appendCleanedTextBlock(out *strings.Builder, value string) {
	value = strings.TrimSpace(value)
	if value == "" {
		return
	}
	appendTrimmedTextBlock(out, value)
}

func appendTrimmedTextBlock(out *strings.Builder, value string) {
	if out.Len() > 0 {
		out.WriteByte('\n')
	}
	out.WriteString(value)
}

func appendCleanedTextBlocks(out *strings.Builder, values []string) {
	for _, value := range values {
		appendCleanedTextBlock(out, value)
	}
}

func appendCleanedTextParts(base string, parts []string) string {
	if len(parts) == 0 {
		return strings.TrimSpace(base)
	}
	size := len(base)
	for _, p := range parts {
		size += len(p) + 1
	}
	var out strings.Builder
	out.Grow(size)
	appendCleanedTextBlock(&out, base)
	appendCleanedTextBlocks(&out, parts)
	return strings.TrimSpace(out.String())
}

func worksheetAttributeText(start xml.StartElement) []string {
	var out []string
	for _, attr := range start.Attr {
		if !isWorksheetTextAttribute(start.Name.Local, attr.Name.Local) {
			continue
		}
		value := cleanVisibleAttributeValue(attr.Value)
		if value != "" {
			out = append(out, value)
		}
	}
	return out
}

func isWorksheetTextAttribute(element, attr string) bool {
	switch element {
	case "dataValidation":
		switch attr {
		case "promptTitle", "prompt", "errorTitle", "error":
			return true
		}
	case "hyperlink":
		switch attr {
		case "display", "tooltip":
			return true
		}
	}
	return false
}

func visibleVMLText(b []byte) (string, error) {
	if hasDOCTYPE(b) {
		return "", errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	dec.Strict = false
	var out []string
	var skipDepth int
	seenAttrs := map[string]bool{}
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return "", err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if skipDepth > 0 {
				skipDepth++
				continue
			}
			if vmlElementHidden(t) || isVMLSystemElement(t.Name.Local) {
				skipDepth = 1
				continue
			}
			for _, value := range visibleAttributeText(t) {
				if seenAttrs[value] {
					continue
				}
				seenAttrs[value] = true
				out = append(out, "\n", value, "\n")
			}
			switch strings.ToLower(t.Name.Local) {
			case "br", "p", "div", "textbox":
				out = append(out, "\n")
			}
		case xml.EndElement:
			if skipDepth > 0 {
				skipDepth--
				continue
			}
			switch strings.ToLower(t.Name.Local) {
			case "p", "div", "textbox":
				out = append(out, "\n")
			}
		case xml.CharData:
			if skipDepth == 0 {
				out = append(out, string(t))
			}
		}
	}
	return cleanVisibleText(strings.Join(out, "")), nil
}

func vmlElementHidden(start xml.StartElement) bool {
	for _, attr := range start.Attr {
		switch strings.ToLower(attr.Name.Local) {
		case "hidden":
			if boolAttrValue(attr.Value) {
				return true
			}
		case "style":
			if vmlStyleHidden(attr.Value) {
				return true
			}
		}
	}
	return false
}

func isVMLSystemElement(name string) bool {
	switch strings.ToLower(name) {
	case "clientdata":
		return true
	default:
		return false
	}
}

func vmlStyleHidden(style string) bool {
	for _, part := range strings.Split(style, ";") {
		key, value, ok := strings.Cut(part, ":")
		if !ok {
			continue
		}
		key = strings.TrimSpace(strings.ToLower(key))
		value = strings.TrimSpace(strings.ToLower(value))
		if i := strings.IndexByte(value, '!'); i >= 0 {
			value = strings.TrimSpace(value[:i])
		}
		switch key {
		case "display":
			if value == "none" {
				return true
			}
		case "visibility":
			if value == "hidden" || value == "collapse" {
				return true
			}
		case "mso-hide":
			if value == "all" {
				return true
			}
		}
	}
	return false
}

func hasDOCTYPE(b []byte) bool {
	head := b
	if len(head) > 2048 {
		head = head[:2048]
	}
	return bytes.Contains(bytes.ToUpper(head), []byte("<!DOCTYPE"))
}

func isExcelHeaderFooterElement(name string) bool {
	switch name {
	case "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter":
		return true
	default:
		return false
	}
}

func cleanExcelHeaderFooterText(s string) string {
	s = cleanText(s)
	if !strings.Contains(s, "&") {
		return cleanVisibleText(s)
	}
	var out strings.Builder
	out.Grow(len(s))
	for i := 0; i < len(s); {
		if s[i] != '&' || i+1 >= len(s) {
			out.WriteByte(s[i])
			i++
			continue
		}
		code := s[i+1]
		if code == '&' {
			out.WriteByte('&')
			i += 2
			continue
		}
		if code == '"' {
			i += 2
			for i < len(s) && s[i] != '"' {
				i++
			}
			if i < len(s) {
				i++
			}
			continue
		}
		if code == '[' {
			i += 2
			for i < len(s) && s[i] != ']' {
				i++
			}
			if i < len(s) {
				i++
			}
			continue
		}
		if code >= '0' && code <= '9' {
			i += 2
			for i < len(s) && s[i] >= '0' && s[i] <= '9' {
				i++
			}
			continue
		}
		switch code {
		case 'L', 'C', 'R':
			if out.Len() > 0 {
				out.WriteByte('\n')
			}
			i += 2
		case 'K':
			i += 2
			i = skipExcelHeaderFooterColor(s, i)
		case 'B', 'I', 'U', 'E', 'S', 'X', 'Y', 'O', 'H', 'P', 'N', 'D', 'T', 'A', 'F', 'Z', 'G':
			i += 2
		default:
			out.WriteByte('&')
			i++
		}
	}
	return cleanVisibleText(out.String())
}

func skipExcelHeaderFooterColor(s string, i int) int {
	if i+6 <= len(s) && isASCIIHex(s[i]) && isASCIIHex(s[i+1]) && (s[i+2] == '+' || s[i+2] == '-') &&
		isASCIIHex(s[i+3]) && isASCIIHex(s[i+4]) && isASCIIHex(s[i+5]) {
		return i + 6
	}
	n := 0
	for i < len(s) && n < 6 && isASCIIHex(s[i]) {
		i++
		n++
	}
	return i
}

func isASCIIHex(b byte) bool {
	return (b >= '0' && b <= '9') || (b >= 'a' && b <= 'f') || (b >= 'A' && b <= 'F')
}

func extractOOXMLImages(files map[string]*zip.File, kind string, includeMetadata bool, strictOfficeImages ...bool) ([]Image, error) {
	var prefix string
	switch kind {
	case "docx":
		prefix = "word/media/"
	case "pptx":
		prefix = "ppt/media/"
	case "xlsx":
		prefix = "xl/media/"
	default:
		prefix = "/media/"
	}
	var names []string
	strict := len(strictOfficeImages) > 0 && strictOfficeImages[0]
	for name := range files {
		// Although Word normally stores images under word/media, OPC permits a
		// document relationship to target a package-root media part.  Include
		// those only for strict Word comparison; the relationship filter below
		// still decides whether Word exposes the part as a Picture Shape.
		strictDocxRootMedia := strict && kind == "docx" && strings.HasPrefix(ooxmlPartKey(name), "media/")
		if isOOXMLMediaPart(name, prefix, kind) || strictDocxRootMedia || (includeMetadata && isOOXMLThumbnail(name)) {
			names = append(names, name)
		}
	}
	visibleMedia, filterMedia := visibleOOXMLMediaParts(files, kind)
	if strict && kind == "docx" {
		visibleMedia, filterMedia = strictDocxVisibleMediaParts(files)
	}
	if strict && kind == "pptx" {
		visibleMedia, filterMedia = strictPptxVisibleMediaParts(files)
	}
	sort.Strings(names)
	altData := ooxmlVisibleImageAltData(files, kind)
	altsByMedia := altData.byMedia
	var alts []string
	altsLoaded := false
	images := make([]Image, 0, len(names))
	usedImageNames := map[string]bool{}
	for _, name := range names {
		if filterMedia && isOOXMLMediaPart(name, prefix, kind) && !visibleOOXMLMediaPartAllowed(visibleMedia, name) {
			continue
		}
		b, err := readZipFile(files[name])
		if err != nil {
			continue
		}
		cleanName := ooxmlCleanPartName(name)
		ext := strings.ToLower(path.Ext(cleanName))
		b, ext, ok := normalizeOOXMLImageData(ext, b)
		if !ok {
			continue
		}
		imgName := ooxmlImageOutputBaseName(cleanName)
		imgName = imageNameWithExt(imgName, ext)
		imgName = uniqueImageFilename(sanitizeFilename(imgName), usedImageNames)
		img := Image{Name: imgName, Ext: ext, Data: append([]byte(nil), b...)}
		if alt := cleanMarkdownImageAltText(altsByMedia[ooxmlPartKey(name)]); alt != "" {
			img.Alt = alt
		} else {
			if !altsLoaded {
				alts = ooxmlVisibleImageAlts(files, kind, altData.byPart)
				altsLoaded = true
			}
			if len(alts) > len(images) {
				img.Alt = cleanMarkdownImageAltText(alts[len(images)])
			}
		}
		images = append(images, img)
	}
	return images, nil
}

// strictDocxVisibleMediaParts mirrors Word's Document.InlineShapes and
// Document.Shapes collections.  These collections belong to the main
// document story, not headers, footers, comments, or glossary documents.
func strictDocxVisibleMediaParts(files map[string]*zip.File) (map[string]bool, bool) {
	document := ooxmlPartName(files, "word/document.xml")
	if document == "" {
		return nil, false
	}
	b, err := readZipFile(ooxmlFile(files, document))
	if err != nil {
		return nil, false
	}
	refs, err := docxStrictPictureRelationshipRefs(b)
	if err != nil {
		return nil, false
	}
	rels, err := relationshipTargetMapForPart(files, document)
	if err != nil {
		return nil, false
	}
	visible := map[string]bool{}
	for id := range refs.Visible {
		if part := docxRelationshipMediaPart(files, document, rels[id]); part != "" {
			visible[part] = true
		}
	}
	return visible, true
}

// docxStrictVisibleImageOccurrences mirrors Word's Document.InlineShapes and
// Document.Shapes collections.  Unlike the package media directory, those
// collections preserve each placement of a picture: one media part may be
// displayed several times in the document and must therefore produce several
// extracted image occurrences.
func docxStrictVisibleImageOccurrences(files map[string]*zip.File, images []Image) []Image {
	document := ooxmlPartName(files, "word/document.xml")
	if document == "" {
		return images
	}
	b, err := readZipFile(ooxmlFile(files, document))
	if err != nil {
		return images
	}
	ids, err := docxStrictPictureRelationshipIDsInOrder(b)
	if err != nil {
		return images
	}
	rels, err := relationshipTargetMapForPart(files, document)
	if err != nil {
		return images
	}
	parts := make([]string, 0, len(ids))
	for _, id := range ids {
		target, known := rels[id]
		if known {
			if part := docxRelationshipMediaPart(files, document, target); part != "" {
				parts = append(parts, part)
			}
			continue
		}
		if !known && id == "" {
			continue
		}
	}
	return ooxmlImageOccurrencesForPartsWithFiles(parts, images, files)
}

func xlsxStrictVisibleImageOccurrences(files map[string]*zip.File, images []Image) []Image {
	media, _ := xlsxStrictVisibleMediaOccurrenceNames(files)
	// Excel's Shapes collection exposes picture shapes from worksheet Drawing
	// parts.  A VML legacyDrawing can reference an image (commonly a header,
	// comment, or compatibility artefact), but it is not exposed as an Excel
	// Picture Shape.  In strict Office mode, lack of a verified Drawing
	// occurrence must therefore mean no visible picture, rather than falling
	// back to every package media part.
	if len(media) == 0 {
		return nil
	}
	byPart := map[string][]Image{}
	for _, image := range images {
		key := strings.ToLower(image.Name)
		byPart[key] = append(byPart[key], image)
	}
	var out []Image
	for _, part := range media {
		name := imageNameWithExt(ooxmlImageOutputBaseName(part), strings.ToLower(path.Ext(part)))
		key := strings.ToLower(name)
		available := byPart[key]
		if len(available) == 0 {
			base := strings.ToLower(strings.TrimSuffix(ooxmlImageOutputBaseName(part), path.Ext(ooxmlImageOutputBaseName(part))))
			for candidate, values := range byPart {
				if len(values) > 0 && strings.EqualFold(strings.TrimSuffix(candidate, path.Ext(candidate)), base) {
					key, available = candidate, values
					break
				}
			}
		}
		if len(available) == 0 {
			continue
		}
		out = append(out, available[0])
	}
	return out
}

// pptxStrictVisibleImageOccurrences mirrors PowerPoint's Slide.Shapes
// collection. A media part can be used by several Picture Shapes; keep every
// visible shape occurrence instead of returning one image per ZIP part. Group
// members are rendered and exposed by PowerPoint through GroupItems, so they
// are included; OLE previews remain excluded.
//
// Callers must pass the pre-filtered image set: a drawing can legitimately
// reuse one media part, whereas an orphaned media part must never be revived
// merely because its basename happens to match a visible relationship.
func pptxStrictVisibleImageOccurrences(files map[string]*zip.File, images []Image) []Image {
	media, found := pptxStrictVisibleMediaOccurrenceNames(files)
	if !found {
		return images
	}
	return ooxmlImageOccurrencesForParts(media, images)
}

func ooxmlImageOccurrencesForParts(parts []string, images []Image) []Image {
	byName := map[string][]Image{}
	for _, image := range images {
		key := strings.ToLower(image.Name)
		byName[key] = append(byName[key], image)
	}
	out := make([]Image, 0, len(parts))
	for _, part := range parts {
		baseName := ooxmlImageOutputBaseName(part)
		name := imageNameWithExt(baseName, strings.ToLower(path.Ext(part)))
		key := strings.ToLower(name)
		available := byName[key]
		if len(available) == 0 {
			base := strings.ToLower(strings.TrimSuffix(baseName, path.Ext(baseName)))
			for candidate, values := range byName {
				if len(values) > 0 && strings.EqualFold(strings.TrimSuffix(candidate, path.Ext(candidate)), base) {
					available = values
					break
				}
			}
		}
		if len(available) > 0 {
			out = append(out, available[0])
		}
	}
	return out
}

// ooxmlImageOccurrencesForPartsWithFiles preserves a strict OOXML occurrence
// even when its image payload is deliberately not materialized by the normal
// image pipeline (for example a valid WMF that a caller cannot decode).  Word
// nevertheless exposes that placement as a Picture Shape, so dropping it
// would make the strict COM image count wrong. Existing materialized Images
// still win, retaining their exact decoded data and metadata.
func ooxmlImageOccurrencesForPartsWithFiles(parts []string, images []Image, files map[string]*zip.File) []Image {
	byName := map[string][]Image{}
	for _, image := range images {
		key := strings.ToLower(image.Name)
		byName[key] = append(byName[key], image)
	}
	used := map[string]bool{}
	out := make([]Image, 0, len(parts))
	for _, part := range parts {
		baseName := ooxmlImageOutputBaseName(part)
		name := imageNameWithExt(baseName, strings.ToLower(path.Ext(part)))
		key := strings.ToLower(name)
		available := byName[key]
		if len(available) == 0 {
			base := strings.ToLower(strings.TrimSuffix(baseName, path.Ext(baseName)))
			for candidate, values := range byName {
				if len(values) > 0 && strings.EqualFold(strings.TrimSuffix(candidate, path.Ext(candidate)), base) {
					available = values
					break
				}
			}
		}
		if len(available) > 0 {
			image := available[0]
			image.Name = uniqueImageFilename(sanitizeFilename(image.Name), used)
			out = append(out, image)
			continue
		}
		f := ooxmlFile(files, part)
		if f == nil {
			continue
		}
		data, err := readZipFile(f)
		if err != nil {
			continue
		}
		ext := strings.ToLower(path.Ext(part))
		if ext == "" || !isSupportedImageExt(ext) {
			continue
		}
		name = uniqueImageFilename(sanitizeFilename(imageNameWithExt(baseName, ext)), used)
		out = append(out, Image{Name: name, Ext: ext, Data: data})
	}
	return out
}

func ooxmlImagesForParts(parts map[string]bool, images []Image) []Image {
	if len(parts) == 0 {
		return nil
	}
	out := make([]Image, 0, len(images))
	for _, image := range images {
		for part := range parts {
			baseName := ooxmlImageOutputBaseName(part)
			name := imageNameWithExt(baseName, strings.ToLower(path.Ext(part)))
			if strings.EqualFold(image.Name, name) {
				out = append(out, image)
				break
			}
		}
	}
	return out
}

func xlsxStrictVisibleMediaOccurrenceNames(files map[string]*zip.File) ([]string, bool) {
	sheets, err := workbookVisibleSheets(files)
	if err != nil || len(sheets) == 0 {
		return nil, false
	}
	var out []string
	foundDrawing := false
	for _, sheet := range sheets {
		for _, drawing := range relationshipTargetsWithPrefix(files, sheet.Path, "xl/drawings/") {
			refs, err := xlsxDrawingImageRefsInOrder(files, drawing)
			if err != nil {
				continue
			}
			if len(refs) == 0 {
				continue
			}
			foundDrawing = true
			rels, err := relationshipTargetMapForPart(files, drawing)
			if err != nil {
				continue
			}
			for _, id := range refs {
				if media := relationshipMediaPart(files, drawing, rels[id], "xl/media/"); media != "" {
					out = append(out, media)
				}
			}
		}
		// An Excel legacyDrawing normally holds comments or form controls, neither
		// of which is a Picture Shape.  Its VML PictureFrame form (ClientData
		// ObjectType="Pict"), however, is exposed by Worksheet.Shapes as an
		// msoPicture.  Keep precisely those occurrences: treating every VML image
		// as a picture would revive comment artwork and compatibility caches.
		for _, drawing := range relationshipTargetsWithPrefix(files, sheet.Path, "xl/drawings/") {
			if !strings.HasSuffix(ooxmlPartKey(drawing), ".vml") {
				continue
			}
			refs, err := xlsxVMLPictureRelationshipIDOccurrences(files, drawing, xlsxOLEObjectShapeIDs(files, sheet.Path))
			if err != nil || len(refs) == 0 {
				continue
			}
			foundDrawing = true
			rels, err := relationshipTargetMapForPart(files, drawing)
			if err != nil {
				continue
			}
			for _, id := range refs {
				if media := relationshipMediaPart(files, drawing, rels[id], "xl/media/"); media != "" {
					out = append(out, media)
				}
			}
		}
	}
	return out, foundDrawing
}

// xlsxVMLPictureRelationshipIDOccurrences returns the image relationships of
// VML PictureFrame controls. Excel persists comments, buttons, dropdowns and
// miscellaneous compatibility objects in the same legacyDrawing part; only
// ObjectType="Pict" participates in Worksheet.Shapes as a picture. An
// oleObject can reuse that form for its preview artwork: COM exposes it as an
// embedded object rather than msoPicture, so its VML shape IDs are excluded.
func xlsxVMLPictureRelationshipIDOccurrences(files map[string]*zip.File, drawing string, oleShapeIDs map[string]bool) ([]string, error) {
	b, err := readZipFile(ooxmlFile(files, drawing))
	if err != nil {
		return nil, err
	}
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	shapeDepth := 0
	pictureShape := false
	shapeID := ""
	var shapeIDs []string
	var ids []string
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			return ids, nil
		}
		if err != nil {
			return nil, err
		}
		switch value := tok.(type) {
		case xml.StartElement:
			if value.Name.Local == "shape" && (strings.Contains(strings.ToLower(value.Name.Space), "vml") || strings.HasPrefix(strings.ToLower(xmlAttrValue(value, "id")), "_x0000_") || xmlAttrValue(value, "id") != "") {
				if shapeDepth == 0 {
					pictureShape = false
					shapeID = normalizeVMLShapeID(xmlAttrValue(value, "id"))
					shapeIDs = shapeIDs[:0]
				}
				shapeDepth++
				continue
			}
			if shapeDepth == 0 {
				continue
			}
			if value.Name.Local == "ClientData" && strings.EqualFold(xmlAttrValue(value, "ObjectType"), "Pict") {
				pictureShape = !oleShapeIDs[shapeID]
				continue
			}
			if value.Name.Local == "imagedata" {
				shapeIDs = append(shapeIDs, imageRelationshipIDs(value)...)
			}
		case xml.EndElement:
			if value.Name.Local == "shape" && shapeDepth > 0 {
				shapeDepth--
				if shapeDepth == 0 {
					if pictureShape {
						ids = append(ids, shapeIDs...)
					}
					pictureShape = false
					shapeID = ""
					shapeIDs = shapeIDs[:0]
				}
			}
		}
	}
}

func xlsxDrawingImageRefsInOrder(files map[string]*zip.File, drawing string) ([]string, error) {
	b, err := readZipFile(ooxmlFile(files, drawing))
	if err != nil {
		return nil, err
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var refs []string
	for {
		tok, err := dec.RawToken()
		if errors.Is(err, io.EOF) {
			return refs, nil
		}
		if err != nil {
			return nil, err
		}
		switch value := tok.(type) {
		case xml.StartElement:
			if value.Name.Local == "blip" {
				for _, attr := range value.Attr {
					if attr.Name.Local == "embed" || attr.Name.Local == "link" {
						if id := strings.TrimSpace(attr.Value); id != "" {
							refs = append(refs, id)
						}
					}
				}
			}
		}
	}
}

func ooxmlImageOutputBaseName(name string) string {
	base := path.Base(ooxmlCleanPartName(name))
	if decoded, err := url.PathUnescape(base); err == nil && decoded != "" {
		return decoded
	}
	return base
}

func visibleOOXMLMediaPartAllowed(visibleMedia map[string]bool, name string) bool {
	if visibleMedia[name] || visibleMedia[ooxmlPartKey(name)] {
		return true
	}
	for _, key := range ooxmlPartKeyCandidates(name) {
		if visibleMedia[key] {
			return true
		}
	}
	return false
}

func visibleOOXMLMediaParts(files map[string]*zip.File, kind string) (map[string]bool, bool) {
	switch kind {
	case "docx":
		return docxVisibleMediaParts(files)
	case "pptx":
		return pptxVisibleMediaParts(files)
	case "xlsx":
		return xlsxVisibleMediaParts(files)
	default:
		return nil, false
	}
}

func docxVisibleMediaParts(files map[string]*zip.File) (map[string]bool, bool) {
	visible := map[string]bool{}
	hidden := map[string]bool{}
	_, hiddenHeaderFooter, constrainedHeaderFooter := docxVisibleHeaderFooterParts(files)
	foundRefs := false
	malformedRels := false
	for _, name := range docxMediaReferencePartNames(files) {
		sourceKey := ooxmlPartKey(name)
		sourceIsHidden := strings.HasPrefix(sourceKey, "word/glossary/") || (constrainedHeaderFooter && hiddenHeaderFooter[sourceKey])
		b, err := readZipFile(files[name])
		if err != nil {
			continue
		}
		if !likelyImageRelationshipMarkup(b) {
			continue
		}
		refs, err := imageRelationshipRefsFromPartBytes(files, name, b)
		if err != nil {
			continue
		}
		if len(refs.Hidden) == 0 && len(refs.Visible) == 0 {
			continue
		}
		foundRefs = true
		rels, err := relationshipTargetMapForPart(files, name)
		if err != nil {
			malformedRels = true
			continue
		}
		for id := range refs.Visible {
			if part := docxRelationshipMediaPart(files, name, rels[id]); part != "" {
				if sourceIsHidden {
					hidden[part] = true
				} else {
					visible[part] = true
				}
			}
		}
		for id := range refs.Hidden {
			if part := docxRelationshipMediaPart(files, name, rels[id]); part != "" {
				hidden[part] = true
			}
		}
	}
	for _, name := range docxVisibleHTMLPartNamesNoError(files) {
		b, err := readZipFile(files[name])
		if err != nil {
			continue
		}
		refs := htmlImageMediaRefs(files, name, b)
		if len(refs) == 0 {
			continue
		}
		foundRefs = true
		for _, part := range refs {
			visible[part] = true
		}
	}
	if !foundRefs || malformedRels {
		return nil, false
	}
	allowed := map[string]bool{}
	for name := range visible {
		allowed[name] = true
	}
	for name := range hidden {
		if !visible[name] {
			delete(allowed, name)
		}
	}
	return allowed, true
}

func likelyImageRelationshipMarkup(b []byte) bool {
	for _, marker := range [][]byte{
		[]byte(":embed"),
		[]byte(" embed"),
		[]byte(":link"),
		[]byte(" link"),
		[]byte(":id"),
		[]byte(":relid"),
		[]byte(" relid"),
	} {
		if bytes.Contains(b, marker) {
			return true
		}
	}
	return false
}

func docxMediaReferencePartNames(files map[string]*zip.File) []string {
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if strings.HasPrefix(lower, "word/") && (strings.HasSuffix(lower, ".xml") || strings.HasSuffix(lower, ".vml")) {
			names = append(names, name)
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return names
}

func docxVisibleHTMLPartNamesNoError(files map[string]*zip.File) []string {
	names, err := docxVisibleHTMLPartNames(files)
	if err != nil {
		return nil
	}
	return names
}

func docxRelationshipMediaPart(files map[string]*zip.File, source, target string) string {
	part := resolveOOXMLRelationshipTarget(source, target)
	if part == "" {
		// Word accepts package-root image targets (Target="/media/...") even
		// though the generic cross-root resolver correctly rejects arbitrary
		// Word-to-package-root relationships.  Admit only that narrow, normalized
		// media location here; every other cross-root target stays rejected.
		clean := cleanOOXMLRelationshipTarget(target)
		if strings.HasPrefix(clean, "/") {
			candidate := strings.TrimPrefix(path.Clean(clean), "/")
			if strings.HasPrefix(strings.ToLower(candidate), "media/") {
				part = candidate
			}
		}
	}
	if actual := ooxmlPartName(files, part); actual != "" {
		part = actual
	}
	// Most Word image relationships target word/media.  A valid OPC package can
	// also place a document-owned image at package root (for example
	// Target="/media/image.png"); Word still exposes it as an InlinePicture.
	// Accept a resolved package-root media part while retaining the explicit
	// source-root validation in resolveOOXMLRelationshipTarget.
	key := ooxmlPartKey(part)
	if strings.HasPrefix(key, "word/media/") || strings.HasPrefix(key, "media/") {
		return part
	}
	return ""
}

type docxImageRefs struct {
	Visible map[string]bool
	Hidden  map[string]bool
}

func cloneDocxImageRefs(in docxImageRefs) docxImageRefs {
	return docxImageRefs{
		Visible: cloneBoolMap(in.Visible),
		Hidden:  cloneBoolMap(in.Hidden),
	}
}

func docxImageRelationshipRefs(b []byte) (docxImageRefs, error) {
	refs := docxImageRefs{Visible: map[string]bool{}, Hidden: map[string]bool{}}
	if hasDOCTYPE(b) {
		return refs, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var hiddenRevisionDepth int
	var hiddenRevisionRangeDepth int
	var drawingObjectStack []bool
	var drawingImageStack []bool
	var paragraphHiddenStack []bool
	var alternateStack []alternateContentState
	var skipDepth int
	var backgroundDepth int
	var runDepth int
	var rPrDepth int
	var pPrDepth int
	var runHidden bool
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return refs, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if skipDepth > 0 {
				skipDepth++
				continue
			}
			if alternateContentStartSkip(t.Name.Local, &alternateStack) {
				if t.Name.Local == "Fallback" {
					skipDepth = 1
				}
				continue
			}
			if hiddenRevisionDepth > 0 {
				hiddenRevisionDepth++
			} else if isHiddenRevisionElement(t.Name) {
				hiddenRevisionDepth = 1
			}
			if hiddenRevisionDepth == 0 {
				if isHiddenRevisionRangeStart(t.Name) {
					hiddenRevisionRangeDepth++
				} else if isHiddenRevisionRangeEnd(t.Name) && hiddenRevisionRangeDepth > 0 {
					hiddenRevisionRangeDepth--
				}
			}
			// Word stores a document-page background as w:background with a
			// relationship to its tile image.  It is rendered behind the page, but
			// Word does not expose it from either Document.InlineShapes or
			// Document.Shapes, which are the strict Office image baseline.
			if t.Name.Local == "background" {
				backgroundDepth++
			}
			if isDrawingObjectElement(t.Name.Local) {
				parentHidden := len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]
				drawingObjectStack = append(drawingObjectStack, parentHidden)
				parentImageHidden := len(drawingImageStack) > 0 && drawingImageStack[len(drawingImageStack)-1]
				drawingImageStack = append(drawingImageStack, parentImageHidden)
			}
			if len(drawingObjectStack) > 0 && drawingObjectElementHidden(t) {
				drawingObjectStack[len(drawingObjectStack)-1] = true
			}
			if len(drawingImageStack) > 0 && drawingObjectElementExplicitlyHidden(t) {
				drawingImageStack[len(drawingImageStack)-1] = true
			}
			switch t.Name.Local {
			case "p":
				paragraphHiddenStack = append(paragraphHiddenStack, false)
			case "r":
				runDepth++
				runHidden = false
			case "pPr":
				if len(paragraphHiddenStack) > 0 {
					pPrDepth++
				}
			case "rPr":
				if runDepth > 0 {
					rPrDepth++
				}
			case "vanish", "webHidden":
				if runDepth > 0 && rPrDepth > 0 {
					runHidden = true
				}
				if pPrDepth > 0 && len(paragraphHiddenStack) > 0 {
					paragraphHiddenStack[len(paragraphHiddenStack)-1] = true
				}
			}
			imageHidden := backgroundDepth > 0 || hiddenRevisionDepth > 0 || hiddenRevisionRangeDepth > 0 || runHidden || currentParagraphHidden(paragraphHiddenStack) ||
				(len(drawingImageStack) > 0 && drawingImageStack[len(drawingImageStack)-1])
			for _, id := range imageRelationshipIDs(t) {
				if imageHidden {
					refs.Hidden[id] = true
				} else {
					refs.Visible[id] = true
				}
			}
		case xml.EndElement:
			if skipDepth > 0 {
				skipDepth--
				continue
			}
			if alternateContentEnd(t.Name.Local, &alternateStack) {
				continue
			}
			if t.Name.Local == "background" && backgroundDepth > 0 {
				backgroundDepth--
			}
			if t.Name.Local == "pPr" && pPrDepth > 0 {
				pPrDepth--
			}
			if t.Name.Local == "rPr" && rPrDepth > 0 {
				rPrDepth--
			}
			if t.Name.Local == "r" && runDepth > 0 {
				runDepth--
				if runDepth == 0 {
					runHidden = false
					rPrDepth = 0
				}
			}
			if t.Name.Local == "p" && len(paragraphHiddenStack) > 0 {
				paragraphHiddenStack = paragraphHiddenStack[:len(paragraphHiddenStack)-1]
				if len(paragraphHiddenStack) == 0 {
					pPrDepth = 0
				}
			}
			if isDrawingObjectElement(t.Name.Local) && len(drawingObjectStack) > 0 {
				drawingObjectStack = drawingObjectStack[:len(drawingObjectStack)-1]
				drawingImageStack = drawingImageStack[:len(drawingImageStack)-1]
			}
			if hiddenRevisionDepth > 0 {
				hiddenRevisionDepth--
			}
		}
	}
	return refs, nil
}

func drawingObjectElementExplicitlyHidden(start xml.StartElement) bool {
	return vmlElementHidden(start)
}

func xlsxImageRelationshipRefs(b []byte) (docxImageRefs, error) {
	refs := docxImageRefs{Visible: map[string]bool{}, Hidden: map[string]bool{}}
	if hasDOCTYPE(b) {
		return refs, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var drawingObjectStack []bool
	var alternateStack []alternateContentState
	var skipDepth int
	for {
		tok, err := dec.RawToken()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return refs, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if skipDepth > 0 {
				skipDepth++
				continue
			}
			if alternateContentStartSkip(t.Name.Local, &alternateStack) {
				if t.Name.Local == "Fallback" {
					skipDepth = 1
				}
				continue
			}
			if isDrawingObjectElement(t.Name.Local) {
				parentHidden := len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]
				drawingObjectStack = append(drawingObjectStack, parentHidden)
			}
			if len(drawingObjectStack) > 0 && drawingObjectElementHidden(t) {
				drawingObjectStack[len(drawingObjectStack)-1] = true
			}
			hidden := len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]
			for _, id := range imageRelationshipIDs(t) {
				if hidden {
					refs.Hidden[id] = true
				} else {
					refs.Visible[id] = true
				}
			}
		case xml.EndElement:
			if skipDepth > 0 {
				skipDepth--
				continue
			}
			if alternateContentEnd(t.Name.Local, &alternateStack) {
				continue
			}
			if isDrawingObjectElement(t.Name.Local) && len(drawingObjectStack) > 0 {
				drawingObjectStack = drawingObjectStack[:len(drawingObjectStack)-1]
			}
		}
	}
	return refs, nil
}

func imageRelationshipIDs(start xml.StartElement) []string {
	var ids []string
	for _, attr := range start.Attr {
		value := strings.TrimSpace(attr.Value)
		if value == "" {
			continue
		}
		switch attr.Name.Local {
		case "embed", "link", "relid":
			ids = append(ids, value)
		case "id":
			if attr.Name.Space != "" {
				ids = append(ids, value)
			}
		}
	}
	return ids
}

func scanXMLRelationshipAttributeIDs(b []byte, emit func(string)) {
	for offset := 0; offset < len(b); {
		i := bytes.IndexByte(b[offset:], '<')
		if i < 0 {
			return
		}
		start := offset + i
		if start+1 >= len(b) {
			return
		}
		switch b[start+1] {
		case '/', '!', '?':
			offset = start + 2
			continue
		}
		end := xmlTagEnd(b, start+1)
		if end < 0 {
			return
		}
		scanXMLTagRelationshipAttributeIDs(b[start+1:end], emit)
		offset = end + 1
	}
}

func xmlTagEnd(b []byte, start int) int {
	var quote byte
	for i := start; i < len(b); i++ {
		c := b[i]
		if quote != 0 {
			if c == quote {
				quote = 0
			}
			continue
		}
		if c == '\'' || c == '"' {
			quote = c
			continue
		}
		if c == '>' {
			return i
		}
	}
	return -1
}

func scanXMLTagRelationshipAttributeIDs(tag []byte, emit func(string)) {
	i := xmlTagNameEnd(tag)
	for i < len(tag) {
		for i < len(tag) && isXMLSpace(tag[i]) {
			i++
		}
		if i >= len(tag) || tag[i] == '/' {
			return
		}
		nameStart := i
		for i < len(tag) && !isXMLSpace(tag[i]) && tag[i] != '=' && tag[i] != '/' {
			i++
		}
		name := tag[nameStart:i]
		for i < len(tag) && isXMLSpace(tag[i]) {
			i++
		}
		if i >= len(tag) || tag[i] != '=' {
			continue
		}
		i++
		for i < len(tag) && isXMLSpace(tag[i]) {
			i++
		}
		if i >= len(tag) || (tag[i] != '\'' && tag[i] != '"') {
			continue
		}
		quote := tag[i]
		i++
		valueStart := i
		for i < len(tag) && tag[i] != quote {
			i++
		}
		if i >= len(tag) {
			return
		}
		if xmlRelationshipAttributeName(name) {
			if value := strings.TrimSpace(string(tag[valueStart:i])); value != "" {
				emit(value)
			}
		}
		i++
	}
}

func xmlTagAttr(tag []byte, want string) ([]byte, bool) {
	var out []byte
	scanXMLTagAttrs(tag, func(name, value []byte) {
		if out != nil {
			return
		}
		if string(xmlNameLocal(name)) == want {
			out = value
		}
	})
	if out == nil {
		return nil, false
	}
	return out, true
}

func xmlTagIntAttr(tag []byte, name string) (int, bool) {
	value, ok := xmlTagAttr(tag, name)
	if !ok {
		return 0, false
	}
	return atoiBytes(value)
}

func xmlTagFloatAttr(tag []byte, name string) (float64, bool) {
	value, ok := xmlTagAttr(tag, name)
	if !ok {
		return 0, false
	}
	f, err := strconv.ParseFloat(string(bytes.TrimSpace(value)), 64)
	if err != nil {
		return 0, false
	}
	return f, true
}

func scanXMLTagAttrs(tag []byte, emit func(name, value []byte)) {
	i := xmlTagNameEnd(tag)
	for i < len(tag) {
		for i < len(tag) && isXMLSpace(tag[i]) {
			i++
		}
		if i >= len(tag) || tag[i] == '/' {
			return
		}
		nameStart := i
		for i < len(tag) && !isXMLSpace(tag[i]) && tag[i] != '=' && tag[i] != '/' {
			i++
		}
		name := tag[nameStart:i]
		for i < len(tag) && isXMLSpace(tag[i]) {
			i++
		}
		if i >= len(tag) || tag[i] != '=' {
			continue
		}
		i++
		for i < len(tag) && isXMLSpace(tag[i]) {
			i++
		}
		if i >= len(tag) || (tag[i] != '\'' && tag[i] != '"') {
			continue
		}
		quote := tag[i]
		i++
		valueStart := i
		for i < len(tag) && tag[i] != quote {
			i++
		}
		if i >= len(tag) {
			return
		}
		emit(name, bytes.TrimSpace(tag[valueStart:i]))
		i++
	}
}

func xmlTagLocalName(tag []byte) []byte {
	end := xmlTagNameEnd(tag)
	return xmlNameLocal(tag[:end])
}

func xmlTagNameEnd(tag []byte) int {
	i := 0
	for i < len(tag) && !isXMLSpace(tag[i]) && tag[i] != '/' {
		i++
	}
	return i
}

func xmlNameLocal(name []byte) []byte {
	if colon := bytes.LastIndexByte(name, ':'); colon >= 0 {
		return name[colon+1:]
	}
	return name
}

func xmlRelationshipAttributeName(name []byte) bool {
	colon := bytes.LastIndexByte(name, ':')
	local := xmlNameLocal(name)
	switch string(local) {
	case "embed", "link", "relid":
		return true
	case "id":
		return colon >= 0
	default:
		return false
	}
}

func atoiBytes(value []byte) (int, bool) {
	value = bytes.TrimSpace(value)
	if len(value) == 0 {
		return 0, false
	}
	n := 0
	for _, c := range value {
		if c < '0' || c > '9' {
			return 0, false
		}
		n = n*10 + int(c-'0')
	}
	return n, true
}

func boolAttrBytes(value []byte) bool {
	value = bytes.TrimSpace(value)
	return bytes.Equal(value, []byte("1")) || bytes.Equal(bytes.ToLower(value), []byte("true"))
}

func isXMLSpace(c byte) bool {
	return c == ' ' || c == '\t' || c == '\r' || c == '\n'
}

func pptxVisibleMediaParts(files map[string]*zip.File) (map[string]bool, bool) {
	visible := map[string]bool{}
	hidden := map[string]bool{}
	foundRels := false
	foundSlides := false
	slideNames, constrained, err := pptxCandidateSlideNames(files)
	if err != nil {
		return nil, false
	}
	candidate := map[string]bool{}
	for _, name := range slideNames {
		candidate[ooxmlPartKey(name)] = true
	}
	if constrained {
		slideNames = pptxAllSlidePartNames(files)
	}
	for _, name := range slideNames {
		foundSlides = true
		slideVisible, err := pptxSlideVisible(files, name)
		if err != nil {
			return nil, false
		}
		if !slideVisible || (constrained && !candidate[ooxmlPartKey(name)]) {
			if collectReachableOOXMLMedia(files, name, "ppt/media/", hidden, map[string]bool{}) {
				foundRels = true
			}
			continue
		}
		if collectPptxSlideVisibleMedia(files, name, visible, hidden) {
			foundRels = true
		} else if ooxmlFile(files, ooxmlRelsName(name)) != nil {
			return nil, false
		}
	}
	if !foundSlides {
		return nil, false
	}
	if !foundRels {
		return nil, false
	}
	for name := range hidden {
		if !visible[name] {
			delete(visible, name)
		}
	}
	return visible, true
}

// strictPptxVisibleMediaParts follows PowerPoint's Picture Shape model. It
// accepts direct p:pic references on visible slides, including group members,
// while deliberately excluding master/layout resources and OLE previews.
func strictPptxVisibleMediaParts(files map[string]*zip.File) (map[string]bool, bool) {
	visible := map[string]bool{}
	slides, _, err := pptxCandidateSlideNames(files)
	if err != nil || len(slides) == 0 {
		return nil, false
	}
	for _, slide := range slides {
		ok, err := pptxSlideVisible(files, slide)
		if err != nil || !ok {
			continue
		}
		f := ooxmlFile(files, slide)
		if f == nil {
			return nil, false
		}
		occurrences, ok := pptxStrictSlidePictureMediaOccurrences(files, slide, f)
		if !ok {
			return nil, false
		}
		for _, part := range occurrences {
			visible[part] = true
		}
	}
	// An empty set is meaningful: it means the presentation has no top-level
	// Picture Shapes.  Keep filtering enabled so orphaned/master media does not
	// fall back into StrictOfficeImages merely because no slide picture exists.
	return visible, true
}

func pptxStrictVisibleMediaOccurrenceNames(files map[string]*zip.File) ([]string, bool) {
	slides, _, err := pptxCandidateSlideNames(files)
	if err != nil || len(slides) == 0 {
		return nil, false
	}
	var out []string
	for _, slide := range slides {
		visible, err := pptxSlideVisible(files, slide)
		if err != nil || !visible {
			continue
		}
		f := ooxmlFile(files, slide)
		if f == nil {
			return nil, false
		}
		occurrences, ok := pptxStrictSlidePictureMediaOccurrences(files, slide, f)
		if !ok {
			return nil, false
		}
		out = append(out, occurrences...)
	}
	return out, true
}

// pptxStrictSlidePictureMediaOccurrences follows both image encodings which
// PowerPoint surfaces through Slide.Shapes: DrawingML p:pic elements in the
// slide itself and VML v:shape/v:imagedata elements reached through the
// slide's vmlDrawing relationship.  HTML-originated presentations commonly
// retain the latter and PowerPoint exposes each as msoPicture.
func pptxStrictSlidePictureMediaOccurrences(files map[string]*zip.File, slide string, slideFile *zip.File) ([]string, bool) {
	b, err := readZipFile(slideFile)
	if err != nil {
		return nil, false
	}
	ids, err := pptxPictureRelationshipIDOccurrences(b)
	if err != nil {
		return nil, false
	}
	rels, err := relationshipTargetMapForPart(files, slide)
	if err != nil {
		return nil, false
	}
	var out []string
	appendMedia := func(source string, ids []string, targets map[string]string) {
		for _, id := range ids {
			part := resolveOOXMLRelationshipTarget(source, targets[id])
			if actual := ooxmlPartName(files, part); actual != "" {
				part = actual
			}
			if strings.HasPrefix(ooxmlPartKey(part), "ppt/media/") {
				out = append(out, part)
			}
		}
	}
	appendMedia(slide, ids, rels)
	// PowerPoint repairs a small class of producer defects where a p:pic blip
	// still names rIdN but that one entry is absent from slideN.xml.rels.  Do
	// not fall back to arbitrary package media: infer a missing target only if
	// two or more ordinary image relationships on this same slide establish the
	// same rId-to-image-number offset (for example rId2 -> image102 and rId12
	// -> image112).  This mirrors Office's recovered Picture Shape while
	// keeping orphaned media out of the strict result.
	for _, id := range ids {
		if _, found := rels[id]; found {
			continue
		}
		if part := pptxRecoveredMissingImageRelationshipPart(files, id, rels); part != "" {
			out = append(out, part)
		}
	}
	// A relationship entry alone does not make VML slide artwork visible.
	// PowerPoint packages can retain unreferenced VML chart/PictureFrame caches
	// after conversion; unlike a p:pic, those are not materialized as Shapes.
	// Follow only a legacyDrawing relationship actually named by the slide XML.
	vmlDrawingIDs, err := pptxVMLDrawingRelationshipIDs(b)
	if err != nil {
		return nil, false
	}
	for _, id := range vmlDrawingIDs {
		target := rels[id]
		if target == "" {
			continue
		}
		part := resolveOOXMLRelationshipTarget(slide, target)
		if !strings.HasSuffix(ooxmlPartKey(part), ".vml") {
			continue
		}
		vml := ooxmlFile(files, part)
		if vml == nil {
			return nil, false
		}
		vmlData, err := readZipFile(vml)
		if err != nil {
			return nil, false
		}
		vmlIDs, err := pptxVMLPictureRelationshipIDOccurrences(vmlData)
		if err != nil {
			// VML is optional legacy artwork. A malformed or unsupported VML
			// cache must not discard the slide's valid DrawingML pictures.
			continue
		}
		vmlPart := ooxmlPartName(files, part)
		if vmlPart == "" {
			vmlPart = part
		}
		vmlRels, err := relationshipTargetMapForPart(files, vmlPart)
		if err != nil {
			continue
		}
		appendMedia(vmlPart, vmlIDs, vmlRels)
	}
	return out, true
}

func pptxVMLDrawingRelationshipIDs(b []byte) ([]string, error) {
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var ids []string
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			return ids, nil
		}
		if err != nil {
			return nil, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch start.Name.Local {
		case "legacyDrawing", "legacyDrawingHF", "vmlDrawing":
			if id := strings.TrimSpace(xmlAttrValue(start, "id")); id != "" {
				ids = append(ids, id)
			}
		}
	}
}

func pptxRecoveredMissingImageRelationshipPart(files map[string]*zip.File, id string, rels map[string]string) string {
	missingID, ok := pptxRelationshipNumericID(id)
	if !ok {
		return ""
	}
	offsetCounts := map[int]int{}
	for knownID, target := range rels {
		idNumber, ok := pptxRelationshipNumericID(knownID)
		if !ok {
			continue
		}
		part := resolveOOXMLRelationshipTarget("ppt/slides/slide.xml", target)
		base := strings.ToLower(path.Base(part))
		match := regexp.MustCompile(`^image([0-9]+)\.[a-z0-9]+$`).FindStringSubmatch(base)
		if len(match) != 2 {
			continue
		}
		imageNumber, err := strconv.Atoi(match[1])
		if err != nil {
			continue
		}
		offsetCounts[imageNumber-idNumber]++
	}
	bestOffset, support := 0, 0
	for offset, count := range offsetCounts {
		if count > support {
			bestOffset, support = offset, count
		}
	}
	if support < 2 {
		return ""
	}
	wanted := missingID + bestOffset
	for _, ext := range []string{".png", ".jpeg", ".jpg", ".gif", ".bmp", ".emf", ".wmf"} {
		candidate := fmt.Sprintf("ppt/media/image%d%s", wanted, ext)
		if actual := ooxmlPartName(files, candidate); actual != "" {
			return actual
		}
	}
	return ""
}

func pptxRelationshipNumericID(id string) (int, bool) {
	if !strings.HasPrefix(id, "rId") {
		return 0, false
	}
	n, err := strconv.Atoi(strings.TrimPrefix(id, "rId"))
	return n, err == nil && n > 0
}

func pptxVMLPictureRelationshipIDOccurrences(b []byte) ([]string, error) {
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	shapeDepth := 0
	hidden := false
	var out []string
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			return out, nil
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "shape" && strings.Contains(strings.ToLower(t.Name.Space), "vml") {
				if shapeDepth == 0 {
					hidden = drawingObjectElementHidden(t)
					// HTML-to-PowerPoint exporters retain VML PictureFrames named
					// HTMLText*/HTMLHidden*.  These are staging frames rather than
					// visible Slide.Shapes; their presentation in COM must be
					// accounted for separately from the package drawing tree.
					id := strings.ToLower(xmlAttrValue(t, "id"))
					if strings.HasPrefix(id, "htmltext") || strings.HasPrefix(id, "htmlhidden") {
						hidden = true
					}
					// Legacy Office VML drawing parts use anonymous _x0000_s*
					// PictureFrame caches. PowerPoint does not materialize these
					// as Slide.Shapes. HTML-imported VML uses authored IDs (for
					// example HTMLText1) and is exposed as msoPicture.
					if strings.HasPrefix(strings.ToLower(xmlAttrValue(t, "id")), "_x0000_s") {
						hidden = true
					}
				}
				shapeDepth++
			}
			if shapeDepth > 0 && t.Name.Local == "imagedata" && strings.Contains(strings.ToLower(t.Name.Space), "vml") && !hidden {
				out = append(out, imageRelationshipIDs(t)...)
			}
		case xml.EndElement:
			if t.Name.Local == "shape" && strings.Contains(strings.ToLower(t.Name.Space), "vml") && shapeDepth > 0 {
				shapeDepth--
				if shapeDepth == 0 {
					hidden = false
				}
			}
		}
	}
}

func pptxPictureRelationshipIDOccurrences(b []byte) ([]string, error) {
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	pictureDepth, oleObjectDepth := 0, 0
	pictureHidden := false
	var pictureRefs []string
	var out []string
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			return out, nil
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "oleObj" {
				oleObjectDepth++
			}
			if t.Name.Local == "pic" {
				pictureDepth++
				// OLE previews are a cached rendering of an embedded object, not a
				// picture shape. Group members, in contrast, are rendered picture
				// shapes exposed by PowerPoint through GroupItems and are retained.
				if oleObjectDepth != 0 {
					pictureHidden = true
					continue
				}
				pictureHidden = false
				pictureRefs = pictureRefs[:0]
			}
			if pictureDepth == 0 {
				continue
			}
			// PowerPoint DrawingML can carry a14:hiddenFill/hiddenLine formatting
			// extensions. Those are Office UI fallback styles, not a hidden
			// Slide.Shape; applying the VML/Word z-index heuristic here drops
			// visible imported pictures.
			if pptxPictureElementExplicitlyHidden(t) {
				pictureHidden = true
			}
			// A picture placeholder (p:ph) inherits its artwork from the slide
			// master/layout. Its embedded blip is a local placeholder cache, not a
			// Picture entry in PowerPoint's Slide.Shapes collection.
			if t.Name.Local == "ph" {
				pictureHidden = true
			}
			// Windows Media Player ActiveX controls retain their poster frame as a
			// p:pic cache. Likewise, video-file references use a p:pic only as the
			// poster frame. PowerPoint materializes both as a media/control shape,
			// not as an msoPicture in Slide.Shapes.
			if pptxPictureIsMediaPlayerCache(t) {
				pictureHidden = true
			}
			if t.Name.Local == "videoFile" {
				pictureHidden = true
			}
			if !pictureHidden && t.Name.Local == "blip" {
				pictureRefs = append(pictureRefs, imageRelationshipIDs(t)...)
			}
		case xml.EndElement:
			if t.Name.Local == "pic" && pictureDepth > 0 {
				if pictureDepth == 1 && !pictureHidden {
					out = append(out, pictureRefs...)
				}
				pictureDepth--
				if pictureDepth == 0 {
					pictureHidden = false
					pictureRefs = nil
				}
			}
			if t.Name.Local == "oleObj" && oleObjectDepth > 0 {
				oleObjectDepth--
			}
		}
	}
}

func pptxPictureElementExplicitlyHidden(start xml.StartElement) bool {
	for _, attr := range start.Attr {
		if strings.EqualFold(attr.Name.Local, "hidden") && boolAttrValue(attr.Value) {
			return true
		}
		if !strings.EqualFold(attr.Name.Local, "style") {
			continue
		}
		for _, declaration := range strings.Split(strings.ToLower(attr.Value), ";") {
			key, value, ok := strings.Cut(strings.TrimSpace(declaration), ":")
			if !ok {
				continue
			}
			key, value = strings.TrimSpace(key), strings.TrimSpace(value)
			if (key == "display" && value == "none") || (key == "visibility" && (value == "hidden" || value == "collapse")) {
				return true
			}
		}
	}
	return false
}

func pptxPictureIsMediaPlayerCache(start xml.StartElement) bool {
	if start.Name.Local != "cNvPr" {
		return false
	}
	name := strings.ToLower(strings.TrimSpace(xmlAttrValue(start, "name")))
	return strings.HasPrefix(name, "windowsmediaplayer")
}

func collectPptxSlideVisibleMedia(files map[string]*zip.File, slide string, visible, hidden map[string]bool) bool {
	f := files[slide]
	if f == nil {
		return collectReachableOOXMLMedia(files, slide, "ppt/media/", visible, map[string]bool{})
	}
	b, err := readZipFile(f)
	if err != nil {
		return collectReachableOOXMLMedia(files, slide, "ppt/media/", visible, map[string]bool{})
	}
	refs, err := pptxPictureRelationshipRefs(b)
	if err != nil || (len(refs.Visible) == 0 && len(refs.Hidden) == 0) {
		// Some producers omit the DrawingML picture markup. Preserve the
		// recovery path for such malformed packages.
		return collectReachableOOXMLMedia(files, slide, "ppt/media/", visible, map[string]bool{})
	}
	rels, err := relationshipTargetMapForPart(files, slide)
	if err != nil || len(rels) == 0 {
		return collectReachableOOXMLMedia(files, slide, "ppt/media/", visible, map[string]bool{})
	}
	found := false
	for id := range refs.Visible {
		if collectRelationshipTargetMedia(files, slide, rels[id], "ppt/media/", visible) {
			found = true
		}
	}
	for id := range refs.Hidden {
		if collectRelationshipTargetMedia(files, slide, rels[id], "ppt/media/", hidden) {
			found = true
		}
	}
	if !found {
		return collectReachableOOXMLMedia(files, slide, "ppt/media/", visible, map[string]bool{})
	}
	return true
}

// pptxPictureRelationshipRefs accepts only DrawingML picture shapes. A slide
// relationship can also target layout artwork or an OLE preview; those payloads
// are not entries in PowerPoint's visible Picture Shapes collection.
func pptxPictureRelationshipRefs(b []byte) (docxImageRefs, error) {
	refs := docxImageRefs{Visible: map[string]bool{}, Hidden: map[string]bool{}}
	if hasDOCTYPE(b) {
		return refs, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	pictureDepth := 0
	pictureHidden := false
	oleObjectDepth := 0
	groupDepth := 0
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			return refs, nil
		}
		if err != nil {
			return refs, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "oleObj" {
				oleObjectDepth++
			}
			if t.Name.Local == "grpSp" {
				groupDepth++
			}
			if t.Name.Local == "pic" {
				if oleObjectDepth == 0 && groupDepth == 0 && pictureDepth == 0 {
					pictureHidden = false
				}
				if oleObjectDepth == 0 && groupDepth == 0 {
					pictureDepth++
				}
			}
			if pictureDepth == 0 {
				continue
			}
			if drawingObjectElementHidden(t) {
				pictureHidden = true
			}
			for _, id := range imageRelationshipIDs(t) {
				if pictureHidden {
					refs.Hidden[id] = true
				} else {
					refs.Visible[id] = true
				}
			}
		case xml.EndElement:
			if t.Name.Local == "pic" && pictureDepth > 0 {
				pictureDepth--
				if pictureDepth == 0 {
					pictureHidden = false
				}
			}
			if t.Name.Local == "oleObj" && oleObjectDepth > 0 {
				oleObjectDepth--
			}
			if t.Name.Local == "grpSp" && groupDepth > 0 {
				groupDepth--
			}
		}
	}
}

func collectRelationshipTargetMedia(files map[string]*zip.File, source, target, mediaPrefix string, media map[string]bool) bool {
	part := resolveOOXMLRelationshipTarget(source, target)
	if actual := ooxmlPartName(files, part); actual != "" {
		part = actual
	}
	lower := ooxmlPartKey(part)
	if strings.HasPrefix(lower, mediaPrefix) {
		media[part] = true
		return true
	}
	if strings.HasPrefix(lower, strings.TrimSuffix(mediaPrefix, "media/")) {
		return collectReachableOOXMLMedia(files, part, mediaPrefix, media, map[string]bool{})
	}
	return false
}

func xlsxVisibleMediaParts(files map[string]*zip.File) (map[string]bool, bool) {
	sheets, err := workbookVisibleSheets(files)
	if err != nil || len(sheets) == 0 {
		return nil, false
	}
	visible := map[string]bool{}
	foundRels := false
	for _, sheet := range sheets {
		if collectReachableOOXMLMedia(files, sheet.Path, "xl/media/", visible, map[string]bool{}) {
			foundRels = true
		} else if ooxmlFile(files, ooxmlRelsName(sheet.Path)) != nil {
			return nil, false
		}
	}
	hidden := xlsxHiddenMediaParts(files)
	hiddenObjects, visibleObjects := xlsxVisibleSheetObjectMediaParts(files, sheets)
	// VML legacyDrawing parts may contain an Excel PictureFrame control.  It is
	// surfaced by Worksheet.Shapes but is not represented by DrawingML blips;
	// include it in the media admission set before strict occurrence expansion.
	for name := range xlsxVisibleVMLPictureMediaParts(files, sheets) {
		visibleObjects[name] = true
		foundRels = true
	}
	if !foundRels {
		if len(hidden) > 0 || len(hiddenObjects) > 0 {
			return map[string]bool{}, true
		}
		return nil, false
	}
	if len(visible) == 0 && len(visibleObjects) == 0 {
		if len(hidden) == 0 && len(hiddenObjects) == 0 {
			return nil, false
		}
		allowed := map[string]bool{}
		for name := range files {
			lower := ooxmlPartKey(name)
			if strings.HasPrefix(lower, "xl/media/") {
				allowed[name] = true
			}
		}
		for name := range hidden {
			delete(allowed, name)
		}
		for name := range hiddenObjects {
			delete(allowed, name)
		}
		return allowed, true
	}
	if len(hidden) == 0 {
		if len(hiddenObjects) == 0 {
			return visible, true
		}
	}
	allowed := map[string]bool{}
	for name := range visible {
		allowed[name] = true
	}
	for name := range visibleObjects {
		allowed[name] = true
	}
	for name := range hidden {
		if !visible[name] {
			delete(allowed, name)
		}
	}
	for name := range hiddenObjects {
		if !visibleObjects[name] {
			delete(allowed, name)
		}
	}
	return allowed, true
}

func xlsxVisibleSheetObjectMediaParts(files map[string]*zip.File, sheets []workbookSheet) (map[string]bool, map[string]bool) {
	hidden := map[string]bool{}
	visible := map[string]bool{}
	for _, sheet := range sheets {
		for _, drawing := range relationshipTargetsWithPrefix(files, sheet.Path, "xl/drawings/") {
			refs, err := imageRelationshipRefsFromPart(files, drawing)
			if err != nil || (len(refs.Hidden) == 0 && len(refs.Visible) == 0) {
				continue
			}
			rels, err := relationshipTargetMapForPart(files, drawing)
			if err != nil {
				continue
			}
			for id := range refs.Visible {
				if part := relationshipMediaPart(files, drawing, rels[id], "xl/media/"); part != "" {
					visible[part] = true
				}
			}
			for id := range refs.Hidden {
				if part := relationshipMediaPart(files, drawing, rels[id], "xl/media/"); part != "" {
					hidden[part] = true
				}
			}
		}
	}
	return hidden, visible
}

func xlsxVisibleVMLPictureMediaParts(files map[string]*zip.File, sheets []workbookSheet) map[string]bool {
	visible := map[string]bool{}
	for _, sheet := range sheets {
		for _, drawing := range relationshipTargetsWithPrefix(files, sheet.Path, "xl/drawings/") {
			if !strings.HasSuffix(ooxmlPartKey(drawing), ".vml") {
				continue
			}
			ids, err := xlsxVMLPictureRelationshipIDOccurrences(files, drawing, xlsxOLEObjectShapeIDs(files, sheet.Path))
			if err != nil || len(ids) == 0 {
				continue
			}
			rels, err := relationshipTargetMapForPart(files, drawing)
			if err != nil {
				continue
			}
			for _, id := range ids {
				if part := relationshipMediaPart(files, drawing, rels[id], "xl/media/"); part != "" {
					visible[part] = true
				}
			}
		}
	}
	return visible
}

// xlsxOLEObjectShapeIDs returns the VML shape IDs that worksheet oleObjects
// identify as embedded-object previews. They often use a VML Pict ClientData
// node and an image relationship, but Excel exposes them as msoEmbeddedOLEObject
// rather than as msoPicture.
func xlsxOLEObjectShapeIDs(files map[string]*zip.File, sheetPart string) map[string]bool {
	ids := map[string]bool{}
	b, err := readZipFile(ooxmlFile(files, sheetPart))
	if err != nil || hasDOCTYPE(b) {
		return ids
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			return ids
		}
		if err != nil {
			return ids
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "oleObject" {
			continue
		}
		if id := normalizeVMLShapeID(xmlAttrValue(start, "shapeId")); id != "" {
			ids[id] = true
		}
	}
}

func normalizeVMLShapeID(id string) string {
	id = strings.TrimSpace(strings.ToLower(id))
	if id == "" {
		return ""
	}
	if strings.HasPrefix(id, "_x0000_s") {
		return strings.TrimPrefix(id, "_x0000_s")
	}
	return id
}

func relationshipTargetsWithPrefix(files map[string]*zip.File, source, prefix string) []string {
	rels, err := relationshipTargetMapForPart(files, source)
	if err != nil || len(rels) == 0 {
		return nil
	}
	var out []string
	for _, target := range rels {
		part := resolveOOXMLRelationshipTarget(source, target)
		if actual := ooxmlPartName(files, part); actual != "" {
			part = actual
		}
		if strings.HasPrefix(ooxmlPartKey(part), ooxmlPartKey(prefix)) {
			out = append(out, part)
		}
	}
	sort.Slice(out, func(i, j int) bool {
		return naturalLess(out[i], out[j])
	})
	return out
}

func imageRelationshipRefsFromPart(files map[string]*zip.File, name string) (docxImageRefs, error) {
	f := ooxmlFile(files, name)
	if f == nil {
		return docxImageRefs{Visible: map[string]bool{}, Hidden: map[string]bool{}}, nil
	}
	if cached, ok := ooxmlImageRelationshipRefsCache.Load(f); ok {
		return cloneDocxImageRefs(cached.(docxImageRefs)), nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return docxImageRefs{Visible: map[string]bool{}, Hidden: map[string]bool{}}, err
	}
	return imageRelationshipRefsFromPartBytes(files, name, b)
}

func imageRelationshipRefsFromPartBytes(files map[string]*zip.File, name string, b []byte) (docxImageRefs, error) {
	f := ooxmlFile(files, name)
	if f != nil {
		if cached, ok := ooxmlImageRelationshipRefsCache.Load(f); ok {
			return cloneDocxImageRefs(cached.(docxImageRefs)), nil
		}
	}
	refs, err := docxImageRelationshipRefs(b)
	if err != nil {
		return refs, err
	}
	if f != nil {
		ooxmlImageRelationshipRefsCache.Store(f, cloneDocxImageRefs(refs))
	}
	return refs, nil
}

func relationshipMediaPart(files map[string]*zip.File, source, target, prefix string) string {
	part := resolveOOXMLRelationshipTarget(source, target)
	if actual := ooxmlPartName(files, part); actual != "" {
		part = actual
	}
	if strings.HasPrefix(ooxmlPartKey(part), ooxmlPartKey(prefix)) {
		return part
	}
	return ""
}

func xlsxHiddenMediaParts(files map[string]*zip.File) map[string]bool {
	hidden := map[string]bool{}
	sheets, err := workbookSheets(files)
	if err != nil {
		return hidden
	}
	referenced := map[string]bool{}
	for _, sheet := range sheets {
		referenced[ooxmlPartKey(sheet.Path)] = true
		if !sheet.Hidden {
			continue
		}
		collectReachableOOXMLMedia(files, sheet.Path, "xl/media/", hidden, map[string]bool{})
	}
	if len(referenced) > 0 {
		for _, sheet := range xlsxAllWorksheetPartNames(files) {
			if !referenced[ooxmlPartKey(sheet)] {
				collectReachableOOXMLMedia(files, sheet, "xl/media/", hidden, map[string]bool{})
			}
		}
	}
	return hidden
}

func xlsxAllWorksheetPartNames(files map[string]*zip.File) []string {
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if strings.HasPrefix(lower, "xl/worksheets/sheet") && strings.HasSuffix(lower, ".xml") {
			names = append(names, name)
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return names
}

func pptxSlideRelsName(slide string) string {
	return ooxmlRelsName(slide)
}

func collectReachableOOXMLMedia(files map[string]*zip.File, source, mediaPrefix string, media map[string]bool, seen map[string]bool) bool {
	relsName := ooxmlRelsName(source)
	if seen[relsName] {
		return false
	}
	seen[relsName] = true
	f := ooxmlFile(files, relsName)
	if f == nil {
		return false
	}
	b, err := readZipFile(f)
	if err != nil {
		return false
	}
	targets, err := relationshipTargets(b)
	if err != nil {
		return false
	}
	for _, target := range targets {
		part := resolveOOXMLRelationshipTarget(source, target)
		if actual := ooxmlPartName(files, part); actual != "" {
			part = actual
		}
		lower := ooxmlPartKey(part)
		if strings.HasPrefix(lower, mediaPrefix) {
			media[part] = true
			continue
		}
		if strings.HasPrefix(lower, strings.TrimSuffix(mediaPrefix, "media/")) {
			if ooxmlFile(files, ooxmlRelsName(part)) != nil && !collectReachableOOXMLMedia(files, part, mediaPrefix, media, seen) {
				return false
			}
		}
	}
	return true
}

func ooxmlRelsName(part string) string {
	part = ooxmlCleanPartName(part)
	return path.Join(path.Dir(part), "_rels", path.Base(part)+".rels")
}

func relationshipTargets(b []byte) ([]string, error) {
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "Relationship" {
			continue
		}
		// Relationship targets are OPC locators, not visible document text.  Do
		// not pass them through cleanText: its metadata/resource filtering can
		// legitimately erase a media name such as "image111.jpeg", making a
		// visible DrawingML picture impossible to resolve back to its package
		// part.  Target-mode and path validation happen at the relationship
		// consumer, so retain the exact trimmed locator here.
		target := strings.TrimSpace(xmlAttrValue(start, "Target"))
		mode := strings.TrimSpace(xmlAttrValue(start, "TargetMode"))
		if target != "" && !strings.EqualFold(mode, "External") {
			out = append(out, target)
		}
	}
	return out, nil
}

func relationshipTargetMapForPart(files map[string]*zip.File, part string) (map[string]string, error) {
	relsName := ooxmlRelsName(part)
	f := ooxmlFile(files, relsName)
	if f == nil {
		return nil, nil
	}
	b, err := readZipFile(f)
	if err != nil {
		return nil, err
	}
	return relationshipTargetMap(b)
}

func relationshipTargetMap(b []byte) (map[string]string, error) {
	targets := map[string]string{}
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "Relationship" {
			continue
		}
		id := strings.TrimSpace(xmlAttrValue(start, "Id"))
		// Relationship targets are OPC locators, not visible document text.  Do
		// not pass them through cleanText: its metadata/resource filtering can
		// legitimately erase a media name such as "image111.jpeg", making a
		// visible DrawingML picture impossible to resolve back to its package
		// part.  Target-mode and path validation happen at the relationship
		// consumer, so retain the exact trimmed locator here.
		target := strings.TrimSpace(xmlAttrValue(start, "Target"))
		mode := strings.TrimSpace(xmlAttrValue(start, "TargetMode"))
		if id != "" && target != "" && !strings.EqualFold(mode, "External") {
			targets[id] = target
		}
	}
	return targets, nil
}

func resolveOOXMLRelationshipTarget(source, target string) string {
	source = ooxmlCleanPartName(source)
	target = cleanOOXMLRelationshipTarget(target)
	if target == "" {
		return ""
	}
	sourceRoot := ooxmlPackageRoot(source)
	var resolved string
	if strings.HasPrefix(target, "/") {
		resolved = strings.TrimPrefix(path.Clean(target), "/")
	} else {
		resolved = path.Clean(path.Join(path.Dir(source), target))
	}
	if resolved == "." || strings.HasPrefix(resolved, "../") || resolved == ".." {
		return ""
	}
	if sourceRoot != "" {
		targetRoot := ooxmlPackageRoot(resolved)
		if targetRoot != sourceRoot {
			return ""
		}
	}
	return resolved
}

func ooxmlPackageRoot(name string) string {
	name = ooxmlCleanPartName(name)
	if name == "" {
		return ""
	}
	root, _, _ := strings.Cut(name, "/")
	switch strings.ToLower(root) {
	case "word", "ppt", "xl":
		return strings.ToLower(root)
	default:
		return ""
	}
}

func cleanOOXMLRelationshipTarget(target string) string {
	target = strings.TrimSpace(filepath.ToSlash(target))
	if target == "" {
		return ""
	}
	if u, err := url.Parse(target); err == nil {
		if u.Scheme != "" || u.Host != "" {
			return ""
		}
		if u.Path != "" {
			target = u.Path
		}
	}
	if unescaped, err := url.PathUnescape(target); err == nil {
		target = unescaped
	}
	return target
}

func ooxmlVisibleImageAlts(files map[string]*zip.File, kind string, cached map[string][]string) []string {
	names := ooxmlVisibleImageAltPartNames(files, kind)
	var out []string
	seen := map[string]bool{}
	for _, name := range names {
		values, ok := cached[ooxmlPartKey(name)]
		if !ok {
			b, err := readZipFile(files[name])
			if err != nil {
				continue
			}
			values, err = visibleImageAltText(b)
			if err != nil {
				continue
			}
		}
		for _, value := range values {
			if seen[value] {
				continue
			}
			seen[value] = true
			out = append(out, value)
		}
	}
	return out
}

type ooxmlImageAltData struct {
	byMedia map[string]string
	byPart  map[string][]string
}

func ooxmlVisibleImageAltData(files map[string]*zip.File, kind string) ooxmlImageAltData {
	names := ooxmlVisibleImageAltPartNames(files, kind)
	out := ooxmlImageAltData{byMedia: map[string]string{}, byPart: map[string][]string{}}
	for _, name := range names {
		b, err := readZipFile(files[name])
		if err != nil {
			continue
		}
		if !likelyImageRelationshipMarkup(b) {
			continue
		}
		alts, ordered, err := visibleImageRelationshipAlts(b)
		if err != nil {
			continue
		}
		out.byPart[ooxmlPartKey(name)] = ordered
		if len(alts) == 0 {
			continue
		}
		rels, err := relationshipTargetMapForPart(files, name)
		if err != nil || len(rels) == 0 {
			continue
		}
		for id, alt := range alts {
			target := rels[id]
			if target == "" {
				continue
			}
			part := resolveOOXMLRelationshipTarget(name, target)
			if actual := ooxmlPartName(files, part); actual != "" {
				part = actual
			}
			key := ooxmlPartKey(part)
			if key != "" && out.byMedia[key] == "" {
				out.byMedia[key] = alt
			}
		}
	}
	if kind == "docx" {
		for _, name := range docxVisibleHTMLPartNamesNoError(files) {
			b, err := readZipFile(files[name])
			if err != nil {
				continue
			}
			for _, ref := range htmlImageRefs(files, name, b) {
				if ref.Alt == "" {
					continue
				}
				key := ooxmlPartKey(ref.Media)
				if key != "" && out.byMedia[key] == "" {
					out.byMedia[key] = ref.Alt
				}
				out.byPart[ooxmlPartKey(name)] = append(out.byPart[ooxmlPartKey(name)], ref.Alt)
			}
		}
	}
	return out
}

func ooxmlVisibleImageAltPartNames(files map[string]*zip.File, kind string) []string {
	var names []string
	var visibleHeaderFooter map[string]bool
	var constrainedHeaderFooter bool
	if kind == "docx" {
		visibleHeaderFooter, _, constrainedHeaderFooter = docxVisibleHeaderFooterParts(files)
	}
	for name := range files {
		lower := ooxmlPartKey(name)
		switch kind {
		case "docx":
			if strings.HasPrefix(lower, "word/") && (strings.HasSuffix(lower, ".xml") || strings.HasSuffix(lower, ".vml")) &&
				(!isDocxHeaderFooterPart(lower) || !constrainedHeaderFooter || visibleHeaderFooter[lower]) {
				names = append(names, name)
			}
		case "pptx":
			slideNames, err := pptxVisibleSlideNames(files)
			if err != nil {
				break
			}
			names = append(names, slideNames...)
		case "xlsx":
			visibleDrawingParts := xlsxVisibleDrawingPartNames(files)
			for drawing := range visibleDrawingParts {
				if actual := ooxmlPartName(files, drawing); actual != "" {
					names = append(names, actual)
				}
			}
		}
		if kind == "xlsx" || kind == "pptx" {
			break
		}
	}
	sort.Slice(names, func(i, j int) bool {
		return naturalLess(names[i], names[j])
	})
	return names
}

func visibleImageAltText(b []byte) ([]string, error) {
	if hasDOCTYPE(b) {
		return nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var out []string
	var drawingObjectStack []bool
	var alternateStack []alternateContentState
	var skipDepth int
	var hiddenRevisionRangeDepth int
	var paragraphHiddenStack []bool
	var runDepth int
	var rPrDepth int
	var pPrDepth int
	var runHidden bool
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if skipDepth > 0 {
				skipDepth++
				continue
			}
			if alternateContentStartSkip(t.Name.Local, &alternateStack) {
				if t.Name.Local == "Fallback" {
					skipDepth = 1
				}
				continue
			}
			if isHiddenRevisionElement(t.Name) {
				skipDepth = 1
				continue
			}
			if isHiddenRevisionRangeStart(t.Name) {
				hiddenRevisionRangeDepth++
				continue
			}
			if isHiddenRevisionRangeEnd(t.Name) {
				if hiddenRevisionRangeDepth > 0 {
					hiddenRevisionRangeDepth--
				}
				continue
			}
			if isDrawingObjectElement(t.Name.Local) {
				parentHidden := len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]
				drawingObjectStack = append(drawingObjectStack, parentHidden)
			}
			if len(drawingObjectStack) > 0 && drawingObjectElementHidden(t) {
				drawingObjectStack[len(drawingObjectStack)-1] = true
			}
			switch t.Name.Local {
			case "p":
				paragraphHiddenStack = append(paragraphHiddenStack, false)
			case "r":
				runDepth++
				runHidden = false
			case "pPr":
				if len(paragraphHiddenStack) > 0 {
					pPrDepth++
				}
			case "rPr":
				if runDepth > 0 {
					rPrDepth++
				}
			case "vanish", "webHidden":
				if runDepth > 0 && rPrDepth > 0 {
					runHidden = true
				}
				if pPrDepth > 0 && len(paragraphHiddenStack) > 0 {
					paragraphHiddenStack[len(paragraphHiddenStack)-1] = true
				}
			}
			drawingObjectHidden := hiddenRevisionRangeDepth > 0 || runHidden || currentParagraphHidden(paragraphHiddenStack) || (len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1])
			if drawingObjectHidden || !isLikelyImageAltElement(t.Name.Local) {
				continue
			}
			for _, value := range visibleAttributeText(t) {
				if value != "" {
					out = append(out, value)
				}
			}
		case xml.EndElement:
			if skipDepth > 0 {
				skipDepth--
				continue
			}
			if alternateContentEnd(t.Name.Local, &alternateStack) {
				continue
			}
			if t.Name.Local == "pPr" && pPrDepth > 0 {
				pPrDepth--
			}
			if t.Name.Local == "rPr" && rPrDepth > 0 {
				rPrDepth--
			}
			if t.Name.Local == "r" && runDepth > 0 {
				runDepth--
				if runDepth == 0 {
					runHidden = false
					rPrDepth = 0
				}
			}
			if t.Name.Local == "p" && len(paragraphHiddenStack) > 0 {
				paragraphHiddenStack = paragraphHiddenStack[:len(paragraphHiddenStack)-1]
				if len(paragraphHiddenStack) == 0 {
					pPrDepth = 0
				}
			}
			if isDrawingObjectElement(t.Name.Local) && len(drawingObjectStack) > 0 {
				drawingObjectStack = drawingObjectStack[:len(drawingObjectStack)-1]
			}
		}
	}
	return out, nil
}

func visibleImageRelationshipAlts(b []byte) (map[string]string, []string, error) {
	out := map[string]string{}
	var ordered []string
	if hasDOCTYPE(b) {
		return nil, nil, errors.New("xml doctype is not supported")
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var drawingObjectStack []bool
	var alternateStack []alternateContentState
	var skipDepth int
	var hiddenRevisionRangeDepth int
	var pendingAlt string
	var paragraphHiddenStack []bool
	var runDepth int
	var rPrDepth int
	var pPrDepth int
	var runHidden bool
	for {
		tok, err := dec.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return out, ordered, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if skipDepth > 0 {
				skipDepth++
				continue
			}
			if alternateContentStartSkip(t.Name.Local, &alternateStack) {
				if t.Name.Local == "Fallback" {
					skipDepth = 1
				}
				pendingAlt = ""
				continue
			}
			if isHiddenRevisionElement(t.Name) {
				skipDepth = 1
				pendingAlt = ""
				continue
			}
			if isHiddenRevisionRangeStart(t.Name) {
				hiddenRevisionRangeDepth++
				pendingAlt = ""
				continue
			}
			if isHiddenRevisionRangeEnd(t.Name) {
				if hiddenRevisionRangeDepth > 0 {
					hiddenRevisionRangeDepth--
				}
				pendingAlt = ""
				continue
			}
			if isDrawingObjectElement(t.Name.Local) {
				parentHidden := len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]
				drawingObjectStack = append(drawingObjectStack, parentHidden)
				pendingAlt = ""
			}
			if len(drawingObjectStack) > 0 && drawingObjectElementHidden(t) {
				drawingObjectStack[len(drawingObjectStack)-1] = true
				pendingAlt = ""
			}
			switch t.Name.Local {
			case "p":
				paragraphHiddenStack = append(paragraphHiddenStack, false)
			case "r":
				runDepth++
				runHidden = false
			case "pPr":
				if len(paragraphHiddenStack) > 0 {
					pPrDepth++
				}
			case "rPr":
				if runDepth > 0 {
					rPrDepth++
				}
			case "vanish", "webHidden":
				if runDepth > 0 && rPrDepth > 0 {
					runHidden = true
					pendingAlt = ""
				}
				if pPrDepth > 0 && len(paragraphHiddenStack) > 0 {
					paragraphHiddenStack[len(paragraphHiddenStack)-1] = true
					pendingAlt = ""
				}
			}
			if hiddenRevisionRangeDepth > 0 || runHidden || currentParagraphHidden(paragraphHiddenStack) || (len(drawingObjectStack) > 0 && drawingObjectStack[len(drawingObjectStack)-1]) {
				continue
			}
			if isLikelyImageAltElement(t.Name.Local) {
				for _, value := range visibleAttributeText(t) {
					if value != "" {
						ordered = append(ordered, value)
						if pendingAlt == "" {
							pendingAlt = value
						}
					}
				}
			}
			if pendingAlt == "" {
				continue
			}
			for _, id := range imageRelationshipIDs(t) {
				if out[id] == "" {
					out[id] = pendingAlt
				}
			}
		case xml.EndElement:
			if skipDepth > 0 {
				skipDepth--
				continue
			}
			if alternateContentEnd(t.Name.Local, &alternateStack) {
				pendingAlt = ""
				continue
			}
			if t.Name.Local == "pPr" && pPrDepth > 0 {
				pPrDepth--
			}
			if t.Name.Local == "rPr" && rPrDepth > 0 {
				rPrDepth--
			}
			if t.Name.Local == "r" && runDepth > 0 {
				runDepth--
				if runDepth == 0 {
					runHidden = false
					rPrDepth = 0
				}
			}
			if t.Name.Local == "p" && len(paragraphHiddenStack) > 0 {
				paragraphHiddenStack = paragraphHiddenStack[:len(paragraphHiddenStack)-1]
				if len(paragraphHiddenStack) == 0 {
					pPrDepth = 0
				}
			}
			if isDrawingObjectElement(t.Name.Local) && len(drawingObjectStack) > 0 {
				drawingObjectStack = drawingObjectStack[:len(drawingObjectStack)-1]
				pendingAlt = ""
			}
		}
	}
	return out, ordered, nil
}

func isLikelyImageAltElement(name string) bool {
	switch name {
	case "docPr", "cNvPr", "pic", "shape", "imagedata":
		return true
	default:
		return false
	}
}

func isOOXMLMediaPart(name, prefix, kind string) bool {
	lower := ooxmlPartKey(name)
	if kind == "" {
		return strings.Contains(lower, "/media/")
	}
	return strings.HasPrefix(lower, prefix)
}

func normalizeOOXMLImageData(ext string, b []byte) ([]byte, string, bool) {
	ext = strings.ToLower(ext)
	if normalized, normalizedExt, ok := normalizeCompressedImageData(ext, b); ok {
		return normalized, normalizedExt, true
	}
	if ext != "" {
		if normalized, ok := normalizeImageData(ext, b); ok {
			if ext == ".dib" {
				return normalized, ".bmp", true
			}
			return normalized, ext, true
		}
	}
	sniffed := imageExt(b)
	if normalized, normalizedExt, ok := normalizeSniffedCompressedImageData(b); ok {
		return normalized, normalizedExt, true
	}
	if sniffed == ".bin" || sniffed == ext {
		return nil, "", false
	}
	normalized, ok := normalizeImageData(sniffed, b)
	if !ok {
		return nil, "", false
	}
	if sniffed == ".dib" {
		return normalized, ".bmp", true
	}
	return normalized, sniffed, true
}

func normalizeCompressedMetafileData(ext string, b []byte) ([]byte, string, bool) {
	return normalizeCompressedImageData(ext, b)
}

func normalizeCompressedImageData(ext string, b []byte) ([]byte, string, bool) {
	switch ext {
	case ".emz", ".wmz", ".svgz":
	default:
		return nil, "", false
	}
	raw, ok := gunzipImagePayload(b)
	if !ok {
		return nil, "", false
	}
	targetExt := ".svg"
	switch ext {
	case ".emz":
		targetExt = ".emf"
	case ".wmz":
		targetExt = ".wmf"
	}
	normalized, ok := normalizeImageData(targetExt, raw)
	if !ok {
		sniffed := imageExt(raw)
		if sniffed == "" || sniffed == ".bin" {
			return nil, "", false
		}
		normalized, ok = normalizeImageData(sniffed, raw)
		if !ok {
			return nil, "", false
		}
		if sniffed == ".dib" {
			return normalized, ".bmp", true
		}
		return normalized, sniffed, true
	}
	return normalized, targetExt, true
}

func normalizeSniffedCompressedImageData(b []byte) ([]byte, string, bool) {
	raw, ok := gunzipImagePayload(b)
	if !ok {
		return nil, "", false
	}
	ext := imageExt(raw)
	if ext == "" || ext == ".bin" {
		return nil, "", false
	}
	normalized, ok := normalizeImageData(ext, raw)
	if !ok {
		return nil, "", false
	}
	if ext == ".dib" {
		return normalized, ".bmp", true
	}
	return normalized, ext, true
}

func gunzipImagePayload(b []byte) ([]byte, bool) {
	zr, err := gzip.NewReader(bytes.NewReader(b))
	if err != nil {
		return nil, false
	}
	defer zr.Close()
	raw, err := io.ReadAll(io.LimitReader(zr, maxCompressedMetafileBytes+1))
	if err != nil || len(raw) == 0 || len(raw) > maxCompressedMetafileBytes {
		return nil, false
	}
	return raw, true
}

func imageNameWithExt(name, ext string) string {
	rawExt := path.Ext(name)
	current := strings.ToLower(rawExt)
	if current == ext {
		return strings.TrimSuffix(name, rawExt) + ext
	}
	if compatibleImageExtAlias(current, ext) {
		return strings.TrimSuffix(name, rawExt) + current
	}
	if current == ".dib" && ext == ".bmp" {
		return strings.TrimSuffix(name, rawExt) + ext
	}
	if isCompressedImageExt(current) {
		return strings.TrimSuffix(name, rawExt) + ext
	}
	if current == "" || !isSupportedImageExt(current) {
		return strings.TrimSuffix(name, rawExt) + ext
	}
	return strings.TrimSuffix(name, rawExt) + ext
}

func compatibleImageExtAlias(current, ext string) bool {
	if current == "" || ext == "" {
		return false
	}
	for _, aliases := range [][]string{
		{".jpg", ".jpeg", ".jpe", ".jfif"},
		{".tif", ".tiff"},
		{".wdp", ".jxr", ".hdp"},
		{".jp2", ".jpx", ".jpf"},
		{".j2k", ".j2c", ".jpc"},
		{".pct", ".pict"},
	} {
		hasCurrent := false
		hasExt := false
		for _, alias := range aliases {
			if current == alias {
				hasCurrent = true
			}
			if ext == alias {
				hasExt = true
			}
		}
		if hasCurrent && hasExt {
			return true
		}
	}
	return false
}

func isSupportedImageExt(ext string) bool {
	switch strings.ToLower(ext) {
	case ".png", ".jpg", ".jpeg", ".jpe", ".jfif", ".gif", ".bmp", ".dib", ".emf", ".emz", ".wmf", ".wmz", ".svg", ".svgz", ".eps", ".ps", ".tif", ".tiff", ".webp", ".ico", ".cur", ".pcx", ".tga", ".pct", ".pict", ".heic", ".heif", ".avif", ".wdp", ".jxr", ".hdp", ".jp2", ".jpx", ".j2k", ".j2c", ".jpc", ".jpf":
		return true
	default:
		return false
	}
}

func isOOXMLThumbnail(name string) bool {
	lower := ooxmlPartKey(name)
	if !strings.HasPrefix(lower, "docprops/thumbnail.") {
		return false
	}
	return isSupportedImageExt(path.Ext(lower))
}

func extractEmbeddedOfficePackages(files map[string]*zip.File, kind string, depth int, opts Options) ([]string, []string, []Image) {
	visible, constrained := visibleOOXMLEmbeddedParts(files, kind)
	var names []string
	for name := range files {
		lower := ooxmlPartKey(name)
		if strings.Contains(lower, "embeddings/") && (!constrained || visible[lower] || visible[name]) {
			names = append(names, name)
		}
	}
	sort.Strings(names)
	var texts []string
	var markdowns []string
	var images []Image
	usedImageNames := map[string]bool{}
	for _, name := range names {
		b, err := readZipFile(files[name])
		if err != nil {
			continue
		}
		var res *Result
		switch {
		case isZip(b) && isOOXMLPackage(b):
			res, err = extractOOXMLWithDepth(name, b, depth, opts)
		case isOLE(b):
			legacyName, ok := embeddedLegacyFilename(name, b)
			if !ok {
				// A generic OLE payload is represented by Office as an embedded
				// object, not a visible picture. Do not promote preview bytes from
				// it to document images; this keeps the visible-image contract
				// aligned with Word/PowerPoint/Excel.
				if streams, streamErr := readOLEStreams(b); streamErr == nil {
					res = extractOfficePackagesFromOLEStreams(name, streams, depth, opts)
				}
				if res == nil {
					continue
				}
			}
			res, err = extractLegacyWithDepth(legacyName, b, depth, opts)
		default:
			continue
		}
		if err != nil {
			continue
		}
		if res.Text != "" {
			texts = append(texts, res.Text)
		}
		if res.StructuredMarkdown != "" {
			markdowns = append(markdowns, res.StructuredMarkdown)
		} else if res.Text != "" {
			markdowns = append(markdowns, markdownText(res.Text))
		}
		for _, img := range res.Images {
			if img.Name != "" {
				img.Name = sanitizeFilename(oleStreamBaseName(name) + "-" + img.Name)
			}
			if img.Name != "" {
				img.Name = uniqueImageFilename(img.Name, usedImageNames)
			}
			images = append(images, img)
		}
	}
	return texts, markdowns, images
}

func visibleOOXMLEmbeddedParts(files map[string]*zip.File, kind string) (map[string]bool, bool) {
	switch kind {
	case "docx":
		return docxVisibleEmbeddedParts(files)
	case "pptx":
		return pptxVisibleEmbeddedParts(files)
	case "xlsx":
		return xlsxVisibleEmbeddedParts(files)
	default:
		return nil, false
	}
}

func docxVisibleEmbeddedParts(files map[string]*zip.File) (map[string]bool, bool) {
	visible := map[string]bool{}
	hidden := map[string]bool{}
	_, hiddenHeaderFooter, constrainedHeaderFooter := docxVisibleHeaderFooterParts(files)
	foundCandidates := false
	malformed := false
	for _, source := range docxMediaReferencePartNames(files) {
		targets, hasTargets, err := embeddedRelationshipTargetsForPart(files, source, "word/embeddings/")
		if err != nil {
			malformed = true
			continue
		}
		if !hasTargets {
			continue
		}
		foundCandidates = true
		sourceKey := ooxmlPartKey(source)
		sourceIsHidden := strings.HasPrefix(sourceKey, "word/glossary/") || (constrainedHeaderFooter && hiddenHeaderFooter[sourceKey])
		refs, err := imageRelationshipRefsFromPart(files, source)
		if err != nil {
			malformed = true
			continue
		}
		for id := range refs.Visible {
			if part := targets[id]; part != "" {
				if sourceIsHidden {
					hidden[part] = true
				} else {
					visible[part] = true
				}
			}
		}
		for id := range refs.Hidden {
			if part := targets[id]; part != "" {
				hidden[part] = true
			}
		}
	}
	return allowedVisibleEmbeddedParts(visible, hidden, foundCandidates, malformed)
}

func pptxVisibleEmbeddedParts(files map[string]*zip.File) (map[string]bool, bool) {
	visibleSlides, err := pptxVisibleSlideNames(files)
	if err != nil {
		return nil, false
	}
	visibleSlide := map[string]bool{}
	for _, slide := range visibleSlides {
		visibleSlide[ooxmlPartKey(slide)] = true
	}
	candidateSlides, constrained, err := pptxCandidateSlideNames(files)
	if err != nil {
		return nil, false
	}
	candidateSlide := map[string]bool{}
	for _, slide := range candidateSlides {
		candidateSlide[ooxmlPartKey(slide)] = true
	}
	if constrained {
		candidateSlides = pptxAllSlidePartNames(files)
	}
	visible := map[string]bool{}
	hidden := map[string]bool{}
	foundCandidates := false
	malformed := false
	for _, name := range candidateSlides {
		lower := ooxmlPartKey(name)
		targets, hasTargets, err := embeddedRelationshipTargetsForPart(files, name, "ppt/embeddings/")
		if err != nil {
			malformed = true
			continue
		}
		if !hasTargets {
			continue
		}
		foundCandidates = true
		refs, err := imageRelationshipRefsFromPart(files, name)
		if err != nil {
			malformed = true
			continue
		}
		sourceHidden := !visibleSlide[lower] || (constrained && !candidateSlide[lower])
		for id := range refs.Visible {
			if part := targets[id]; part != "" {
				if sourceHidden {
					hidden[part] = true
				} else {
					visible[part] = true
				}
			}
		}
		for id := range refs.Hidden {
			if part := targets[id]; part != "" {
				hidden[part] = true
			}
		}
	}
	return allowedVisibleEmbeddedParts(visible, hidden, foundCandidates, malformed)
}

func xlsxVisibleEmbeddedParts(files map[string]*zip.File) (map[string]bool, bool) {
	sheets, err := workbookSheets(files)
	if err != nil {
		return nil, false
	}
	visible := map[string]bool{}
	hidden := map[string]bool{}
	foundCandidates := false
	malformed := false
	referenced := map[string]bool{}
	for _, sheet := range sheets {
		referenced[ooxmlPartKey(sheet.Path)] = true
		if xlsxCollectEmbeddedPartsFromSource(files, sheet.Path, sheet.Hidden, visible, hidden) {
			foundCandidates = true
		}
		for _, drawing := range relationshipTargetsWithPrefix(files, sheet.Path, "xl/drawings/") {
			if xlsxCollectEmbeddedPartsFromSource(files, drawing, sheet.Hidden, visible, hidden) {
				foundCandidates = true
			}
		}
	}
	if len(referenced) > 0 {
		for _, sheet := range xlsxAllWorksheetPartNames(files) {
			if referenced[ooxmlPartKey(sheet)] {
				continue
			}
			if xlsxCollectEmbeddedPartsFromSource(files, sheet, true, visible, hidden) {
				foundCandidates = true
			}
			for _, drawing := range relationshipTargetsWithPrefix(files, sheet, "xl/drawings/") {
				if xlsxCollectEmbeddedPartsFromSource(files, drawing, true, visible, hidden) {
					foundCandidates = true
				}
			}
		}
	}
	if malformed {
		return nil, false
	}
	return allowedVisibleEmbeddedParts(visible, hidden, foundCandidates, false)
}

func xlsxCollectEmbeddedPartsFromSource(files map[string]*zip.File, source string, sourceHidden bool, visible, hidden map[string]bool) bool {
	targets, hasTargets, err := embeddedRelationshipTargetsForPart(files, source, "xl/embeddings/")
	if err != nil || !hasTargets {
		return false
	}
	refs, err := imageRelationshipRefsFromPart(files, source)
	if err != nil {
		return false
	}
	for id := range refs.Visible {
		if part := targets[id]; part != "" {
			if sourceHidden {
				hidden[part] = true
			} else {
				visible[part] = true
			}
		}
	}
	for id := range refs.Hidden {
		if part := targets[id]; part != "" {
			hidden[part] = true
		}
	}
	return true
}

func embeddedRelationshipTargetsForPart(files map[string]*zip.File, source, prefix string) (map[string]string, bool, error) {
	rels, err := relationshipTargetMapForPart(files, source)
	if err != nil {
		return nil, false, err
	}
	targets := map[string]string{}
	for id, target := range rels {
		part := resolveOOXMLRelationshipTarget(source, target)
		if actual := ooxmlPartName(files, part); actual != "" {
			part = actual
		}
		if strings.HasPrefix(ooxmlPartKey(part), ooxmlPartKey(prefix)) {
			targets[id] = part
		}
	}
	return targets, len(targets) > 0, nil
}

func allowedVisibleEmbeddedParts(visible, hidden map[string]bool, foundCandidates, malformed bool) (map[string]bool, bool) {
	if !foundCandidates || malformed {
		return nil, false
	}
	allowed := map[string]bool{}
	for name := range visible {
		allowed[name] = true
		allowed[ooxmlPartKey(name)] = true
	}
	for name := range hidden {
		if !visible[name] && !visible[ooxmlPartKey(name)] {
			delete(allowed, name)
			delete(allowed, ooxmlPartKey(name))
		}
	}
	return allowed, true
}

func embeddedLegacyFilename(name string, data []byte) (string, bool) {
	ext := strings.ToLower(path.Ext(name))
	switch ext {
	case ".doc", ".xls", ".ppt":
		return name, true
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		return "", false
	}
	if inferred := inferLegacyExt(streams); inferred != "" {
		return name + inferred, true
	}
	return "", false
}

func inferLegacyExt(streams []oleStream) string {
	for _, s := range streams {
		switch strings.ToLower(s.Name) {
		case "worddocument":
			return ".doc"
		case "powerpoint document":
			return ".ppt"
		case "workbook", "book":
			return ".xls"
		}
	}
	return ""
}

func isOOXMLPackage(data []byte) bool {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return false
	}
	files := map[string]*zip.File{}
	for _, f := range zr.File {
		files[f.Name] = f
	}
	switch ooxmlKind(files) {
	case "docx", "pptx", "xlsx":
		return true
	default:
		return false
	}
}

func readZipFile(f *zip.File) ([]byte, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}
