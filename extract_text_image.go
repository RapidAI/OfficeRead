package officeread

import (
	"bytes"
	"encoding/binary"
	"encoding/xml"
	"html"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"net/url"
	"path/filepath"
	"sort"
	"strings"
	"unicode"
	"unicode/utf16"
	"unicode/utf8"
)

func cleanText(s string) string {
	if fast, ok := cleanTextFastPath(s); ok {
		return fast
	}
	s = strings.ToValidUTF8(s, "")
	if text, ok := rtfVisibleText(s); ok {
		s = text
	}
	s = decodeOOXMLTextEscapes(s)
	s = strings.Map(cleanTextRune, s)
	// 鎼?is an internal sentinel for the legacy Word possessive mojibake handled
	// by cleanTextRune. Expand it only after rune cleanup so ordinary Unicode
	// text is never mistaken for the sentinel.
	s = strings.ReplaceAll(s, "\uE000", "'s")
	s = spaceRE.ReplaceAllString(s, " ")
	lines := strings.Split(s, "\n")
	for i := range lines {
		lines[i] = strings.TrimSpace(lines[i])
		lines[i] = repairWindows1251MojibakeLine(lines[i])
		lines[i] = repairWindows1252UTF8MojibakePunctuationLine(lines[i])
		lines[i] = repairGBKMojibakePunctuationLine(lines[i])
		lines[i] = repairUnbalancedASCIIQuoteLine(lines[i])
		lines[i] = repairMojibakeContractionLine(lines[i])
		lines[i] = repairGBKDecodedUTF8LatinAccentsLine(lines[i])
		lines[i] = stripWordFieldInstructions(lines[i])
		if lines[i] != "" && looksLikeDiscardableBinaryControlLine(lines[i]) {
			lines[i] = ""
		}
	}
	s = strings.Join(lines, "\n")
	s = blankLineRE.ReplaceAllString(s, "\n\n")
	return strings.TrimSpace(s)
}

func rtfVisibleText(s string) (string, bool) {
	trimmed := strings.TrimSpace(s)
	lower := strings.ToLower(trimmed)
	if !strings.HasPrefix(lower, `{\rtf`) && !strings.HasPrefix(lower, `\rtf`) {
		return "", false
	}
	var out strings.Builder
	skipStack := []bool{false}
	hiddenStack := []bool{false}
	optionalDestinationStack := []bool{false}
	ucSkip := 1
	codePage := uint16(1252)
	var pendingHighSurrogate rune
	var hexBytes []byte
	flushHexBytes := func(skip bool) {
		if len(hexBytes) == 0 {
			return
		}
		if !skip {
			out.WriteString(decodeCodePageBytes(hexBytes, codePage))
		}
		hexBytes = hexBytes[:0]
	}
	skipNextGroup := false
	for i := 0; i < len(trimmed); {
		skip := skipStack[len(skipStack)-1] || hiddenStack[len(hiddenStack)-1]
		switch trimmed[i] {
		case '{':
			flushHexBytes(skip)
			groupSkip := skipStack[len(skipStack)-1]
			if skipNextGroup {
				groupSkip = true
				skipNextGroup = false
			}
			skipStack = append(skipStack, groupSkip)
			hiddenStack = append(hiddenStack, hiddenStack[len(hiddenStack)-1])
			optionalDestinationStack = append(optionalDestinationStack, false)
			i++
			continue
		case '}':
			flushHexBytes(skip)
			if len(skipStack) > 1 {
				skipStack = skipStack[:len(skipStack)-1]
				hiddenStack = hiddenStack[:len(hiddenStack)-1]
				optionalDestinationStack = optionalDestinationStack[:len(optionalDestinationStack)-1]
			}
			i++
			continue
		case '\\':
			if next, ok := parseRTFBinaryControl(trimmed, i); ok {
				flushHexBytes(skip)
				pendingHighSurrogate = 0
				i = next
				continue
			}
			if next, b, ok := parseRTFHexByteControl(trimmed, i); ok {
				pendingHighSurrogate = 0
				if !skip {
					hexBytes = append(hexBytes, b)
				}
				i = next
				continue
			}
			if next, code, n, ok := parseRTFUnicodeControl(trimmed, i, ucSkip); ok {
				flushHexBytes(skip)
				i = next
				if !skip {
					switch {
					case utf16.IsSurrogate(code) && code >= 0xd800 && code <= 0xdbff:
						pendingHighSurrogate = code
					case pendingHighSurrogate != 0 && utf16.IsSurrogate(code):
						decoded := utf16.DecodeRune(pendingHighSurrogate, code)
						if decoded != unicode.ReplacementChar {
							out.WriteRune(decoded)
						}
						pendingHighSurrogate = 0
					default:
						pendingHighSurrogate = 0
						if code != 0 && code != unicode.ReplacementChar {
							out.WriteRune(code)
						}
					}
					if n > 0 {
						i = skipRTFFallbackChars(trimmed, i, n)
					}
				}
				continue
			}
			next, n, text, controlSkip, controlCodePage, hiddenState, destination, controlWord, optionalDestination := parseRTFControl(trimmed, i, ucSkip)
			flushHexBytes(skip)
			i = next
			if optionalDestination {
				optionalDestinationStack[len(optionalDestinationStack)-1] = true
				continue
			}
			if optionalDestinationStack[len(optionalDestinationStack)-1] {
				if !rtfVisibleOptionalDestination(controlWord) {
					skipStack[len(skipStack)-1] = true
					skip = true
				}
				optionalDestinationStack[len(optionalDestinationStack)-1] = false
			}
			if controlWord == "upr" {
				skipNextGroup = true
			}
			if destination {
				skipStack[len(skipStack)-1] = true
				continue
			}
			if controlSkip >= 0 {
				ucSkip = controlSkip
			}
			if controlCodePage != 0 {
				codePage = controlCodePage
			}
			if hiddenState >= 0 {
				hiddenStack[len(hiddenStack)-1] = hiddenState > 0
				skip = skipStack[len(skipStack)-1] || hiddenStack[len(hiddenStack)-1]
			}
			if !skip && text != "" {
				pendingHighSurrogate = 0
				out.WriteString(text)
				if n > 0 {
					i = skipRTFFallbackChars(trimmed, i, n)
				}
			}
			continue
		default:
			flushHexBytes(skip)
			if !skip {
				pendingHighSurrogate = 0
				out.WriteByte(trimmed[i])
			}
			i++
		}
	}
	flushHexBytes(skipStack[len(skipStack)-1] || hiddenStack[len(hiddenStack)-1])
	text := cleanTextNoMojibakeRepair(out.String())
	if text == "" {
		return "", true
	}
	return text, true
}

func parseRTFHexByteControl(s string, i int) (next int, b byte, ok bool) {
	if i+3 >= len(s) || s[i] != '\\' || s[i+1] != '\'' {
		return 0, 0, false
	}
	b, ok = parseHexByte(s[i+2 : i+4])
	if !ok {
		return 0, 0, false
	}
	return i + 4, b, true
}

func parseRTFBinaryControl(s string, i int) (int, bool) {
	if i+4 >= len(s) || s[i] != '\\' || !strings.HasPrefix(strings.ToLower(s[i+1:]), "bin") {
		return 0, false
	}
	j := i + 4
	numStart := j
	for j < len(s) && s[j] >= '0' && s[j] <= '9' {
		j++
	}
	if j == numStart {
		return 0, false
	}
	n, ok := atoi(s[numStart:j])
	if !ok || n < 0 {
		return 0, false
	}
	if j < len(s) && s[j] == ' ' {
		j++
	}
	if n > len(s)-j {
		return len(s), true
	}
	return j + n, true
}

func parseRTFUnicodeControl(s string, i, ucSkip int) (next int, code rune, fallbackSkip int, ok bool) {
	if i+2 >= len(s) || s[i] != '\\' || s[i+1] != 'u' {
		return 0, 0, 0, false
	}
	j := i + 2
	if j < len(s) && ((s[j] >= 'A' && s[j] <= 'Z') || (s[j] >= 'a' && s[j] <= 'z')) {
		return 0, 0, 0, false
	}
	sign := 1
	if j < len(s) && s[j] == '-' {
		sign = -1
		j++
	}
	numStart := j
	for j < len(s) && s[j] >= '0' && s[j] <= '9' {
		j++
	}
	if j == numStart {
		return 0, 0, 0, false
	}
	num, ok := atoi(s[numStart:j])
	if !ok {
		return 0, 0, 0, false
	}
	num *= sign
	if num < 0 {
		num += 65536
	}
	if num < 0 || num > 65535 {
		return 0, 0, 0, false
	}
	if j < len(s) && s[j] == ' ' {
		j++
	}
	return j, rune(num), ucSkip, true
}

func parseRTFControl(s string, i, ucSkip int) (next int, fallbackSkip int, text string, newUCSkip int, newCodePage uint16, hiddenState int, destination bool, word string, optionalDestination bool) {
	newUCSkip = -1
	hiddenState = -1
	if i+1 >= len(s) {
		return i + 1, 0, "", newUCSkip, 0, hiddenState, false, "", false
	}
	c := s[i+1]
	switch c {
	case '\\', '{', '}':
		return i + 2, 0, string(c), newUCSkip, 0, hiddenState, false, "", false
	case '~':
		return i + 2, 0, " ", newUCSkip, 0, hiddenState, false, "", false
	case '_':
		return i + 2, 0, "-", newUCSkip, 0, hiddenState, false, "", false
	case '-':
		return i + 2, 0, "", newUCSkip, 0, hiddenState, false, "", false
	case '*':
		return i + 2, 0, "", newUCSkip, 0, hiddenState, false, "", true
	}
	if !((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')) {
		return i + 2, 0, "", newUCSkip, 0, hiddenState, false, "", false
	}
	j := i + 1
	for j < len(s) && ((s[j] >= 'A' && s[j] <= 'Z') || (s[j] >= 'a' && s[j] <= 'z')) {
		j++
	}
	word = strings.ToLower(s[i+1 : j])
	sign := 1
	if j < len(s) && s[j] == '-' {
		sign = -1
		j++
	}
	numStart := j
	for j < len(s) && s[j] >= '0' && s[j] <= '9' {
		j++
	}
	hasNum := j > numStart
	num := 0
	if hasNum {
		if parsed, ok := atoi(s[numStart:j]); ok {
			num = parsed * sign
		}
	}
	if j < len(s) && s[j] == ' ' {
		j++
	}
	if rtfDestinationControl(word) {
		return j, 0, "", newUCSkip, 0, hiddenState, true, word, false
	}
	switch word {
	case "par", "line", "page", "sect", "column":
		return j, 0, "\n", newUCSkip, 0, hiddenState, false, word, false
	case "tab", "cell", "enspace", "emspace", "qmspace":
		return j, 0, " ", newUCSkip, 0, hiddenState, false, word, false
	case "row":
		return j, 0, "\n", newUCSkip, 0, hiddenState, false, word, false
	case "emdash":
		return j, 0, " - ", newUCSkip, 0, hiddenState, false, word, false
	case "endash":
		return j, 0, " - ", newUCSkip, 0, hiddenState, false, word, false
	case "bullet":
		return j, 0, "* ", newUCSkip, 0, hiddenState, false, word, false
	case "lquote", "rquote":
		return j, 0, "'", newUCSkip, 0, hiddenState, false, word, false
	case "ldblquote", "rdblquote":
		return j, 0, `"`, newUCSkip, 0, hiddenState, false, word, false
	case "chftn":
		return j, 0, "[footnote] ", newUCSkip, 0, hiddenState, false, word, false
	case "chatn":
		return j, 0, "[comment] ", newUCSkip, 0, hiddenState, false, word, false
	case "uc":
		if hasNum && num >= 0 && num <= 32 {
			newUCSkip = num
		}
		return j, 0, "", newUCSkip, 0, hiddenState, false, word, false
	case "ansicpg":
		if hasNum && num > 0 && num <= 65535 {
			newCodePage = uint16(num)
		}
		return j, 0, "", newUCSkip, newCodePage, hiddenState, false, word, false
	case "fcharset":
		if hasNum {
			newCodePage = rtfCharsetCodePage(num)
		}
		return j, 0, "", newUCSkip, newCodePage, hiddenState, false, word, false
	case "v", "webhidden", "deleted":
		if hasNum && num == 0 {
			hiddenState = 0
		} else {
			hiddenState = 1
		}
		return j, 0, "", newUCSkip, 0, hiddenState, false, word, false
	case "plain":
		hiddenState = 0
		return j, 0, "", newUCSkip, 0, hiddenState, false, word, false
	case "u":
		if !hasNum {
			return j, 0, "", newUCSkip, 0, hiddenState, false, word, false
		}
		if num < 0 {
			num += 65536
		}
		r := rune(num)
		if r == 0 || r == unicode.ReplacementChar {
			return j, ucSkip, "", newUCSkip, 0, hiddenState, false, word, false
		}
		return j, ucSkip, string(r), newUCSkip, 0, hiddenState, false, word, false
	default:
		return j, 0, "", newUCSkip, 0, hiddenState, false, word, false
	}
}

func rtfCharsetCodePage(charset int) uint16 {
	switch charset {
	case 0, 1:
		return 1252
	case 77:
		return 10000
	case 128:
		return 932
	case 129:
		return 949
	case 134:
		return 936
	case 136:
		return 950
	case 161:
		return 1253
	case 162:
		return 1254
	case 163:
		return 1258
	case 177:
		return 1255
	case 178:
		return 1256
	case 186:
		return 1257
	case 204:
		return 1251
	case 238:
		return 1250
	default:
		return 0
	}
}

func rtfVisibleOptionalDestination(word string) bool {
	switch word {
	case "ud", "pntext", "listtext":
		return true
	default:
		return false
	}
}

func rtfDestinationControl(word string) bool {
	switch word {
	case "fonttbl", "colortbl", "stylesheet", "info", "pict", "object", "objdata", "datastore",
		"themedata", "colorschememapping", "generator", "revtbl", "listtable", "listoverridetable",
		"xmlnstbl", "datfield", "fldinst", "atnid", "atnauthor", "atndate", "atntime",
		"atnref", "atnparent", "atnstatus", "atnicn":
		return true
	default:
		return false
	}
}

func skipRTFFallbackChars(s string, i, n int) int {
	for n > 0 && i < len(s) {
		if s[i] == '\\' && i+1 < len(s) && s[i+1] == '\'' && i+3 < len(s) {
			i += 4
		} else {
			_, size := utf8.DecodeRuneInString(s[i:])
			if size <= 0 {
				return i
			}
			i += size
		}
		n--
	}
	return i
}

func parseHexByte(s string) (byte, bool) {
	if len(s) != 2 {
		return 0, false
	}
	var out byte
	for _, r := range s {
		out <<= 4
		switch {
		case r >= '0' && r <= '9':
			out += byte(r - '0')
		case r >= 'a' && r <= 'f':
			out += byte(r - 'a' + 10)
		case r >= 'A' && r <= 'F':
			out += byte(r - 'A' + 10)
		default:
			return 0, false
		}
	}
	return out, true
}

func cleanTextRune(r rune) rune {
	if r == utf8.RuneError {
		return -1
	}
	if isInvisibleFormatControlRune(r) {
		return -1
	}
	// Legacy Word's single-byte apostrophe/footnote control can be decoded as
	// this unrelated Hangul syllable when a compressed piece is read through a
	// mixed code page. It is not document text and otherwise glues to a word
	// (for example, "Commission鎴?policy") instead of the visible possessive.
	if r == '\ubb69' {
		return '\uE000'
	}
	if r == '\r' {
		return '\n'
	}
	if r == '\v' || r == '\f' {
		return '\n'
	}
	if r == '\t' || r == '\n' {
		return r
	}
	if unicode.IsSpace(r) {
		return ' '
	}
	if unicode.IsPrint(r) {
		return r
	}
	return -1
}

func isInvisibleFormatControlRune(r rune) bool {
	switch {
	case r == '\u034f':
		return true
	case r >= '\u180b' && r <= '\u180d':
		return true
	case r == '\u180f':
		return true
	case r >= '\ufe00' && r <= '\ufe0f':
		return true
	case r >= 0xe0100 && r <= 0xe01ef:
		return true
	default:
		return false
	}
}

func cleanTextFastPath(s string) (string, bool) {
	s = strings.TrimSpace(s)
	if s == "" {
		return "", true
	}
	if cleanTextFastPathControlFragment(s) {
		return "", false
	}
	if strings.ContainsAny(s, "/\\%[]()") || containsRIDFold(s) {
		return "", false
	}
	prevSpace := false
	for i := 0; i < len(s); i++ {
		c := s[i]
		switch {
		case c < 0x20 || c >= utf8.RuneSelf:
			return "", false
		case c == '\\' || c == '<' || c == '>' || c == '#':
			return "", false
		case c == '_' && i+1 < len(s) && (s[i+1] == 'x' || s[i+1] == 'X'):
			return "", false
		case c == ' ':
			if prevSpace {
				return "", false
			}
			prevSpace = true
		default:
			prevSpace = false
		}
	}
	return s, true
}

func cleanTextFastPathControlFragment(s string) bool {
	if !maybeControlFragmentText(s) {
		return false
	}
	return looksLikeBinaryControlFragment(s)
}

func maybeControlFragmentText(s string) bool {
	if len(s) < len("chart") || len(s) > 80 {
		return false
	}
	if isLegacyObjectReference(s) {
		return true
	}
	if looksLikeOLEIdentifierFragment(s) {
		return true
	}
	if looksLikeOLEWrapperStreamName(s) {
		return true
	}
	if looksLikeOOXMLMarkupNameFragment(s) {
		return true
	}
	switch s[0] {
	case '0', '1':
		return len(s) == len("1table") && strings.EqualFold(s[1:], "table")
	case 'a', 'A':
		return strings.EqualFold(s, "adobe acrobat document") ||
			strings.EqualFold(s, "adobe photoshop image") ||
			strings.EqualFold(s, "acrobat document") ||
			strings.EqualFold(s, "acroexch.document") ||
			hasPrefixFold(s, "acroexch.document.")
	case 'b', 'B':
		return strings.EqualFold(s, "bitmap image")
	case 'c', 'C':
		return strings.EqualFold(s, "cachelastmodifiedfactor.1") ||
			strings.EqualFold(s, "chart") ||
			strings.EqualFold(s, "current user") ||
			strings.EqualFold(s, "coreldraw") ||
			strings.EqualFold(s, "coreldraw 10.0 graphic") ||
			hasPrefixFold(s, "coreldraw.graphic.")
	case 'd', 'D':
		return strings.EqualFold(s, "document")
	case 'e', 'E':
		return strings.EqualFold(s, "endstream") ||
			strings.EqualFold(s, "endobj") ||
			strings.EqualFold(s, "equation") ||
			hasPrefixFold(s, "equation.") ||
			hasPrefixFold(s, "excel.sheet.") ||
			hasPrefixFold(s, "excel.chart.")
	case 'f', 'F':
		return hasPrefixFold(s, "forms.")
	case 'h', 'H':
		return strings.EqualFold(s, "html document") ||
			strings.EqualFold(s, "htmldocument") ||
			strings.EqualFold(s, "htmlfile")
	case 'i', 'I':
		return strings.EqualFold(s, "internet explorer_server")
	case 'm', 'M':
		return strings.EqualFold(s, "mathtype equation") ||
			strings.EqualFold(s, "macromedia flash factory object") ||
			strings.EqualFold(s, "media clip") ||
			strings.EqualFold(s, "windows media player") ||
			(hasPrefixFold(s, "mathtype ") && hasSuffixFold(s, " equation")) ||
			strings.EqualFold(s, "microsoft equation") ||
			strings.EqualFold(s, "microsoft equation 3.0") ||
			strings.EqualFold(s, "microsoft excel") ||
			strings.EqualFold(s, "microsoft excel worksheet") ||
			strings.EqualFold(s, "microsoft excel 97-2003 worksheet") ||
			strings.EqualFold(s, "microsoft excel 2007 worksheet") ||
			strings.EqualFold(s, "microsoft excel 2007 workbook") ||
			strings.EqualFold(s, "microsoft excel chart") ||
			strings.EqualFold(s, "microsoft office excel worksheet") ||
			strings.EqualFold(s, "microsoft office excel 97-2003 worksheet") ||
			strings.EqualFold(s, "microsoft office excel 2007 worksheet") ||
			strings.EqualFold(s, "microsoft office excel 2007 workbook") ||
			strings.EqualFold(s, "microsoft office powerpoint") ||
			strings.EqualFold(s, "microsoft powerpoint presentation") ||
			strings.EqualFold(s, "microsoft powerpoint 97-2003 presentation") ||
			strings.EqualFold(s, "microsoft powerpoint 2007 presentation") ||
			strings.EqualFold(s, "microsoft office powerpoint presentation") ||
			strings.EqualFold(s, "microsoft office powerpoint 97-2003 presentation") ||
			strings.EqualFold(s, "microsoft office powerpoint 2007 presentation") ||
			strings.EqualFold(s, "microsoft powerpoint slide") ||
			strings.EqualFold(s, "microsoft word document") ||
			strings.EqualFold(s, "microsoft word 97-2003 document") ||
			strings.EqualFold(s, "microsoft word 2007 document") ||
			strings.EqualFold(s, "microsoft office word document") ||
			strings.EqualFold(s, "microsoft office word 97-2003 document") ||
			strings.EqualFold(s, "microsoft office word 2007 document") ||
			strings.EqualFold(s, "microsoft graph chart") ||
			strings.EqualFold(s, "microsoft graph 97 chart") ||
			strings.EqualFold(s, "microsoft graph 2000 chart") ||
			hasPrefixFold(s, "microsoft forms 2.0 ") ||
			strings.EqualFold(s, "microsoft photo editor 3.0 photo") ||
			strings.EqualFold(s, "microsoft visio drawing") ||
			strings.EqualFold(s, "microsoft excel 2003 worksheet") ||
			strings.EqualFold(s, "ms graph chart") ||
			strings.EqualFold(s, "ms org chart") ||
			strings.EqualFold(s, "ms organization chart 2.0") ||
			hasPrefixFold(s, "mscomctl.") ||
			hasPrefixFold(s, "mscomctllib.") ||
			hasPrefixFold(s, "mscomct2.") ||
			hasPrefixFold(s, "mscomctl2.") ||
			hasPrefixFold(s, "msforms.") ||
			hasPrefixFold(s, "msgraph.chart.") ||
			hasPrefixFold(s, "ms_clipart_gallery.") ||
			hasPrefixFold(s, "msphotoed.") ||
			hasPrefixFold(s, "mediaplayer.mediaplayer.")
	case 'o', 'O':
		return strings.EqualFold(s, "organization chart") ||
			strings.EqualFold(s, "outlook.fileattach") ||
			strings.EqualFold(s, "outlook.message") ||
			hasPrefixFold(s, "orgpluswopx.") ||
			hasPrefixFold(s, "outlook.fileattach.") ||
			hasPrefixFold(s, "outlook.message.")
	case 'p', 'P':
		return strings.EqualFold(s, "package") ||
			strings.EqualFold(s, "package object") ||
			strings.EqualFold(s, "packager shell object") ||
			strings.EqualFold(s, "paint.picture") ||
			strings.EqualFold(s, "pdf document") ||
			strings.EqualFold(s, "photo editor photo") ||
			strings.EqualFold(s, "pictures") ||
			strings.EqualFold(s, "powerpoint document") ||
			strings.EqualFold(s, "powerpoint presentation") ||
			strings.EqualFold(s, "powerpoint slide") ||
			strings.EqualFold(s, "presentation") ||
			hasPrefixFold(s, "powerpoint.show.") ||
			hasPrefixFold(s, "powerpoint.slide.") ||
			hasPrefixFold(s, "powerpoint.presentation.") ||
			hasPrefixFold(s, "powerpoint.template.") ||
			hasPrefixFold(s, "photoshop.image.") ||
			hasPrefixFold(s, "pdf.document.")
	case 'r', 'R':
		return strings.EqualFold(s, "root entry") ||
			strings.EqualFold(s, "richedit document") ||
			hasPrefixFold(s, "richedit.document.")
	case 's', 'S':
		return strings.EqualFold(s, "shell explorer") ||
			strings.EqualFold(s, "slide") ||
			strings.EqualFold(s, "shockwave flash object") ||
			strings.EqualFold(s, "stream") ||
			strings.EqualFold(s, "shell.explorer") ||
			hasPrefixFold(s, "shell.explorer.") ||
			strings.EqualFold(s, "smartdraw") ||
			strings.EqualFold(s, "smartdraw drawing") ||
			hasPrefixFold(s, "smartdraw.") ||
			hasPrefixFold(s, "shockwaveflash.shockwaveflash.")
	case 'v', 'V':
		return hasPrefixFold(s, "visio.drawing.")
	case 'w', 'W':
		return strings.EqualFold(s, "worddocument") ||
			strings.EqualFold(s, "wordpad document") ||
			strings.EqualFold(s, "windows media player") ||
			strings.EqualFold(s, "worksheet") ||
			hasPrefixFold(s, "wmplayer.ocx.") ||
			hasPrefixFold(s, "word.document.") ||
			hasPrefixFold(s, "wordpad.document.")
	}
	return false
}

func hasPrefixFold(s, prefix string) bool {
	return len(s) >= len(prefix) && strings.EqualFold(s[:len(prefix)], prefix)
}

func hasSuffixFold(s, suffix string) bool {
	return len(s) >= len(suffix) && strings.EqualFold(s[len(s)-len(suffix):], suffix)
}

func containsASCIIFold(s, substr string) bool {
	return indexASCIIFold(s, substr) >= 0
}

func indexASCIIFold(s, substr string) int {
	if substr == "" {
		return 0
	}
	first := asciiLower(substr[0])
	for i := 0; i+len(substr) <= len(s); i++ {
		if asciiLower(s[i]) != first {
			continue
		}
		matched := true
		for j := 1; j < len(substr); j++ {
			if asciiLower(s[i+j]) != asciiLower(substr[j]) {
				matched = false
				break
			}
		}
		if matched {
			return i
		}
	}
	return -1
}

func asciiLower(b byte) byte {
	if b >= 'A' && b <= 'Z' {
		return b + ('a' - 'A')
	}
	return b
}

func containsRIDFold(s string) bool {
	for i := 0; i+2 < len(s); i++ {
		if s[i] != 'r' && s[i] != 'R' {
			continue
		}
		if (s[i+1] == 'i' || s[i+1] == 'I') && (s[i+2] == 'd' || s[i+2] == 'D') {
			return true
		}
	}
	return false
}

func maybeHiddenOrControlText(s string) bool {
	if s == "" {
		return false
	}
	switch s[0] {
	case '.', '/', '\\', '[', '<', '%', '#':
		return true
	}
	if len(s) >= 3 && s[1] == ':' && (s[2] == '/' || s[2] == '\\') {
		return true
	}
	if strings.ContainsAny(s, "/\\<>#%") {
		return true
	}
	return maybeControlFragmentText(s)
}

func decodeOOXMLTextEscapes(s string) string {
	if !strings.Contains(s, "_x") && !strings.Contains(s, "_X") {
		return s
	}
	var out strings.Builder
	out.Grow(len(s))
	changed := false
	for i := 0; i < len(s); {
		r, size := utf8.DecodeRuneInString(s[i:])
		if r == '_' && i+7 <= len(s) && (s[i+1] == 'x' || s[i+1] == 'X') && s[i+6] == '_' {
			if v, ok := parseOOXMLHex4(s[i+2 : i+6]); ok {
				if utf16.IsSurrogate(v) && i+14 <= len(s) && s[i+7] == '_' && (s[i+8] == 'x' || s[i+8] == 'X') && s[i+13] == '_' {
					if next, ok := parseOOXMLHex4(s[i+9 : i+13]); ok && utf16.IsSurrogate(next) {
						decoded := utf16.DecodeRune(v, next)
						if decoded != unicode.ReplacementChar {
							out.WriteRune(decoded)
							i += 14
							changed = true
							continue
						}
					}
				}
				writeOOXMLTextEscape(&out, v)
				i = skipRawNewlineAfterOOXMLBreak(s, i+7, v)
				changed = true
				continue
			}
		}
		out.WriteString(s[i : i+size])
		i += size
	}
	if !changed {
		return s
	}
	return out.String()
}

func skipRawNewlineAfterOOXMLBreak(s string, i int, v rune) int {
	if !isOOXMLWhitespaceEscape(v) || v == '\t' || i >= len(s) {
		return i
	}
	if s[i] == '\r' {
		i++
		if i < len(s) && s[i] == '\n' {
			i++
		}
		return i
	}
	if s[i] == '\n' {
		return i + 1
	}
	return i
}

func parseOOXMLHex4(s string) (rune, bool) {
	if len(s) != 4 {
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
	return v, true
}

func isOOXMLWhitespaceEscape(v rune) bool {
	return v == '\t' || v == '\n' || v == '\r' || v == '\v' || v == '\f'
}

func writeOOXMLTextEscape(out *strings.Builder, v rune) {
	switch v {
	case '\t':
		out.WriteByte('\t')
	case '\n', '\r', '\v', '\f':
		out.WriteByte('\n')
	default:
		if v < 0x20 || (v >= 0xd800 && v <= 0xdfff) {
			out.WriteByte(' ')
			return
		}
		out.WriteRune(v)
	}
}

func stripWordFieldInstructions(s string) string {
	if s == "" {
		return s
	}
	if !containsWordFieldInstructionMarker(s) {
		return s
	}
	s = wordHyperlinkFieldRE.ReplaceAllString(s, "")
	s = wordStyleRefFieldRE.ReplaceAllString(s, "")
	s = wordNamedFieldRE.ReplaceAllString(s, "")
	s = wordLinkFieldRE.ReplaceAllString(s, "")
	s = wordIncludeTextFieldRE.ReplaceAllString(s, "")
	s = wordEmbedFieldRE.ReplaceAllString(s, "")
	s = wordMacroButtonFieldRE.ReplaceAllString(s, "")
	s = wordTemplateFieldRE.ReplaceAllString(s, "")
	s = wordTOCFieldRE.ReplaceAllString(s, "")
	s = wordTOCBookmarkFieldRE.ReplaceAllString(s, "")
	s = wordInternalBookmarkRE.ReplaceAllString(s, "")
	s = wordTOCInternalBookmarkRE.ReplaceAllString(s, "")
	s = wordSeqFieldRE.ReplaceAllString(s, "")
	s = wordIndexEntryFieldRE.ReplaceAllString(s, "")
	s = wordPromptFieldRE.ReplaceAllString(s, "")
	s = wordSetFieldRE.ReplaceAllString(s, "")
	s = wordIfFieldRE.ReplaceAllString(s, "")
	s = wordBibliographyFieldRE.ReplaceAllString(s, "")
	s = wordDatabaseFieldRE.ReplaceAllString(s, "")
	s = wordAdvanceFieldRE.ReplaceAllString(s, "")
	s = wordPrivateAddinFieldRE.ReplaceAllString(s, "")
	s = wordPicturePathFieldRE.ReplaceAllString(s, "")
	s = wordSymbolFieldRE.ReplaceAllString(s, "")
	s = wordFormattedSimpleFieldRE.ReplaceAllString(s, "")
	s = wordSimpleFieldRE.ReplaceAllString(s, "")
	s = wordFormatSwitchRE.ReplaceAllString(s, "")
	s = wordMergeFormatRE.ReplaceAllString(s, "")
	s = orphanWordFieldTokenRE.ReplaceAllString(s, "")
	s = spaceRE.ReplaceAllString(s, " ")
	s = strings.ReplaceAll(s, " ,", ",")
	s = strings.ReplaceAll(s, " .", ".")
	s = strings.ReplaceAll(s, "( ", "(")
	s = strings.ReplaceAll(s, " )", ")")
	return strings.TrimSpace(s)
}

func stripInlineHiddenOfficeReferences(s string) string {
	if s == "" {
		return s
	}
	if !maybeInlineHiddenOfficeReferenceMarker(s) {
		return s
	}
	original := s
	s = stripInlineHiddenOfficeAssignments(s)
	s = stripInlineHiddenMIMEParameterizedHeaders(s)
	s = stripInlineHiddenOfficeColonAssignments(s)
	s = stripInlineHiddenOfficeURLReferences(s)
	s = stripInlineWrappedHiddenOfficeReferences(s)
	fields := strings.Fields(s)
	if len(fields) == 0 {
		return ""
	}
	out := make([]string, 0, len(fields))
	changed := s != original
	for _, field := range fields {
		if looksLikeInlineHiddenOfficeReferenceToken(field) {
			changed = true
			continue
		}
		out = append(out, field)
	}
	if !changed {
		return s
	}
	return strings.Join(out, " ")
}

func stripInlineHiddenMIMEParameterizedHeaders(s string) string {
	return inlineHiddenMIMEParameterizedHeaderRE.ReplaceAllString(s, "")
}

func stripInlineWrappedHiddenOfficeReferences(s string) string {
	return inlineWrappedHiddenOfficeReferenceRE.ReplaceAllStringFunc(s, func(match string) string {
		candidate := hiddenResourceReferenceCandidate(match)
		if looksLikeRelationshipIDReference(candidate) || looksLikeHiddenResourceReference(candidate) || looksLikeOfficeRelationshipMetadataReference(candidate) || looksLikeOfficeXMLMetadataReference(candidate) {
			return ""
		}
		return match
	})
}

func stripInlineHiddenOfficeAssignments(s string) string {
	return inlineHiddenOfficeAssignmentRE.ReplaceAllStringFunc(s, func(match string) string {
		value, ok := hiddenResourceAssignmentValue(match)
		if !ok {
			return match
		}
		candidate := hiddenResourceReferenceCandidate(value)
		if looksLikeRelationshipIDReference(candidate) || looksLikeHiddenResourceReference(candidate) || looksLikeOfficeRelationshipMetadataReference(match) || looksLikeOfficeRelationshipMetadataReference(candidate) || looksLikeOfficeXMLMetadataReference(match) || looksLikeOfficeXMLMetadataReference(candidate) {
			return ""
		}
		return match
	})
}

func stripInlineHiddenOfficeColonAssignments(s string) string {
	return inlineHiddenOfficeColonAssignmentRE.ReplaceAllStringFunc(s, func(match string) string {
		value, ok := hiddenResourceAssignmentValue(match)
		if !ok {
			return match
		}
		candidate := hiddenResourceReferenceCandidate(value)
		if looksLikeOfficeMHTMLHeaderReference(match, candidate) || looksLikeRelationshipIDReference(candidate) || looksLikeHiddenResourceReference(candidate) || looksLikeOfficeRelationshipMetadataReference(match) || looksLikeOfficeRelationshipMetadataReference(candidate) || looksLikeOfficeXMLMetadataReference(match) || looksLikeOfficeXMLMetadataReference(candidate) {
			return ""
		}
		return match
	})
}

func looksLikeOfficeMHTMLHeaderReference(match, value string) bool {
	key := match
	if i := strings.IndexByte(key, ':'); i >= 0 {
		key = key[:i]
	}
	key = strings.ToLower(strings.TrimSpace(key))
	value = strings.ToLower(strings.TrimSpace(strings.Trim(value, `"'<>`)))
	switch key {
	case "content-id", "content-transfer-encoding", "content-disposition", "content-description", "content-base", "mime-version":
		return true
	case "content-location":
		return value == "" || looksLikeHiddenResourceReference(value) || looksLikeOfficeXMLMetadataReference(value)
	case "content-type":
		return strings.HasPrefix(value, "application/vnd.") ||
			strings.HasPrefix(value, "application/xml") ||
			strings.HasPrefix(value, "text/xml") ||
			strings.HasPrefix(value, "image/")
	default:
		return false
	}
}

func stripInlineHiddenOfficeURLReferences(s string) string {
	return inlineHiddenOfficeURLReferenceRE.ReplaceAllStringFunc(s, func(match string) string {
		candidate := hiddenResourceReferenceCandidate(match)
		if looksLikeRelationshipIDReference(candidate) || looksLikeHiddenResourceReference(candidate) || looksLikeOfficeRelationshipMetadataReference(candidate) || looksLikeOfficeXMLMetadataReference(candidate) {
			return ""
		}
		return match
	})
}

func maybeInlineHiddenOfficeReferenceMarker(s string) bool {
	if s == "" {
		return false
	}
	hasEq := strings.IndexByte(s, '=') >= 0
	hasColon := strings.IndexByte(s, ':') >= 0
	hasSlash := strings.IndexByte(s, '/') >= 0 || strings.IndexByte(s, '\\') >= 0
	hasWrap := strings.ContainsAny(s, "<>{}[]()")
	hasPercent := strings.IndexByte(s, '%') >= 0
	if !hasEq && !hasColon && !hasSlash && !hasWrap && !hasPercent && !containsRIDFold(s) && !containsASCIIFold(s, "url(") {
		return false
	}
	if containsRIDFold(s) || containsASCIIFold(s, "url(") {
		return true
	}
	if hasEq {
		for _, marker := range []string{
			"target=", "targetmode=", "type=", "contenttype=", "partname=",
			"href=", "src=", "embed=", "link=", "xmlns", "schemaLocation", "mc:Ignorable",
		} {
			if containsASCIIFold(s, marker) {
				return true
			}
		}
	}
	if hasColon {
		for _, marker := range []string{
			"content-", "contenttype:", "partname:", "targetmode:", "target:", "type:",
			"mime-version", "schemaLocation", "mc:Ignorable", "r:id", "r:embed", "r:link",
			"file:/", "file:\\", "pack://", "opc://", "ms-word:", "cid:", "mid:",
		} {
			if containsASCIIFold(s, marker) {
				return true
			}
		}
	}
	if hasSlash {
		if strings.Contains(s, "../") || strings.Contains(s, "..\\") || strings.Contains(s, "://") {
			return true
		}
		for _, marker := range []string{"word/", "ppt/", "xl/", "_rels/", "media/", "[content_types].xml"} {
			if containsASCIIFold(s, marker) {
				return true
			}
		}
	}
	if containsASCIIFold(s, "[content_types].xml") {
		return true
	}
	if hasPercent && (containsASCIIFold(s, "%2f") || containsASCIIFold(s, "%5c")) {
		return true
	}
	return hasWrap && maybeHiddenOrControlText(s)
}

func looksLikeInlineHiddenOfficeReferenceToken(s string) bool {
	if !maybeInlineHiddenOfficeReferenceMarker(s) {
		return false
	}
	s = strings.TrimSpace(strings.ReplaceAll(s, "\\", "/"))
	if s == "" {
		return false
	}
	if looksLikeRelationshipIDReference(s) {
		return true
	}
	seen := map[string]bool{}
	queue := []string{s}
	for len(queue) > 0 && len(seen) < 32 {
		cur := strings.TrimSpace(strings.ReplaceAll(queue[0], "\\", "/"))
		queue = queue[1:]
		cur = strings.TrimSpace(strings.TrimRight(cur, ".,;:"))
		if cur == "" || seen[cur] {
			continue
		}
		seen[cur] = true
		if looksLikeInlineHiddenResourceReferencePlain(cur) {
			return true
		}
		if looksLikeOfficePartPath(strings.ToLower(strings.TrimPrefix(cur, "/"))) {
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

func looksLikeInlineHiddenResourceReferencePlain(s string) bool {
	lower := strings.ToLower(strings.TrimSpace(s))
	return looksLikeLocalFileURIReference(lower) || strings.HasPrefix(s, "//") || hiddenPackageURIPathCandidate(s) != ""
}

func looksLikeLocalFileURIReference(lower string) bool {
	lower = strings.TrimSpace(lower)
	return strings.HasPrefix(lower, "file:/") || strings.Contains(lower, "|file:/")
}

func containsWordFieldInstructionMarker(s string) bool {
	for _, marker := range []string{
		"HYPERLINK", "INCLUDEPICTURE", "PAGEREF", "REF", "NOTEREF",
		"MERGEFIELD", "DOCPROPERTY", "DOCVARIABLE", "STYLEREF", "INCLUDETEXT",
		"LINK", "EMBED", "MACROBUTTON", "TEMPLATE", "AUTHOR",
		"CREATEDATE", "DATE", "TIME", "FILENAME", "FILESIZE",
		"EDITTIME", "PAGE", "NUMPAGES", "SECTIONPAGES", "SUBJECT",
		"KEYWORDS", "COMMENTS", "LASTSAVEDBY", "PRINTDATE", "SAVEDATE",
		"USERNAME", "USERINITIALS", "SHAPE", "FORMTEXT", "FORMCHECKBOX",
		"SYMBOL", "QUOTE", "AUTOTEXT", "AUTOTEXTLIST", "LISTNUM", "AUTONUM", "AUTONUMLGL", "AUTONUMOUT", "RD",
		"TOC", "SEQ", "XE", "TC", "TA", "ASK", "FILLIN",
		"SET", "IF", "CITATION", "BIBLIOGRAPHY", "DATABASE", "ADVANCE", "ADDIN", "PRIVATE", "MERGEFORMAT", "__RefHeading__", "_Toc",
	} {
		if strings.Contains(s, marker) {
			return true
		}
	}
	return false
}

func repairWindows1251MojibakeLine(s string) string {
	if s == "" {
		return s
	}
	var letters, latin1Letters int
	for _, r := range s {
		if !unicode.IsLetter(r) {
			continue
		}
		letters++
		if r >= 0x00c0 && r <= 0x00ff {
			latin1Letters++
		}
	}
	if latin1Letters < 4 || latin1Letters*2 < letters {
		return s
	}
	raw, ok := windows1252StringBytes(s)
	if !ok || countHighBytes(raw) < 4 {
		return s
	}
	candidate := cleanTextNoMojibakeRepair(decodeCodePageBytes(raw, 1251))
	if countCyrillicLetters(candidate) < 4 || !looksLikeTextFragment(candidate) {
		return s
	}
	return candidate
}

func repairWindows1252UTF8MojibakePunctuationLine(s string) string {
	if !strings.Contains(s, "\u9225") && !strings.Contains(s, "锟") {
		return s
	}
	replacer := strings.NewReplacer(
		"\u9225?", "'",
		"\u9225\u6dd5", "\"G",
		"\u9225\u6de9", "\"R",
		"\u9225\u6dea", "\"S",
		"\u9225\u6a9a", "'s",
		"\u9225\u6e04", "\"d",
		"\u9225\u6e22", "\"t",
		"閳ユ珐", "\"R",
		"閳ユ藩", "\"S",
		"閳ユ窌", "\"G",
		"閳ユ笜", "\"n",
		"閳ユ獨", "'s",
		"閳ユ獩", "'t",
		"閳ユ獧", "'r",
		"閳ユ獡", "'m",
		"閳ユ獟", "'l",
		"閳ユ獓", "'d",
		"閳ユ獫", "'v",
		"锟", "'",
	)
	s = replacer.Replace(s)
	return strings.TrimSpace(s)
}
func countHighBytes(raw []byte) int {
	var n int
	for _, b := range raw {
		if b >= 0x80 {
			n++
		}
	}
	return n
}

func repairGBKMojibakePunctuationLine(s string) string {
	if !strings.ContainsRune(s, '\u9225') {
		return s
	}
	replacer := strings.NewReplacer(
		"\u9225\u6dd5", "\"G",
		"\u9225\u6e04", "\"d",
		"\u9225\u6e22", "\"t",
		"\u9225?", "\"",
	)
	return replacer.Replace(s)
}

func repairUnbalancedASCIIQuoteLine(s string) string {
	if strings.Count(s, "\"")%2 == 0 || !strings.ContainsRune(s, '\'') {
		return s
	}
	runes := []rune(s)
	lastQuote := -1
	for i, r := range runes {
		if r == '"' {
			lastQuote = i
		}
	}
	if lastQuote < 0 {
		return s
	}
	for i := lastQuote + 1; i < len(runes); i++ {
		if runes[i] != '\'' {
			continue
		}
		if i > 0 && isASCIILetter(runes[i-1]) && (i+1 == len(runes) || unicode.IsSpace(runes[i+1]) || strings.ContainsRune(".,;:!?)", runes[i+1])) {
			runes[i] = '"'
			return string(runes)
		}
	}
	return s
}

func repairMojibakeContractionLine(s string) string {
	if !strings.Contains(s, "\u2019c") {
		return s
	}
	return strings.NewReplacer(
		"It\u2019c", "It\u2019s",
		"it\u2019c", "it\u2019s",
		"That\u2019c", "That\u2019s",
		"that\u2019c", "that\u2019s",
	).Replace(s)
}

func repairGBKDecodedUTF8LatinAccentsLine(s string) string {
	if s == "" {
		return s
	}
	accentMap := map[rune]rune{
		'\u8c29': '\u00e1',
		'\u8305': '\u00e9',
		'\u94c6': '\u00ed',
		'\u8d38': '\u00f3',
		'\u7164': '\u00fa',
		'\u5e3d': '\u00f1',
		'\u8302': '\u00e3',
		'\u83bd': '\u00f5',
		'\u8d42': '\u00f6',
		'\u9c81': '\u00fc',
		'\u8c28': '\u00e0',
		'\u7f8c': '\u00e8',
		'\u8d32': '\u00f2',
		'\u94c3': '\u00ec',
		'\u8def': '\u00f9',
	}
	runes := []rune(s)
	changed := false
	for i, r := range runes {
		repl, ok := accentMap[r]
		if !ok {
			continue
		}
		if (i > 0 && isASCIILetter(runes[i-1])) || (i+1 < len(runes) && isASCIILetter(runes[i+1])) {
			runes[i] = repl
			changed = true
		}
	}
	if !changed {
		return s
	}
	return string(runes)
}

func isASCIILetter(r rune) bool {
	return (r >= 'A' && r <= 'Z') || (r >= 'a' && r <= 'z')
}

func cleanTextNoMojibakeRepair(s string) string {
	s = strings.ToValidUTF8(s, "")
	s = strings.Map(cleanTextRune, s)
	s = spaceRE.ReplaceAllString(s, " ")
	lines := strings.Split(s, "\n")
	for i := range lines {
		lines[i] = strings.TrimSpace(lines[i])
	}
	s = strings.Join(lines, "\n")
	s = blankLineRE.ReplaceAllString(s, "\n\n")
	return strings.TrimSpace(s)
}

func windows1252StringBytes(s string) ([]byte, bool) {
	out := make([]byte, 0, len(s))
	for _, r := range s {
		if r <= 0x7f || (r >= 0x00a0 && r <= 0x00ff) {
			out = append(out, byte(r))
			continue
		}
		if b, ok := windows1252RuneByte(r); ok {
			out = append(out, b)
			continue
		}
		return nil, false
	}
	return out, true
}

func windows1252RuneByte(r rune) (byte, bool) {
	for b := 0x80; b <= 0x9f; b++ {
		if mapped, ok := windows1252ByteRune(byte(b)); ok && mapped == r {
			return byte(b), true
		}
	}
	return 0, false
}

func countCyrillicLetters(s string) int {
	var n int
	for _, r := range s {
		if unicode.Is(unicode.Cyrillic, r) && unicode.IsLetter(r) {
			n++
		}
	}
	return n
}

func joinText(parts []string) string {
	var out []string
	for _, p := range parts {
		p = strings.TrimSpace(cleanVisibleText(p))
		if p != "" {
			out = append(out, p)
		}
	}
	return cleanVisibleText(strings.Join(out, "\n"))
}

func joinCleanedText(parts []string) string {
	var out strings.Builder
	if len(parts) > 0 {
		size := len(parts) - 1
		for _, p := range parts {
			size += len(p)
		}
		out.Grow(size)
	}
	for _, p := range parts {
		p = strings.TrimSpace(p)
		if p == "" {
			continue
		}
		if maybeHiddenOrControlText(p) && (looksLikeBinaryControlFragment(p) || looksLikeHiddenResourceReference(p)) {
			continue
		}
		if out.Len() > 0 {
			out.WriteByte('\n')
		}
		out.WriteString(p)
	}
	return strings.TrimSpace(out.String())
}

func imageExt(b []byte) string {
	switch {
	case len(b) >= 8 && bytes.Equal(b[:8], []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1a, '\n'}):
		return ".png"
	case len(b) >= 3 && bytes.Equal(b[:3], []byte{0xff, 0xd8, 0xff}):
		return ".jpg"
	case len(b) >= 4 && bytes.Equal(b[:4], []byte("GIF8")):
		return ".gif"
	case len(b) >= 2 && bytes.Equal(b[:2], []byte("BM")):
		return ".bmp"
	case validEMFData(b):
		return ".emf"
	case validWMFData(b):
		return ".wmf"
	case validSVGData(b):
		return ".svg"
	case validEPSData(b):
		return ".eps"
	case validTIFFData(b):
		return ".tif"
	case validWebPData(b):
		return ".webp"
	case validICOData(b):
		return ".ico"
	case validCURData(b):
		return ".cur"
	case validPCXData(b):
		return ".pcx"
	case validTGAData(b):
		return ".tga"
	case validPICTData(b):
		return ".pict"
	case validDIBData(b):
		return ".dib"
	case validJPEGXRData(b):
		return ".jxr"
	default:
		if ext, ok := jpeg2000ImageExt(b); ok {
			return ext
		}
		if ext, ok := isoBMFFImageExt(b); ok {
			return ext
		}
		return ".bin"
	}
}

func validImageData(ext string, b []byte) bool {
	_, ok := normalizeImageData(ext, b)
	return ok
}

func normalizeImageData(ext string, b []byte) ([]byte, bool) {
	ext = strings.ToLower(ext)
	switch ext {
	case ".png", ".jpg", ".jpeg", ".jpe", ".jfif", ".gif":
		return normalizeRasterImageData(ext, b)
	case ".bmp":
		size, ok := bmpDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".dib":
		return dibToBMP(b)
	case ".emf":
		size, ok := emfDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".wmf":
		size, ok := wmfDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".svg":
		size, ok := svgDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".eps", ".ps":
		size, ok := epsDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".jp2", ".jpx", ".jpf":
		size, ok := jpeg2000ContainerDeclaredSizeForExt(ext, b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".j2k", ".j2c", ".jpc":
		size, ok := jpeg2000CodestreamDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".tif", ".tiff":
		size, ok := tiffDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".webp":
		size, ok := webpDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".ico":
		size, ok := icoDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".cur":
		size, ok := curDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".pcx":
		size, ok := pcxDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".tga":
		size, ok := tgaDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".pct", ".pict":
		size, ok := pictDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".heic", ".heif", ".avif":
		size, ok := isoBMFFDeclaredSizeForExt(ext, b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	case ".wdp", ".jxr", ".hdp":
		size, ok := jpegXRDeclaredSize(b)
		if !ok {
			return nil, false
		}
		return b[:size], true
	default:
		return nil, false
	}
}

func normalizeRasterImageData(ext string, b []byte) ([]byte, bool) {
	size, ok := rasterImageEndOffset(ext, b)
	if !ok {
		return nil, false
	}
	normalized := b[:size]
	if _, _, err := image.DecodeConfig(bytes.NewReader(normalized)); err != nil {
		return nil, false
	}
	return normalized, true
}

func rasterImageEndOffset(ext string, b []byte) (int, bool) {
	switch strings.ToLower(ext) {
	case ".png":
		return pngEndOffset(b)
	case ".jpg", ".jpeg", ".jpe", ".jfif":
		return jpegEndOffset(b)
	case ".gif":
		return gifEndOffset(b)
	default:
		return 0, false
	}
}

func imageEndOffset(ext string, b []byte) (int, bool) {
	switch strings.ToLower(ext) {
	case ".bmp":
		return bmpDeclaredSize(b)
	case ".tif", ".tiff":
		return tiffDeclaredSize(b)
	case ".webp":
		return webpDeclaredSize(b)
	case ".ico":
		return icoDeclaredSize(b)
	case ".cur":
		return curDeclaredSize(b)
	case ".pcx":
		return pcxDeclaredSize(b)
	case ".tga":
		return tgaDeclaredSize(b)
	case ".eps", ".ps":
		return epsDeclaredSize(b)
	case ".pct", ".pict":
		return pictDeclaredSize(b)
	case ".heic", ".heif", ".avif":
		return isoBMFFDeclaredSizeForExt(ext, b)
	case ".wdp", ".jxr", ".hdp":
		return jpegXRDeclaredSize(b)
	case ".jp2", ".jpx", ".jpf":
		return jpeg2000ContainerDeclaredSizeForExt(ext, b)
	case ".j2k", ".j2c", ".jpc":
		return jpeg2000CodestreamDeclaredSize(b)
	default:
		return rasterImageEndOffset(ext, b)
	}
}

func pngEndOffset(b []byte) (int, bool) {
	if len(b) < 8 || !bytes.Equal(b[:8], []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1a, '\n'}) {
		return 0, false
	}
	for off := 8; ; {
		if off+12 > len(b) {
			return 0, false
		}
		chunkLen := binary.BigEndian.Uint32(b[off:])
		if chunkLen > uint32(len(b)-off-12) {
			return 0, false
		}
		size := int(chunkLen)
		end := off + 12 + size
		if bytes.Equal(b[off+4:off+8], []byte("IEND")) {
			return end, chunkLen == 0
		}
		off = end
	}
}

func jpegEndOffset(b []byte) (int, bool) {
	if len(b) < 2 || b[0] != 0xff || b[1] != 0xd8 {
		return 0, false
	}
	off := 2
outer:
	for off < len(b) {
		if b[off] != 0xff {
			return 0, false
		}
		for off < len(b) && b[off] == 0xff {
			off++
		}
		if off >= len(b) {
			return 0, false
		}
		marker := b[off]
		off++
		if marker == 0xd9 {
			return off, true
		}
		if marker == 0x00 {
			return 0, false
		}
		if marker == 0x01 || (marker >= 0xd0 && marker <= 0xd7) {
			continue
		}
		if off+2 > len(b) {
			return 0, false
		}
		size := int(binary.BigEndian.Uint16(b[off:]))
		if size < 2 || off+size > len(b) {
			return 0, false
		}
		off += size
		if marker != 0xda {
			continue
		}
		for off+1 < len(b) {
			if b[off] != 0xff {
				off++
				continue
			}
			next := b[off+1]
			switch {
			case next == 0x00:
				off += 2
			case next == 0xff:
				off++
			case next >= 0xd0 && next <= 0xd7:
				off += 2
			case next == 0xd9:
				return off + 2, true
			default:
				continue outer
			}
		}
		return 0, false
	}
	return 0, false
}

func gifEndOffset(b []byte) (int, bool) {
	if len(b) < 13 || (!bytes.Equal(b[:6], []byte("GIF87a")) && !bytes.Equal(b[:6], []byte("GIF89a"))) {
		return 0, false
	}
	if binary.LittleEndian.Uint16(b[6:]) == 0 || binary.LittleEndian.Uint16(b[8:]) == 0 {
		return 0, false
	}
	off := 13
	if b[10]&0x80 != 0 {
		off += 3 * (1 << ((b[10] & 0x07) + 1))
	}
	for off < len(b) {
		switch b[off] {
		case 0x3b:
			return off + 1, true
		case 0x2c:
			if off+10 > len(b) {
				return 0, false
			}
			if binary.LittleEndian.Uint16(b[off+5:]) == 0 || binary.LittleEndian.Uint16(b[off+7:]) == 0 {
				return 0, false
			}
			packed := b[off+9]
			off += 10
			if packed&0x80 != 0 {
				off += 3 * (1 << ((packed & 0x07) + 1))
			}
			if off >= len(b) {
				return 0, false
			}
			off++
			var ok bool
			off, ok = skipGIFSubBlocks(b, off)
			if !ok {
				return 0, false
			}
		case 0x21:
			if off+2 > len(b) {
				return 0, false
			}
			var ok bool
			off, ok = skipGIFSubBlocks(b, off+2)
			if !ok {
				return 0, false
			}
		default:
			return 0, false
		}
	}
	return 0, false
}

func skipGIFSubBlocks(b []byte, off int) (int, bool) {
	for off < len(b) {
		size := int(b[off])
		off++
		if size == 0 {
			return off, true
		}
		if off+size > len(b) {
			return 0, false
		}
		off += size
	}
	return 0, false
}

func validSVGData(b []byte) bool {
	_, ok := svgDeclaredSize(b)
	return ok
}

func svgDeclaredSize(b []byte) (int, bool) {
	trimmed := bytes.TrimLeftFunc(b, unicode.IsSpace)
	prefixLen := len(b) - len(trimmed)
	if bytes.HasPrefix(trimmed, []byte{0xef, 0xbb, 0xbf}) {
		trimmed = trimmed[3:]
		prefixLen += 3
	}
	if len(trimmed) == 0 || hasDOCTYPE(trimmed) {
		return 0, false
	}
	dec := xml.NewDecoder(bytes.NewReader(trimmed))
	rootDepth := 0
	seenRoot := false
	styleDepth := 0
	for {
		tok, err := dec.Token()
		if err != nil {
			return 0, false
		}
		switch se := tok.(type) {
		case xml.StartElement:
			local := strings.ToLower(strings.TrimSpace(se.Name.Local))
			if !seenRoot {
				if local != "svg" {
					return 0, false
				}
				seenRoot = true
			}
			if svgUnsafeElementOrAttrs(se) {
				return 0, false
			}
			rootDepth++
			if local == "style" {
				styleDepth++
			}
		case xml.EndElement:
			if !seenRoot {
				return 0, false
			}
			if strings.EqualFold(se.Name.Local, "style") && styleDepth > 0 {
				styleDepth--
			}
			rootDepth--
			if rootDepth == 0 {
				size := prefixLen + int(dec.InputOffset())
				if size <= prefixLen || size > len(b) {
					return 0, false
				}
				return size, true
			}
		case xml.CharData:
			if styleDepth > 0 && unsafeSVGAttributeValue(string(se)) {
				return 0, false
			}
		}
	}
}

func svgUnsafeElementOrAttrs(start xml.StartElement) bool {
	element := strings.ToLower(strings.TrimSpace(start.Name.Local))
	if element == "script" || element == "foreignobject" {
		return true
	}
	for _, attr := range start.Attr {
		local := strings.ToLower(strings.TrimSpace(attr.Name.Local))
		value := strings.TrimSpace(strings.ToLower(attr.Value))
		if strings.HasPrefix(local, "on") {
			return true
		}
		if local == "href" {
			if element == "a" {
				if unsafeSVGLinkAttributeValue(value) {
					return true
				}
			} else if unsafeSVGAttributeValue(value) {
				return true
			}
		}
		if local == "style" {
			if unsafeSVGAttributeValue(value) {
				return true
			}
		}
	}
	return false
}

func unsafeSVGLinkAttributeValue(value string) bool {
	value = strings.TrimSpace(strings.ToLower(value))
	if value == "" {
		return false
	}
	compact := strings.Map(func(r rune) rune {
		if unicode.IsSpace(r) || r <= 0x1f || r == 0x7f {
			return -1
		}
		return r
	}, value)
	for _, s := range []string{value, compact} {
		if unsafeSVGActiveContentReference(s) {
			return true
		}
		if unsafeSVGLocalResourceReference(s) {
			return true
		}
	}
	return false
}

func unsafeSVGAttributeValue(value string) bool {
	value = strings.TrimSpace(strings.ToLower(value))
	if value == "" {
		return false
	}
	compact := strings.Map(func(r rune) rune {
		if unicode.IsSpace(r) || r <= 0x1f || r == 0x7f {
			return -1
		}
		return r
	}, value)
	for _, s := range []string{
		value,
		compact,
	} {
		if unsafeSVGActiveContentReference(s) {
			return true
		}
		if unsafeSVGExternalResourceReference(s) {
			return true
		}
		if strings.Contains(s, "expression(") {
			return true
		}
	}
	return false
}

func unsafeSVGActiveContentReference(value string) bool {
	return strings.HasPrefix(value, "javascript:") || strings.Contains(value, "javascript:") ||
		strings.HasPrefix(value, "data:text/html") || strings.Contains(value, "data:text/html") ||
		strings.HasPrefix(value, "data:application/xhtml+xml") || strings.Contains(value, "data:application/xhtml+xml") ||
		strings.HasPrefix(value, "data:image/svg+xml") || strings.Contains(value, "data:image/svg+xml")
}

func unsafeSVGLocalResourceReference(value string) bool {
	value = strings.TrimSpace(strings.ReplaceAll(value, "\\", "/"))
	if value == "" {
		return false
	}
	if looksLikeHiddenResourceReference(value) {
		return true
	}
	lower := strings.ToLower(value)
	return strings.Contains(lower, "file://") || strings.Contains(lower, "file:/")
}

func unsafeSVGExternalResourceReference(value string) bool {
	value = strings.TrimSpace(strings.ReplaceAll(value, "\\", "/"))
	if value == "" {
		return false
	}
	if unsafeSVGLocalResourceReference(value) {
		return true
	}
	lower := strings.ToLower(value)
	if strings.Contains(lower, "http://") || strings.Contains(lower, "https://") || strings.Contains(lower, "ftp://") || strings.HasPrefix(lower, "//") {
		return true
	}
	for _, ref := range svgURLReferences(value) {
		if strings.HasPrefix(strings.TrimSpace(ref), "#") {
			continue
		}
		if looksLikeHiddenResourceReference(ref) || unsafeSVGExternalResourceReference(ref) {
			return true
		}
	}
	return false
}

func svgURLReferences(value string) []string {
	var refs []string
	lower := strings.ToLower(value)
	for offset := 0; ; {
		i := strings.Index(lower[offset:], "url(")
		if i < 0 {
			break
		}
		start := offset + i + len("url(")
		end := strings.IndexByte(value[start:], ')')
		if end < 0 {
			break
		}
		ref := strings.TrimSpace(value[start : start+end])
		ref = strings.Trim(ref, `"'`)
		if ref != "" {
			refs = append(refs, ref)
		}
		offset = start + end + 1
	}
	return refs
}

func svgByteRanges(data []byte) []byteRange {
	var ranges []byteRange
	for _, sig := range [][]byte{[]byte("<svg"), []byte("<?xml")} {
		offset := 0
		for {
			i := bytes.Index(data[offset:], sig)
			if i < 0 {
				break
			}
			start := offset + i
			size, ok := svgDeclaredSize(data[start:])
			if !ok {
				offset = start + len(sig)
				continue
			}
			ranges = append(ranges, byteRange{start: start, end: start + size})
			offset = start + size
		}
	}
	sort.Slice(ranges, func(i, j int) bool {
		if ranges[i].start != ranges[j].start {
			return ranges[i].start < ranges[j].start
		}
		return ranges[i].end > ranges[j].end
	})
	var out []byteRange
	for _, r := range ranges {
		contained := false
		for _, kept := range out {
			if r.start >= kept.start && r.end <= kept.end {
				contained = true
				break
			}
		}
		if !contained {
			out = append(out, r)
		}
	}
	return out
}

func carveSVGImages(data []byte) []imageCandidate {
	var candidates []imageCandidate
	for _, r := range svgByteRanges(data) {
		img, ok := normalizeImageData(".svg", data[r.start:r.end])
		if ok {
			candidates = append(candidates, imageCandidate{start: r.start, end: r.end, ext: ".svg", data: append([]byte(nil), img...)})
		}
	}
	return candidates
}

func validEPSData(b []byte) bool {
	_, ok := epsDeclaredSize(b)
	return ok
}

func epsDeclaredSize(b []byte) (int, bool) {
	if len(b) < 32 || !bytes.HasPrefix(b, []byte("%!PS-Adobe-")) {
		return 0, false
	}
	eof := bytes.Index(b, []byte("%%EOF"))
	if eof < 0 {
		return 0, false
	}
	end := eof + len("%%EOF")
	if end < len(b) && b[end] == '\r' {
		end++
		if end < len(b) && b[end] == '\n' {
			end++
		}
	} else if end < len(b) && b[end] == '\n' {
		end++
	}
	headEnd := bytes.IndexAny(b[:end], "\r\n")
	if headEnd < 0 {
		return 0, false
	}
	header := string(b[:headEnd])
	if !strings.Contains(header, "EPSF-") {
		return 0, false
	}
	if !epsHasValidBoundingBox(b[:end]) {
		return 0, false
	}
	if epsHasUnsafeFileOperator(b[:end]) {
		return 0, false
	}
	return end, true
}

func epsHasValidBoundingBox(b []byte) bool {
	for _, line := range strings.FieldsFunc(string(b), func(r rune) bool { return r == '\r' || r == '\n' }) {
		line = strings.TrimSpace(line)
		if !strings.HasPrefix(line, "%%BoundingBox:") {
			continue
		}
		fields := strings.Fields(strings.TrimSpace(strings.TrimPrefix(line, "%%BoundingBox:")))
		if len(fields) != 4 {
			return false
		}
		var vals [4]int
		for i, f := range fields {
			v, ok := parseEPSInt(f)
			if !ok {
				return false
			}
			vals[i] = v
		}
		return vals[2] > vals[0] && vals[3] > vals[1] && vals[2]-vals[0] <= 100000 && vals[3]-vals[1] <= 100000
	}
	return false
}

func parseEPSInt(s string) (int, bool) {
	if s == "" {
		return 0, false
	}
	sign := 1
	if s[0] == '-' {
		sign = -1
		s = s[1:]
	}
	if s == "" {
		return 0, false
	}
	n := 0
	for _, r := range s {
		if r < '0' || r > '9' {
			return 0, false
		}
		n = n*10 + int(r-'0')
		if n > 1000000 {
			return 0, false
		}
	}
	return sign * n, true
}

func epsHasUnsafeFileOperator(b []byte) bool {
	for _, line := range strings.FieldsFunc(string(b), func(r rune) bool { return r == '\r' || r == '\n' }) {
		line = strings.TrimSpace(line)
		if line == "" || strings.HasPrefix(line, "%") {
			continue
		}
		if epsLineHasUnsafeFileOperator(line) {
			return true
		}
	}
	return false
}

func epsLineHasUnsafeFileOperator(line string) bool {
	for _, field := range strings.Fields(line) {
		field = strings.Trim(field, "[]{}()<>")
		switch field {
		case "run", "file", "deletefile", "renamefile", "filenameforall":
			return true
		}
	}
	return false
}

func epsByteRanges(data []byte) []byteRange {
	var ranges []byteRange
	offset := 0
	for {
		i := bytes.Index(data[offset:], []byte("%!PS-Adobe-"))
		if i < 0 {
			break
		}
		start := offset + i
		size, ok := epsDeclaredSize(data[start:])
		if !ok {
			offset = start + len("%!PS-Adobe-")
			continue
		}
		ranges = append(ranges, byteRange{start: start, end: start + size})
		offset = start + size
	}
	return ranges
}

func carveEPSImages(data []byte) []imageCandidate {
	var candidates []imageCandidate
	for _, r := range epsByteRanges(data) {
		img := data[r.start:r.end]
		if len(img) <= 32 || !validEPSData(img) {
			continue
		}
		candidates = append(candidates, imageCandidate{start: r.start, end: r.end, ext: ".eps", data: append([]byte(nil), img...)})
	}
	return candidates
}

func jpeg2000ImageExt(b []byte) (string, bool) {
	if _, ext, ok := jpeg2000ContainerDeclaredSize(b); ok {
		return ext, true
	}
	if _, ok := jpeg2000CodestreamDeclaredSize(b); ok {
		return ".j2k", true
	}
	return "", false
}

func jpeg2000ContainerDeclaredSizeForExt(ext string, b []byte) (int, bool) {
	size, detectedExt, ok := jpeg2000ContainerDeclaredSize(b)
	if !ok {
		return 0, false
	}
	switch strings.ToLower(ext) {
	case ".jp2":
		if detectedExt != ".jp2" {
			return 0, false
		}
	case ".jpx", ".jpf":
		if detectedExt != ".jpx" && detectedExt != ".jpf" {
			return 0, false
		}
	default:
		return 0, false
	}
	return size, true
}

func jpeg2000ContainerDeclaredSize(b []byte) (int, string, bool) {
	if len(b) < 32 || !bytes.Equal(b[:12], []byte{0, 0, 0, 12, 'j', 'P', ' ', ' ', 0x0d, 0x0a, 0x87, 0x0a}) {
		return 0, "", false
	}
	ftypSize, _, ok := isoBMFFBoxSize(b[12:])
	if !ok || ftypSize < 16 || 12+ftypSize > len(b) || !bytes.Equal(b[16:20], []byte("ftyp")) {
		return 0, "", false
	}
	ext := jpeg2000ExtFromBrands(b[20 : 12+ftypSize])
	if ext == "" {
		return 0, "", false
	}
	off := 12 + ftypSize
	var hasJP2Header, hasCodestream bool
	for off+8 <= len(b) {
		size, header, ok := isoBMFFBoxSize(b[off:])
		if !ok || size < header || size > len(b)-off {
			break
		}
		boxType := string(b[off+4 : off+8])
		switch boxType {
		case "jp2h":
			hasJP2Header = jpeg2000JP2HeaderValid(b[off+header : off+size])
		case "jp2c":
			if codestreamSize, ok := jpeg2000CodestreamDeclaredSize(b[off+header : off+size]); ok && codestreamSize == size-header {
				hasCodestream = true
			}
		}
		off += size
	}
	if !hasJP2Header || !hasCodestream || off <= 12+ftypSize {
		return 0, "", false
	}
	return off, ext, true
}

func jpeg2000JP2HeaderValid(b []byte) bool {
	for off := 0; off+8 <= len(b); {
		size, header, ok := isoBMFFBoxSize(b[off:])
		if !ok || size < header || size > len(b)-off {
			return false
		}
		if string(b[off+4:off+8]) == "ihdr" {
			return jpeg2000ImageHeaderValid(b[off+header : off+size])
		}
		off += size
	}
	return false
}

func jpeg2000ImageHeaderValid(b []byte) bool {
	if len(b) != 14 {
		return false
	}
	height := binary.BigEndian.Uint32(b[0:4])
	width := binary.BigEndian.Uint32(b[4:8])
	components := binary.BigEndian.Uint16(b[8:10])
	bitsPerComponent := b[10]
	compression := b[11]
	unknownColorspace := b[12]
	intellectualProperty := b[13]
	if width == 0 || height == 0 || width > 100000 || height > 100000 || components == 0 {
		return false
	}
	if bitsPerComponent != 0xff && bitsPerComponent&0x7f > 37 {
		return false
	}
	return compression == 7 && unknownColorspace <= 1 && intellectualProperty <= 1
}

func jpeg2000ExtFromBrands(brands []byte) string {
	for off := 0; off+4 <= len(brands); off += 4 {
		switch string(brands[off : off+4]) {
		case "jp2 ":
			return ".jp2"
		case "jpx ", "jpxb", "jpf ":
			return ".jpx"
		}
	}
	return ""
}

func jpeg2000CodestreamDeclaredSize(b []byte) (int, bool) {
	if len(b) < 44 || !bytes.Equal(b[:4], []byte{0xff, 0x4f, 0xff, 0x51}) {
		return 0, false
	}
	segLen := int(binary.BigEndian.Uint16(b[4:6]))
	if segLen < 38 || 4+segLen > len(b) {
		return 0, false
	}
	siz := b[4 : 4+segLen]
	xsiz := binary.BigEndian.Uint32(siz[4:8])
	ysiz := binary.BigEndian.Uint32(siz[8:12])
	xosiz := binary.BigEndian.Uint32(siz[12:16])
	yosiz := binary.BigEndian.Uint32(siz[16:20])
	csiz := int(binary.BigEndian.Uint16(siz[36:38]))
	if xsiz <= xosiz || ysiz <= yosiz || csiz <= 0 || csiz > 16384 || segLen != 38+3*csiz {
		return 0, false
	}
	for off := 38; off+3 <= len(siz); off += 3 {
		ssiz, xrsiz, yrsiz := siz[off], siz[off+1], siz[off+2]
		if ssiz&0x7f > 37 || xrsiz == 0 || yrsiz == 0 {
			return 0, false
		}
	}
	eoc := bytes.Index(b[4+segLen:], []byte{0xff, 0xd9})
	if eoc < 0 {
		return 0, false
	}
	return 4 + segLen + eoc + 2, true
}

func jpeg2000ByteRanges(data []byte) []byteRange {
	var ranges []byteRange
	signature := []byte{0, 0, 0, 12, 'j', 'P', ' ', ' '}
	offset := 0
	for {
		i := bytes.Index(data[offset:], signature)
		if i < 0 {
			break
		}
		start := offset + i
		size, _, ok := jpeg2000ContainerDeclaredSize(data[start:])
		if !ok {
			offset = start + len(signature)
			continue
		}
		ranges = append(ranges, byteRange{start: start, end: start + size})
		offset = start + size
	}
	return ranges
}

func carveJPEG2000Images(data []byte) []imageCandidate {
	var candidates []imageCandidate
	for _, r := range jpeg2000ByteRanges(data) {
		img := data[r.start:r.end]
		ext, ok := jpeg2000ImageExt(img)
		if !ok || len(img) <= 32 {
			continue
		}
		candidates = append(candidates, imageCandidate{start: r.start, end: r.end, ext: ext, data: append([]byte(nil), img...)})
	}
	return candidates
}

func validWebPData(b []byte) bool {
	_, ok := webpDeclaredSize(b)
	return ok
}

func webpDeclaredSize(b []byte) (int, bool) {
	if len(b) < 20 || !bytes.Equal(b[:4], []byte("RIFF")) || !bytes.Equal(b[8:12], []byte("WEBP")) {
		return 0, false
	}
	riffSize := int(binary.LittleEndian.Uint32(b[4:]))
	total := riffSize + 8
	if riffSize < 12 || total < 20 || total > len(b) {
		return 0, false
	}
	if !validWebPChunks(b[12:total]) {
		return 0, false
	}
	return total, true
}

func validWebPChunks(chunks []byte) bool {
	seenImageChunk := false
	seenAnimationFrame := false
	animationFlag := false
	for off := 0; off < len(chunks); {
		if off+8 > len(chunks) {
			return false
		}
		chunkType := string(chunks[off : off+4])
		chunkSize := int(binary.LittleEndian.Uint32(chunks[off+4:]))
		payloadStart := off + 8
		payloadEnd := payloadStart + chunkSize
		if chunkSize < 0 || payloadEnd < payloadStart || payloadEnd > len(chunks) {
			return false
		}
		payload := chunks[payloadStart:payloadEnd]
		switch chunkType {
		case "VP8 ":
			if !validWebPVP8Chunk(payload) {
				return false
			}
			seenImageChunk = true
		case "VP8L":
			if !validWebPVP8LChunk(payload) {
				return false
			}
			seenImageChunk = true
		case "VP8X":
			if !validWebPVP8XChunk(payload) {
				return false
			}
			animationFlag = payload[0]&0x02 != 0
		case "ANMF":
			if !validWebPANMFChunk(payload) {
				return false
			}
			seenAnimationFrame = true
			seenImageChunk = true
		}
		off = payloadEnd
		if chunkSize%2 == 1 {
			if off >= len(chunks) {
				return false
			}
			off++
		}
	}
	return seenImageChunk && (!seenAnimationFrame || animationFlag)
}

func validWebPVP8Chunk(payload []byte) bool {
	if len(payload) < 10 {
		return false
	}
	frameTag := uint32(payload[0]) | uint32(payload[1])<<8 | uint32(payload[2])<<16
	if frameTag&1 != 0 {
		return false
	}
	version := (frameTag >> 1) & 0x7
	if version > 3 || frameTag&0x10 == 0 {
		return false
	}
	firstPartitionLen := int(frameTag >> 5)
	if firstPartitionLen < 7 || firstPartitionLen > len(payload)-3 {
		return false
	}
	if !bytes.Equal(payload[3:6], []byte{0x9d, 0x01, 0x2a}) {
		return false
	}
	width := int(binary.LittleEndian.Uint16(payload[6:])) & 0x3fff
	height := int(binary.LittleEndian.Uint16(payload[8:])) & 0x3fff
	return width > 0 && height > 0
}

func validWebPVP8LChunk(payload []byte) bool {
	if len(payload) < 5 || payload[0] != 0x2f {
		return false
	}
	return payload[4]&0xe0 == 0
}

func validWebPVP8XChunk(payload []byte) bool {
	if len(payload) != 10 {
		return false
	}
	if payload[0]&0xc1 != 0 {
		return false
	}
	width := webp24(payload[4:7]) + 1
	height := webp24(payload[7:10]) + 1
	if width <= 0 || height <= 0 || width > 100000 || height > 100000 {
		return false
	}
	return payload[1] == 0 && payload[2] == 0 && payload[3] == 0
}

func webp24(b []byte) int {
	return int(b[0]) | int(b[1])<<8 | int(b[2])<<16
}

func validWebPANMFChunk(payload []byte) bool {
	if len(payload) < 16 {
		return false
	}
	width := webp24(payload[6:9]) + 1
	height := webp24(payload[9:12]) + 1
	if width <= 0 || height <= 0 || width > 100000 || height > 100000 {
		return false
	}
	if payload[15]&0xfc != 0 {
		return false
	}
	return validWebPChunks(payload[16:])
}

func validICOData(b []byte) bool {
	_, ok := icoDeclaredSize(b)
	return ok
}

func icoDeclaredSize(b []byte) (int, bool) {
	return iconDirectoryDeclaredSize(b, 1)
}

func validCURData(b []byte) bool {
	_, ok := curDeclaredSize(b)
	return ok
}

func curDeclaredSize(b []byte) (int, bool) {
	return iconDirectoryDeclaredSize(b, 2)
}

func iconDirectoryDeclaredSize(b []byte, wantType uint16) (int, bool) {
	if len(b) < 22 || binary.LittleEndian.Uint16(b[0:]) != 0 || binary.LittleEndian.Uint16(b[2:]) != wantType {
		return 0, false
	}
	count := int(binary.LittleEndian.Uint16(b[4:]))
	if count <= 0 || count > 256 {
		return 0, false
	}
	dirEnd := 6 + count*16
	if dirEnd > len(b) {
		return 0, false
	}
	maxEnd := dirEnd
	ranges := make([]intRange, 0, count)
	for i := 0; i < count; i++ {
		entry := b[6+i*16:]
		if entry[3] != 0 {
			return 0, false
		}
		size := int(binary.LittleEndian.Uint32(entry[8:]))
		offset := int(binary.LittleEndian.Uint32(entry[12:]))
		if size <= 0 || offset < dirEnd || offset > len(b) || size > len(b)-offset {
			return 0, false
		}
		end := offset + size
		for _, r := range ranges {
			if offset < r.max && end > r.min {
				return 0, false
			}
		}
		ranges = append(ranges, intRange{min: offset, max: end})
		payload := b[offset : offset+size]
		if !validICOImagePayload(payload) {
			return 0, false
		}
		if w, h, ok := icoImagePayloadDimensions(payload); ok {
			if !iconEntryMatchesPayloadDimensions(entry[0], entry[1], w, h, icoPayloadIsDIB(payload)) {
				return 0, false
			}
		}
		if end > maxEnd {
			maxEnd = end
		}
	}
	return maxEnd, true
}

func iconEntryDimension(v byte) int {
	if v == 0 {
		return 256
	}
	return int(v)
}

func iconEntryMatchesPayloadDimensions(entryWidth, entryHeight byte, payloadWidth, payloadHeight int, dibPayload bool) bool {
	w := iconEntryDimension(entryWidth)
	h := iconEntryDimension(entryHeight)
	if w != payloadWidth {
		return false
	}
	if h == payloadHeight {
		return true
	}
	return dibPayload && payloadHeight == h*2
}

func icoImagePayloadDimensions(b []byte) (int, int, bool) {
	if len(b) >= 8 && bytes.Equal(b[:8], []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1a, '\n'}) {
		size, ok := pngEndOffset(b)
		if !ok || size != len(b) {
			return 0, 0, false
		}
		cfg, _, err := image.DecodeConfig(bytes.NewReader(b))
		if err != nil || cfg.Width <= 0 || cfg.Height <= 0 {
			return 0, 0, false
		}
		return cfg.Width, cfg.Height, true
	}
	w, h, ok := dibDimensions(b)
	return w, h, ok
}

func icoPayloadIsDIB(b []byte) bool {
	if len(b) >= 8 && bytes.Equal(b[:8], []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1a, '\n'}) {
		return false
	}
	return true
}

func validICOImagePayload(b []byte) bool {
	if len(b) >= 8 && bytes.Equal(b[:8], []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1a, '\n'}) {
		size, ok := pngEndOffset(b)
		return ok && size == len(b)
	}
	_, _, ok := dibDeclaredSize(b)
	return ok
}

func validDIBData(b []byte) bool {
	_, _, ok := dibDeclaredSize(b)
	return ok
}

func validPCXData(b []byte) bool {
	_, ok := pcxDeclaredSize(b)
	return ok
}

func validTGAData(b []byte) bool {
	_, ok := tgaDeclaredSize(b)
	return ok
}

func pcxDeclaredSize(b []byte) (int, bool) {
	if len(b) < 128 || b[0] != 0x0a || b[2] != 1 {
		return 0, false
	}
	switch b[1] {
	case 0, 2, 3, 4, 5:
	default:
		return 0, false
	}
	bitsPerPixel := int(b[3])
	switch bitsPerPixel {
	case 1, 2, 4, 8:
	default:
		return 0, false
	}
	xmin := int(binary.LittleEndian.Uint16(b[4:]))
	ymin := int(binary.LittleEndian.Uint16(b[6:]))
	xmax := int(binary.LittleEndian.Uint16(b[8:]))
	ymax := int(binary.LittleEndian.Uint16(b[10:]))
	if xmax < xmin || ymax < ymin {
		return 0, false
	}
	width := xmax - xmin + 1
	height := ymax - ymin + 1
	if width <= 0 || height <= 0 || width > 100000 || height > 100000 {
		return 0, false
	}
	if b[64] != 0 {
		return 0, false
	}
	planes := int(b[65])
	switch planes {
	case 1, 3, 4:
	default:
		return 0, false
	}
	bytesPerLine := int(binary.LittleEndian.Uint16(b[66:]))
	minBytesPerLine := (width*bitsPerPixel + 7) / 8
	if bytesPerLine < minBytesPerLine || bytesPerLine > 65535 {
		return 0, false
	}
	paletteInfo := binary.LittleEndian.Uint16(b[68:])
	if paletteInfo > 2 {
		return 0, false
	}
	totalDecoded := int64(height) * int64(planes) * int64(bytesPerLine)
	if totalDecoded <= 0 || totalDecoded > int64(maxCompressedMetafileBytes) {
		return 0, false
	}
	off := 128
	var decoded int64
	for decoded < totalDecoded {
		if off >= len(b) {
			return 0, false
		}
		c := b[off]
		off++
		count := 1
		if c&0xc0 == 0xc0 {
			count = int(c & 0x3f)
			if count == 0 || off >= len(b) {
				return 0, false
			}
			off++
		}
		if decoded+int64(count) > totalDecoded {
			return 0, false
		}
		decoded += int64(count)
	}
	if b[1] == 5 && bitsPerPixel == 8 && planes == 1 && off+769 <= len(b) && b[off] == 0x0c {
		off += 769
	}
	return off, off > 128
}

func tgaDeclaredSize(b []byte) (int, bool) {
	if len(b) < 18 {
		return 0, false
	}
	idLen := int(b[0])
	colorMapType := b[1]
	imageType := int(b[2])
	colorMapFirstIndex := int(binary.LittleEndian.Uint16(b[3:]))
	colorMapLength := int(binary.LittleEndian.Uint16(b[5:]))
	colorMapEntryBits := int(b[7])
	width := int(binary.LittleEndian.Uint16(b[12:]))
	height := int(binary.LittleEndian.Uint16(b[14:]))
	pixelBits := int(b[16])
	descriptor := b[17]
	if width <= 0 || height <= 0 || width > 100000 || height > 100000 {
		return 0, false
	}
	if colorMapType != 0 && colorMapType != 1 {
		return 0, false
	}
	if descriptor&0xc0 != 0 {
		return 0, false
	}
	offset := 18 + idLen
	if offset < 18 || offset > len(b) {
		return 0, false
	}
	hasColorMap := colorMapType == 1
	switch imageType {
	case 1, 9:
		if !hasColorMap {
			return 0, false
		}
	case 2, 3, 10, 11:
		if colorMapType != 0 {
			return 0, false
		}
		if colorMapFirstIndex != 0 || colorMapLength != 0 || colorMapEntryBits != 0 {
			return 0, false
		}
	default:
		return 0, false
	}
	if hasColorMap {
		if colorMapFirstIndex+colorMapLength > 65536 {
			return 0, false
		}
		if colorMapEntryBits != 15 && colorMapEntryBits != 16 && colorMapEntryBits != 24 && colorMapEntryBits != 32 {
			return 0, false
		}
		entryBytes := (colorMapEntryBits + 7) / 8
		colorMapBytes := colorMapLength * entryBytes
		if colorMapLength <= 0 || colorMapBytes/colorMapLength != entryBytes || offset+colorMapBytes < offset || offset+colorMapBytes > len(b) {
			return 0, false
		}
		offset += colorMapBytes
	}
	pixelBytes, ok := tgaPixelBytes(imageType, pixelBits)
	if !ok {
		return 0, false
	}
	pixels := int64(width) * int64(height)
	if pixels <= 0 || pixels > int64(maxCompressedMetafileBytes) {
		return 0, false
	}
	switch imageType {
	case 1, 2, 3:
		payload := pixels * int64(pixelBytes)
		if payload <= 0 || payload > int64(len(b)-offset) {
			return 0, false
		}
		return offset + int(payload), true
	case 9, 10, 11:
		return tgaRLEEndOffset(b, offset, pixels, pixelBytes)
	default:
		return 0, false
	}
}

func tgaPixelBytes(imageType, pixelBits int) (int, bool) {
	switch imageType {
	case 1, 9:
		switch pixelBits {
		case 8, 16:
			return pixelBits / 8, true
		default:
			return 0, false
		}
	case 2, 10:
		switch pixelBits {
		case 16, 24, 32:
			return pixelBits / 8, true
		default:
			return 0, false
		}
	case 3, 11:
		switch pixelBits {
		case 8, 16:
			return pixelBits / 8, true
		default:
			return 0, false
		}
	default:
		return 0, false
	}
}

func tgaRLEEndOffset(b []byte, offset int, pixels int64, pixelBytes int) (int, bool) {
	remaining := pixels
	for remaining > 0 {
		if offset >= len(b) {
			return 0, false
		}
		header := b[offset]
		offset++
		count := int64(header&0x7f) + 1
		if count > remaining {
			return 0, false
		}
		bytesNeeded := pixelBytes
		if header&0x80 == 0 {
			bytesNeeded = int(count) * pixelBytes
		}
		if bytesNeeded <= 0 || bytesNeeded > len(b)-offset {
			return 0, false
		}
		offset += bytesNeeded
		remaining -= count
	}
	return offset, true
}

func validPICTData(b []byte) bool {
	_, ok := pictDeclaredSize(b)
	return ok
}

func pictDeclaredSize(b []byte) (int, bool) {
	if len(b) >= 512 {
		if size, ok := rawPICTDeclaredSize(b[512:]); ok {
			return 512 + size, true
		}
	}
	return rawPICTDeclaredSize(b)
}

func rawPICTDeclaredSize(b []byte) (int, bool) {
	// A PICT v2 stream starts with its frame followed by the version opcode and
	// the mandatory extended-header opcode/version.  Merely finding the version
	// opcode is too weak: BIFF/Escher data can contain that byte sequence by
	// chance (notably in otherwise text-only .xls files).
	if len(b) < 40 || !bytes.Equal(b[10:14], []byte{0x00, 0x11, 0x02, 0xff}) ||
		!bytes.Equal(b[14:18], []byte{0x0c, 0x00, 0xff, 0xfe}) {
		return 0, false
	}
	top := int(int16(binary.BigEndian.Uint16(b[2:])))
	left := int(int16(binary.BigEndian.Uint16(b[4:])))
	bottom := int(int16(binary.BigEndian.Uint16(b[6:])))
	right := int(int16(binary.BigEndian.Uint16(b[8:])))
	if bottom <= top || right <= left || bottom-top > 100000 || right-left > 100000 {
		return 0, false
	}
	for off := 14; off+2 <= len(b); off += 2 {
		if bytes.Equal(b[off:off+2], []byte{0x00, 0xff}) {
			return off + 2, off > 14
		}
	}
	return 0, false
}

func pictByteRanges(data []byte) []byteRange {
	var ranges []byteRange
	signature := []byte{0x00, 0x11, 0x02, 0xff}
	offset := 0
	for {
		i := bytes.Index(data[offset:], signature)
		if i < 0 {
			break
		}
		sig := offset + i
		for _, start := range []int{sig - 522, sig - 10} {
			if start < 0 {
				continue
			}
			size, ok := pictDeclaredSize(data[start:])
			if ok && start+size <= len(data) {
				ranges = append(ranges, byteRange{start: start, end: start + size})
			}
		}
		offset = sig + len(signature)
	}
	return ranges
}

func carvePICTImages(data []byte) []imageCandidate {
	var candidates []imageCandidate
	for _, r := range pictByteRanges(data) {
		img := data[r.start:r.end]
		if len(img) <= 16 || !validPICTData(img) {
			continue
		}
		candidates = append(candidates, imageCandidate{start: r.start, end: r.end, ext: ".pict", data: append([]byte(nil), img...)})
	}
	return candidates
}

func isoBMFFImageExt(b []byte) (string, bool) {
	_, ext, ok := isoBMFFDeclaredSize(b)
	return ext, ok
}

func isoBMFFDeclaredSizeForExt(ext string, b []byte) (int, bool) {
	size, detectedExt, ok := isoBMFFDeclaredSize(b)
	if !ok {
		return 0, false
	}
	switch strings.ToLower(ext) {
	case ".avif":
		if detectedExt != ".avif" {
			return 0, false
		}
	case ".heic", ".heif":
		if detectedExt != ".heic" && detectedExt != ".heif" {
			return 0, false
		}
	default:
		return 0, false
	}
	return size, true
}

func isoBMFFDeclaredSize(b []byte) (int, string, bool) {
	if len(b) < 24 || !bytes.Equal(b[4:8], []byte("ftyp")) {
		return 0, "", false
	}
	ftypSize := int(binary.BigEndian.Uint32(b[0:4]))
	if ftypSize < 16 || ftypSize%4 != 0 || ftypSize > len(b) {
		return 0, "", false
	}
	ext := isoBMFFImageExtFromBrands(b[8:ftypSize])
	if ext == "" {
		return 0, "", false
	}
	off := ftypSize
	var mediaBoxes int
	var hasMetaPayload bool
	for off+8 <= len(b) {
		size, header, ok := isoBMFFBoxSize(b[off:])
		if !ok || size < header || size > len(b)-off {
			break
		}
		boxType := string(b[off+4 : off+8])
		if boxType == "meta" || boxType == "mdat" || boxType == "moov" {
			mediaBoxes++
		}
		if boxType == "meta" && size >= header+12 {
			hasMetaPayload = true
		}
		off += size
	}
	if !hasMetaPayload || mediaBoxes == 0 || off <= ftypSize {
		return 0, "", false
	}
	return off, ext, true
}

func isoBMFFBoxSize(b []byte) (int, int, bool) {
	if len(b) < 8 {
		return 0, 0, false
	}
	size32 := binary.BigEndian.Uint32(b[:4])
	switch size32 {
	case 0:
		return 0, 0, false
	case 1:
		if len(b) < 16 {
			return 0, 0, false
		}
		size64 := binary.BigEndian.Uint64(b[8:16])
		if size64 > uint64(len(b)) || size64 > uint64(int(^uint(0)>>1)) {
			return 0, 0, false
		}
		return int(size64), 16, true
	default:
		return int(size32), 8, true
	}
}

func isoBMFFImageExtFromBrands(brands []byte) string {
	if len(brands) < 8 {
		return ""
	}
	var sawHEIF bool
	for off := 0; off+4 <= len(brands); off += 4 {
		brand := string(brands[off : off+4])
		switch brand {
		case "avif", "avis":
			return ".avif"
		case "heic", "heix", "hevc", "hevx":
			return ".heic"
		case "mif1", "msf1":
			sawHEIF = true
		}
	}
	if sawHEIF {
		return ".heif"
	}
	return ""
}

func isoBMFFByteRanges(data []byte) []byteRange {
	var ranges []byteRange
	offset := 0
	for {
		i := bytes.Index(data[offset:], []byte("ftyp"))
		if i < 0 {
			break
		}
		ftyp := offset + i
		if ftyp < 4 {
			offset = ftyp + 4
			continue
		}
		start := ftyp - 4
		size, _, ok := isoBMFFDeclaredSize(data[start:])
		if !ok {
			offset = ftyp + 4
			continue
		}
		ranges = append(ranges, byteRange{start: start, end: start + size})
		offset = start + size
	}
	return ranges
}

func carveISOImages(data []byte) []imageCandidate {
	var candidates []imageCandidate
	for _, r := range isoBMFFByteRanges(data) {
		img := data[r.start:r.end]
		ext, ok := isoBMFFImageExt(img)
		if !ok || len(img) <= 32 {
			continue
		}
		candidates = append(candidates, imageCandidate{start: r.start, end: r.end, ext: ext, data: append([]byte(nil), img...)})
	}
	return candidates
}

func validTIFFData(b []byte) bool {
	_, ok := tiffDeclaredSize(b)
	return ok
}

func validJPEGXRData(b []byte) bool {
	_, ok := jpegXRDeclaredSize(b)
	return ok
}

func jpegXRDeclaredSize(b []byte) (int, bool) {
	if len(b) < 8 || !bytes.Equal(b[:2], []byte("II")) || binary.LittleEndian.Uint16(b[2:]) != 0x01bc {
		return 0, false
	}
	offset := int(binary.LittleEndian.Uint32(b[4:]))
	maxEnd := 8
	sawDimensions := false
	sawImagePayload := false
	for ifd := 0; ifd < 16 && offset != 0; ifd++ {
		end, next, hasDimensions, hasPayload, ok := tiffIFDDeclaredEnd(b, binary.LittleEndian, offset)
		if !ok {
			return 0, false
		}
		if end > maxEnd {
			maxEnd = end
		}
		sawDimensions = sawDimensions || hasDimensions
		sawImagePayload = sawImagePayload || hasPayload
		if next < 0 || next > len(b) || (next != 0 && next <= offset) {
			return 0, false
		}
		offset = next
	}
	if offset != 0 || !sawDimensions || !sawImagePayload || maxEnd > len(b) {
		return 0, false
	}
	return maxEnd, true
}

func tiffDeclaredSize(b []byte) (int, bool) {
	if len(b) < 8 {
		return 0, false
	}
	var order binary.ByteOrder
	switch {
	case bytes.Equal(b[:2], []byte("II")):
		order = binary.LittleEndian
	case bytes.Equal(b[:2], []byte("MM")):
		order = binary.BigEndian
	default:
		return 0, false
	}
	magic := order.Uint16(b[2:])
	if magic == 43 {
		return bigTIFFDeclaredSize(b, order)
	}
	if magic != 42 {
		return 0, false
	}
	offset := int(order.Uint32(b[4:]))
	maxEnd := 8
	sawDimensions := false
	for ifd := 0; ifd < 16 && offset != 0; ifd++ {
		end, next, hasDimensions, _, ok := tiffIFDDeclaredEnd(b, order, offset)
		if !ok {
			return 0, false
		}
		if end > maxEnd {
			maxEnd = end
		}
		sawDimensions = sawDimensions || hasDimensions
		if next < 0 || next > len(b) || (next != 0 && next <= offset) {
			return 0, false
		}
		offset = next
	}
	if offset != 0 || !sawDimensions || maxEnd > len(b) {
		return 0, false
	}
	return maxEnd, true
}

func bigTIFFDeclaredSize(b []byte, order binary.ByteOrder) (int, bool) {
	if len(b) < 16 || order.Uint16(b[4:]) != 8 || order.Uint16(b[6:]) != 0 {
		return 0, false
	}
	offset64 := order.Uint64(b[8:])
	if offset64 > uint64(len(b)) || offset64 > uint64(int(^uint(0)>>1)) {
		return 0, false
	}
	offset := int(offset64)
	maxEnd := 16
	sawDimensions := false
	for ifd := 0; ifd < 16 && offset != 0; ifd++ {
		end, next, hasDimensions, _, ok := bigTIFFIFDDeclaredEnd(b, order, offset)
		if !ok {
			return 0, false
		}
		if end > maxEnd {
			maxEnd = end
		}
		sawDimensions = sawDimensions || hasDimensions
		if next < 0 || next > len(b) || (next != 0 && next <= offset) {
			return 0, false
		}
		offset = next
	}
	if offset != 0 || !sawDimensions || maxEnd > len(b) {
		return 0, false
	}
	return maxEnd, true
}

func tiffIFDDeclaredEnd(b []byte, order binary.ByteOrder, offset int) (int, int, bool, bool, bool) {
	if offset < 8 || offset+2 > len(b) {
		return 0, 0, false, false, false
	}
	entries := int(order.Uint16(b[offset:]))
	if entries == 0 || entries > (len(b)-offset-6)/12 {
		return 0, 0, false, false, false
	}
	ifdEnd := offset + 2 + entries*12 + 4
	if ifdEnd > len(b) {
		return 0, 0, false, false, false
	}
	maxEnd := ifdEnd
	hasWidth, hasHeight := false, false
	var stripOffsets, stripByteCounts []uint64
	var tileOffsets, tileByteCounts []uint64
	var jpegOffset, jpegLength uint64
	for i := 0; i < entries; i++ {
		pos := offset + 2 + i*12
		tag := order.Uint16(b[pos:])
		fieldType := order.Uint16(b[pos+2:])
		count := order.Uint32(b[pos+4:])
		typeSize, ok := tiffTypeSize(fieldType)
		if !ok || count == 0 || uint64(count) > uint64(len(b))/uint64(typeSize) {
			return 0, 0, false, false, false
		}
		fieldBytes := uint64(count) * uint64(typeSize)
		if fieldBytes > 4 {
			valueOffset := uint64(order.Uint32(b[pos+8:]))
			if valueOffset > uint64(len(b)) || fieldBytes > uint64(len(b))-valueOffset {
				return 0, 0, false, false, false
			}
			if end := int(valueOffset + fieldBytes); end > maxEnd {
				maxEnd = end
			}
		}
		switch tag {
		case 256:
			hasWidth = hasWidth || tiffDimensionTagIsValid(b, order, pos)
		case 257:
			hasHeight = hasHeight || tiffDimensionTagIsValid(b, order, pos)
		case 273:
			stripOffsets = tiffNumericValues(b, order, pos)
		case 279:
			stripByteCounts = tiffNumericValues(b, order, pos)
		case 324:
			tileOffsets = tiffNumericValues(b, order, pos)
		case 325:
			tileByteCounts = tiffNumericValues(b, order, pos)
		case 513:
			values := tiffNumericValues(b, order, pos)
			if len(values) > 0 {
				jpegOffset = values[0]
			}
		case 514:
			values := tiffNumericValues(b, order, pos)
			if len(values) > 0 {
				jpegLength = values[0]
			}
		}
	}
	if len(stripOffsets) != len(stripByteCounts) || len(tileOffsets) != len(tileByteCounts) {
		return 0, 0, false, false, false
	}
	hasImagePayload := false
	for i, off := range stripOffsets {
		byteCount := stripByteCounts[i]
		if off == 0 || byteCount == 0 || off > uint64(len(b)) || byteCount > uint64(len(b))-off {
			return 0, 0, false, false, false
		}
		hasImagePayload = true
		end := off + byteCount
		if int(end) > maxEnd {
			maxEnd = int(end)
		}
	}
	for i, off := range tileOffsets {
		byteCount := tileByteCounts[i]
		if off == 0 || byteCount == 0 || off > uint64(len(b)) || byteCount > uint64(len(b))-off {
			return 0, 0, false, false, false
		}
		hasImagePayload = true
		end := off + byteCount
		if int(end) > maxEnd {
			maxEnd = int(end)
		}
	}
	if (jpegOffset == 0) != (jpegLength == 0) {
		return 0, 0, false, false, false
	}
	if jpegOffset != 0 {
		end := jpegOffset + jpegLength
		if jpegOffset > uint64(len(b)) || jpegLength > uint64(len(b))-jpegOffset {
			return 0, 0, false, false, false
		}
		hasImagePayload = true
		if int(end) > maxEnd {
			maxEnd = int(end)
		}
	}
	next := int(order.Uint32(b[offset+2+entries*12:]))
	return maxEnd, next, hasWidth && hasHeight, hasImagePayload, true
}

func bigTIFFIFDDeclaredEnd(b []byte, order binary.ByteOrder, offset int) (int, int, bool, bool, bool) {
	if offset < 16 || offset+8 > len(b) {
		return 0, 0, false, false, false
	}
	entries64 := order.Uint64(b[offset:])
	if entries64 == 0 || entries64 > uint64((len(b)-offset-16)/20) || entries64 > uint64(int(^uint(0)>>1)) {
		return 0, 0, false, false, false
	}
	entries := int(entries64)
	ifdEnd := offset + 8 + entries*20 + 8
	if ifdEnd > len(b) {
		return 0, 0, false, false, false
	}
	maxEnd := ifdEnd
	hasWidth, hasHeight := false, false
	var stripOffsets, stripByteCounts []uint64
	var tileOffsets, tileByteCounts []uint64
	var jpegOffset, jpegLength uint64
	for i := 0; i < entries; i++ {
		pos := offset + 8 + i*20
		tag := order.Uint16(b[pos:])
		fieldType := order.Uint16(b[pos+2:])
		count := order.Uint64(b[pos+4:])
		typeSize, ok := tiffTypeSize(fieldType)
		if !ok || count == 0 || count > uint64(len(b))/uint64(typeSize) {
			return 0, 0, false, false, false
		}
		fieldBytes := count * uint64(typeSize)
		if fieldBytes > 8 {
			valueOffset := order.Uint64(b[pos+12:])
			if valueOffset > uint64(len(b)) || fieldBytes > uint64(len(b))-valueOffset {
				return 0, 0, false, false, false
			}
			if end := int(valueOffset + fieldBytes); end > maxEnd {
				maxEnd = end
			}
		}
		switch tag {
		case 256:
			hasWidth = hasWidth || bigTIFFDimensionTagIsValid(b, order, pos)
		case 257:
			hasHeight = hasHeight || bigTIFFDimensionTagIsValid(b, order, pos)
		case 273:
			stripOffsets = bigTIFFNumericValues(b, order, pos)
		case 279:
			stripByteCounts = bigTIFFNumericValues(b, order, pos)
		case 324:
			tileOffsets = bigTIFFNumericValues(b, order, pos)
		case 325:
			tileByteCounts = bigTIFFNumericValues(b, order, pos)
		case 513:
			values := bigTIFFNumericValues(b, order, pos)
			if len(values) > 0 {
				jpegOffset = values[0]
			}
		case 514:
			values := bigTIFFNumericValues(b, order, pos)
			if len(values) > 0 {
				jpegLength = values[0]
			}
		}
	}
	if len(stripOffsets) != len(stripByteCounts) || len(tileOffsets) != len(tileByteCounts) {
		return 0, 0, false, false, false
	}
	hasImagePayload := false
	for i, off := range stripOffsets {
		byteCount := stripByteCounts[i]
		if off == 0 || byteCount == 0 || off > uint64(len(b)) || byteCount > uint64(len(b))-off {
			return 0, 0, false, false, false
		}
		hasImagePayload = true
		end := off + byteCount
		if int(end) > maxEnd {
			maxEnd = int(end)
		}
	}
	for i, off := range tileOffsets {
		byteCount := tileByteCounts[i]
		if off == 0 || byteCount == 0 || off > uint64(len(b)) || byteCount > uint64(len(b))-off {
			return 0, 0, false, false, false
		}
		hasImagePayload = true
		end := off + byteCount
		if int(end) > maxEnd {
			maxEnd = int(end)
		}
	}
	if (jpegOffset == 0) != (jpegLength == 0) {
		return 0, 0, false, false, false
	}
	if jpegOffset != 0 {
		end := jpegOffset + jpegLength
		if jpegOffset > uint64(len(b)) || jpegLength > uint64(len(b))-jpegOffset {
			return 0, 0, false, false, false
		}
		hasImagePayload = true
		if int(end) > maxEnd {
			maxEnd = int(end)
		}
	}
	next64 := order.Uint64(b[offset+8+entries*20:])
	if next64 > uint64(len(b)) || next64 > uint64(int(^uint(0)>>1)) {
		return 0, 0, false, false, false
	}
	return maxEnd, int(next64), hasWidth && hasHeight, hasImagePayload, true
}

func tiffNumericValues(b []byte, order binary.ByteOrder, entry int) []uint64 {
	fieldType := order.Uint16(b[entry+2:])
	count := int(order.Uint32(b[entry+4:]))
	typeSize, ok := tiffTypeSize(fieldType)
	if !ok || count <= 0 {
		return nil
	}
	fieldBytes := count * typeSize
	value := b[entry+8 : entry+12]
	if fieldBytes > 4 {
		offset := int(order.Uint32(value))
		if offset < 0 || offset+fieldBytes > len(b) {
			return nil
		}
		value = b[offset : offset+fieldBytes]
	}
	out := make([]uint64, 0, count)
	for i := 0; i < count; i++ {
		pos := i * typeSize
		switch fieldType {
		case 3:
			out = append(out, uint64(order.Uint16(value[pos:])))
		case 4:
			out = append(out, uint64(order.Uint32(value[pos:])))
		}
	}
	return out
}

func bigTIFFNumericValues(b []byte, order binary.ByteOrder, entry int) []uint64 {
	fieldType := order.Uint16(b[entry+2:])
	count64 := order.Uint64(b[entry+4:])
	typeSize, ok := tiffTypeSize(fieldType)
	if !ok || count64 == 0 || count64 > uint64(int(^uint(0)>>1)) {
		return nil
	}
	count := int(count64)
	fieldBytes64 := count64 * uint64(typeSize)
	if fieldBytes64 > uint64(int(^uint(0)>>1)) {
		return nil
	}
	fieldBytes := int(fieldBytes64)
	value := b[entry+12 : entry+20]
	if fieldBytes > 8 {
		offset64 := order.Uint64(value)
		if offset64 > uint64(len(b)) || fieldBytes64 > uint64(len(b))-offset64 || offset64 > uint64(int(^uint(0)>>1)) {
			return nil
		}
		offset := int(offset64)
		value = b[offset : offset+fieldBytes]
	}
	out := make([]uint64, 0, count)
	for i := 0; i < count; i++ {
		pos := i * typeSize
		switch fieldType {
		case 3:
			out = append(out, uint64(order.Uint16(value[pos:])))
		case 4:
			out = append(out, uint64(order.Uint32(value[pos:])))
		case 16, 17, 18:
			out = append(out, order.Uint64(value[pos:]))
		}
	}
	return out
}

func validTIFFIFD(b []byte, order binary.ByteOrder, offset int) bool {
	if offset < 8 || offset+2 > len(b) {
		return false
	}
	entries := int(order.Uint16(b[offset:]))
	if entries == 0 || entries > (len(b)-offset-6)/12 {
		return false
	}
	hasWidth, hasHeight := false, false
	for i := 0; i < entries; i++ {
		pos := offset + 2 + i*12
		fieldType := order.Uint16(b[pos+2:])
		count := order.Uint32(b[pos+4:])
		typeSize, ok := tiffTypeSize(fieldType)
		if !ok || count == 0 || uint64(count) > uint64(len(b))/uint64(typeSize) {
			return false
		}
		fieldBytes := uint64(count) * uint64(typeSize)
		if fieldBytes > 4 {
			valueOffset := uint64(order.Uint32(b[pos+8:]))
			if valueOffset > uint64(len(b)) || fieldBytes > uint64(len(b))-valueOffset {
				return false
			}
		}
		tag := order.Uint16(b[pos:])
		switch tag {
		case 256:
			hasWidth = hasWidth || tiffDimensionTagIsValid(b, order, pos)
		case 257:
			hasHeight = hasHeight || tiffDimensionTagIsValid(b, order, pos)
		}
	}
	return hasWidth && hasHeight
}

func tiffTypeSize(fieldType uint16) (int, bool) {
	switch fieldType {
	case 1, 2, 6, 7:
		return 1, true
	case 3, 8:
		return 2, true
	case 4, 9, 11:
		return 4, true
	case 5, 10, 12:
		return 8, true
	case 16, 17, 18:
		return 8, true
	default:
		return 0, false
	}
}

func tiffDimensionTagIsValid(b []byte, order binary.ByteOrder, entry int) bool {
	fieldType := order.Uint16(b[entry+2:])
	count := order.Uint32(b[entry+4:])
	if count == 0 {
		return false
	}
	switch fieldType {
	case 3:
		if count == 1 {
			return order.Uint16(b[entry+8:]) > 0
		}
		offset := int(order.Uint32(b[entry+8:]))
		return offset >= 0 && offset+2 <= len(b) && order.Uint16(b[offset:]) > 0
	case 4:
		if count == 1 {
			return order.Uint32(b[entry+8:]) > 0
		}
		offset := int(order.Uint32(b[entry+8:]))
		return offset >= 0 && offset+4 <= len(b) && order.Uint32(b[offset:]) > 0
	default:
		return false
	}
}

func bigTIFFDimensionTagIsValid(b []byte, order binary.ByteOrder, entry int) bool {
	fieldType := order.Uint16(b[entry+2:])
	count := order.Uint64(b[entry+4:])
	if count == 0 {
		return false
	}
	switch fieldType {
	case 3:
		if count == 1 {
			return order.Uint16(b[entry+12:]) > 0
		}
		offset := order.Uint64(b[entry+12:])
		return offset <= uint64(len(b)) && 2 <= uint64(len(b))-offset && order.Uint16(b[int(offset):]) > 0
	case 4:
		if count == 1 {
			return order.Uint32(b[entry+12:]) > 0
		}
		offset := order.Uint64(b[entry+12:])
		return offset <= uint64(len(b)) && 4 <= uint64(len(b))-offset && order.Uint32(b[int(offset):]) > 0
	case 16, 17, 18:
		if count == 1 {
			return order.Uint64(b[entry+12:]) > 0
		}
		offset := order.Uint64(b[entry+12:])
		return offset <= uint64(len(b)) && 8 <= uint64(len(b))-offset && order.Uint64(b[int(offset):]) > 0
	default:
		return false
	}
}

func validBMPData(b []byte) bool {
	_, ok := bmpDeclaredSize(b)
	return ok
}

func dibToBMP(b []byte) ([]byte, bool) {
	size, pixelOffset, ok := dibDeclaredSize(b)
	if !ok {
		return nil, false
	}
	fileSize := 14 + size
	out := make([]byte, fileSize)
	copy(out[14:], b[:size])
	copy(out[:2], []byte("BM"))
	binary.LittleEndian.PutUint32(out[2:], uint32(fileSize))
	binary.LittleEndian.PutUint32(out[10:], uint32(14+pixelOffset))
	return out, true
}

func dibDeclaredSize(b []byte) (int, int, bool) {
	minPixelOffset, pixelBytes, ok := dibImageLayout(b)
	if !ok {
		return 0, 0, false
	}
	total := minPixelOffset + pixelBytes
	if total < minPixelOffset || total > len(b) {
		return 0, 0, false
	}
	return total, minPixelOffset, true
}

func dibDimensions(b []byte) (int, int, bool) {
	if len(b) < 12 {
		return 0, 0, false
	}
	headerSize := int(binary.LittleEndian.Uint32(b))
	switch headerSize {
	case 12:
		width := int(binary.LittleEndian.Uint16(b[4:]))
		height := int(binary.LittleEndian.Uint16(b[6:]))
		if width <= 0 || height <= 0 || width > 100000 || height > 100000 {
			return 0, 0, false
		}
		return width, height, true
	case 40, 108, 124:
		if len(b) < headerSize {
			return 0, 0, false
		}
		width := int(int32(binary.LittleEndian.Uint32(b[4:])))
		height := int(int32(binary.LittleEndian.Uint32(b[8:])))
		if width <= 0 || height == 0 {
			return 0, 0, false
		}
		if height < 0 {
			height = -height
		}
		return width, height, true
	default:
		return 0, 0, false
	}
}

func dibImageLayout(b []byte) (int, int, bool) {
	if len(b) < 12 {
		return 0, 0, false
	}
	headerSize := int(binary.LittleEndian.Uint32(b))
	switch headerSize {
	case 12:
		width := int(binary.LittleEndian.Uint16(b[4:]))
		height := int(binary.LittleEndian.Uint16(b[6:]))
		if width <= 0 || height <= 0 || width > 100000 || height > 100000 {
			return 0, 0, false
		}
		planes := binary.LittleEndian.Uint16(b[8:])
		bitCount := binary.LittleEndian.Uint16(b[10:])
		if planes != 1 {
			return 0, 0, false
		}
		switch bitCount {
		case 1, 4, 8, 24:
		default:
			return 0, 0, false
		}
		paletteEntries := 0
		if bitCount <= 8 {
			paletteEntries = 1 << bitCount
		}
		pixelOffset := headerSize + paletteEntries*3
		if pixelOffset < headerSize || pixelOffset > len(b) {
			return 0, 0, false
		}
		rowBits := int64(width) * int64(bitCount)
		stride := ((rowBits + 31) / 32) * 4
		if stride <= 0 || stride > int64(len(b)) {
			return 0, 0, false
		}
		pixelBytes := int(stride * int64(height))
		return pixelOffset, pixelBytes, true
	case 40, 108, 124:
	default:
		return 0, 0, false
	}
	if len(b) < headerSize {
		return 0, 0, false
	}
	width := int(int32(binary.LittleEndian.Uint32(b[4:])))
	height := int(int32(binary.LittleEndian.Uint32(b[8:])))
	if width <= 0 || height == 0 || width > 100000 || height > 100000 || height < -100000 {
		return 0, 0, false
	}
	planes := binary.LittleEndian.Uint16(b[12:])
	bitCount := binary.LittleEndian.Uint16(b[14:])
	compression := binary.LittleEndian.Uint32(b[16:])
	if planes != 1 {
		return 0, 0, false
	}
	switch bitCount {
	case 1, 4, 8, 16, 24, 32:
	default:
		return 0, 0, false
	}
	maskBytes := 0
	switch compression {
	case 0:
	case 3, 6:
		if bitCount != 16 && bitCount != 32 {
			return 0, 0, false
		}
		if headerSize == 40 {
			maskBytes = 12
			if compression == 6 {
				maskBytes = 16
			}
			if len(b) < headerSize+maskBytes || !validDIBBitfieldMasks(b[headerSize:headerSize+maskBytes], bitCount, compression == 6) {
				return 0, 0, false
			}
		} else {
			if !validDIBBitfieldMasks(b[40:56], bitCount, compression == 6) {
				return 0, 0, false
			}
		}
	default:
		return 0, 0, false
	}
	colorsUsed := int(binary.LittleEndian.Uint32(b[32:]))
	paletteEntries := 0
	if bitCount <= 8 {
		paletteEntries = 1 << bitCount
		if colorsUsed > 0 && colorsUsed < paletteEntries {
			paletteEntries = colorsUsed
		}
	}
	pixelOffset := headerSize + maskBytes + paletteEntries*4
	if pixelOffset < headerSize || pixelOffset > len(b) {
		return 0, 0, false
	}
	absHeight := height
	if absHeight < 0 {
		absHeight = -absHeight
	}
	rowBits := int64(width) * int64(bitCount)
	stride := ((rowBits + 31) / 32) * 4
	if stride <= 0 || stride > int64(len(b)) {
		return 0, 0, false
	}
	pixelBytes := int(stride * int64(absHeight))
	imageSize := int(binary.LittleEndian.Uint32(b[20:]))
	if imageSize > pixelBytes {
		pixelBytes = imageSize
	}
	return pixelOffset, pixelBytes, true
}

func validDIBBitfieldMasks(b []byte, bitCount uint16, requireAlpha bool) bool {
	if len(b) < 12 {
		return false
	}
	limitMask := uint32(0xffffffff)
	if bitCount < 32 {
		limitMask = (uint32(1) << bitCount) - 1
	}
	used := uint32(0)
	for i := 0; i < 3; i++ {
		mask := binary.LittleEndian.Uint32(b[i*4:])
		if mask == 0 || mask&^limitMask != 0 || !contiguousBitMask(mask) || used&mask != 0 {
			return false
		}
		used |= mask
	}
	if len(b) >= 16 {
		alpha := binary.LittleEndian.Uint32(b[12:])
		if alpha != 0 {
			if alpha&^limitMask != 0 || !contiguousBitMask(alpha) || used&alpha != 0 {
				return false
			}
		} else if requireAlpha {
			return false
		}
	} else if requireAlpha {
		return false
	}
	return true
}

func contiguousBitMask(mask uint32) bool {
	for mask&1 == 0 {
		mask >>= 1
	}
	return mask != 0 && (mask&(mask+1)) == 0
}

func bmpDeclaredSize(b []byte) (int, bool) {
	if len(b) < 26 || !bytes.Equal(b[:2], []byte("BM")) {
		return 0, false
	}
	size := int(binary.LittleEndian.Uint32(b[2:]))
	if size < 26 || size > len(b) {
		return 0, false
	}
	if binary.LittleEndian.Uint16(b[6:]) != 0 || binary.LittleEndian.Uint16(b[8:]) != 0 {
		return 0, false
	}
	pixelOffset := int(binary.LittleEndian.Uint32(b[10:]))
	if pixelOffset < 14 || pixelOffset > size {
		return 0, false
	}
	minDIBPixelOffset, pixelBytes, ok := dibImageLayout(b[14:size])
	if !ok {
		return 0, false
	}
	minPixelOffset := 14 + minDIBPixelOffset
	if pixelOffset < minPixelOffset {
		return 0, false
	}
	if pixelOffset+pixelBytes < pixelOffset || pixelOffset+pixelBytes > size {
		return 0, false
	}
	return size, true
}

func validEMFData(b []byte) bool {
	_, ok := emfDeclaredSize(b)
	return ok
}

func emfDeclaredSize(b []byte) (int, bool) {
	if len(b) < 88 {
		return 0, false
	}
	if binary.LittleEndian.Uint32(b) != 1 {
		return 0, false
	}
	headerSize := int(binary.LittleEndian.Uint32(b[4:]))
	if headerSize < 88 || headerSize > len(b) || headerSize%4 != 0 {
		return 0, false
	}
	if !bytes.Equal(b[40:44], []byte{' ', 'E', 'M', 'F'}) {
		return 0, false
	}
	size := headerSize
	if len(b) >= 56 {
		declaredSize := int(binary.LittleEndian.Uint32(b[48:]))
		records := binary.LittleEndian.Uint32(b[52:])
		if declaredSize == 0 || records == 0 {
			return 0, false
		}
		if declaredSize < headerSize || declaredSize > len(b) || declaredSize%4 != 0 {
			return 0, false
		}
		size = declaredSize
	}
	if !validEMFRecordChain(b[:size], headerSize) {
		return 0, false
	}
	return size, true
}

func validEMFRecordChain(b []byte, headerSize int) bool {
	if headerSize < 8 || headerSize > len(b) {
		return false
	}
	if binary.LittleEndian.Uint32(b[0:]) != 1 || int(binary.LittleEndian.Uint32(b[4:])) != headerSize {
		return false
	}
	lastType := uint32(0)
	for off := headerSize; off < len(b); {
		if off+8 > len(b) {
			return false
		}
		recordType := binary.LittleEndian.Uint32(b[off:])
		recordSize := int(binary.LittleEndian.Uint32(b[off+4:]))
		if recordSize < 8 || recordSize%4 != 0 || recordSize > len(b)-off {
			return false
		}
		lastType = recordType
		off += recordSize
	}
	return lastType == 14
}

func validWMFData(b []byte) bool {
	_, ok := wmfDeclaredSize(b)
	return ok
}

func wmfDeclaredSize(b []byte) (int, bool) {
	if len(b) < 18 {
		return 0, false
	}
	if len(b) >= 40 && binary.LittleEndian.Uint32(b) == 0x9ac6cdd7 {
		if !validPlaceableWMFChecksum(b[:22]) {
			return 0, false
		}
		size, ok := standardWMFDeclaredSize(b[22:])
		if !ok {
			return 0, false
		}
		return 22 + size, true
	}
	return standardWMFDeclaredSize(b)
}

func standardWMFDeclaredSize(b []byte) (int, bool) {
	if len(b) < 18 {
		return 0, false
	}
	mtType := binary.LittleEndian.Uint16(b)
	headerSize := binary.LittleEndian.Uint16(b[2:])
	version := binary.LittleEndian.Uint16(b[4:])
	if (mtType != 1 && mtType != 2) || headerSize != 9 || (version != 0x0100 && version != 0x0300) {
		return 0, false
	}
	sizeWords := binary.LittleEndian.Uint32(b[6:])
	sizeBytes := uint64(sizeWords) * 2
	if sizeBytes < 24 || sizeBytes > uint64(len(b)) {
		return 0, false
	}
	size := int(sizeBytes)
	maxRecordWords := binary.LittleEndian.Uint32(b[12:])
	if maxRecordWords < 3 || maxRecordWords > sizeWords {
		return 0, false
	}
	if binary.LittleEndian.Uint32(b[size-6:]) != 3 || binary.LittleEndian.Uint16(b[size-2:]) != 0 {
		return 0, false
	}
	if !validStandardWMFRecords(b[:size], maxRecordWords) {
		return 0, false
	}
	return size, true
}

func validStandardWMFRecords(b []byte, maxRecordWords uint32) bool {
	for off := 18; off+6 <= len(b); {
		recordWords := binary.LittleEndian.Uint32(b[off:])
		function := binary.LittleEndian.Uint16(b[off+4:])
		if recordWords < 3 || recordWords > maxRecordWords {
			return false
		}
		recordBytes := uint64(recordWords) * 2
		if recordBytes > uint64(len(b)-off) {
			return false
		}
		end := off + int(recordBytes)
		if function == 0 {
			return recordWords == 3 && end == len(b)
		}
		off = end
	}
	return false
}

func validPlaceableWMFChecksum(b []byte) bool {
	if len(b) < 22 {
		return false
	}
	var checksum uint16
	for i := 0; i < 20; i += 2 {
		checksum ^= binary.LittleEndian.Uint16(b[i:])
	}
	return checksum == binary.LittleEndian.Uint16(b[20:])
}

func sanitizeFilename(name string) string {
	var b strings.Builder
	for _, r := range name {
		switch {
		case isInvisibleFormatControlRune(r):
			continue
		case r < 0x20 || r == 0x7f:
			b.WriteByte('_')
		case r != ' ' && unicode.IsSpace(r):
			b.WriteByte(' ')
		case !unicode.IsPrint(r):
			continue
		case strings.ContainsRune(`<>:"/\|?*`, r):
			b.WriteByte('_')
		default:
			b.WriteRune(r)
		}
	}
	cleaned := b.String()
	if ext := filepath.Ext(strings.TrimRight(cleaned, " .")); ext != "" {
		base := strings.Trim(strings.TrimSuffix(cleaned, ext), " .")
		if base == "" {
			name = "image" + ext
		} else {
			name = strings.Trim(cleaned, " .")
		}
	} else {
		name = strings.Trim(cleaned, " .")
	}
	if name == "" || name == "." || name == ".." {
		return "image.bin"
	}
	ext := filepath.Ext(name)
	base := strings.TrimSuffix(name, ext)
	if isWindowsReservedFilename(base) {
		name = base + "_" + ext
	}
	return truncateFilenameBytes(name, maxImageFilenameBytes)
}

func isWindowsReservedFilename(base string) bool {
	base = strings.Trim(strings.TrimSpace(base), ".")
	upper := strings.ToUpper(base)
	switch upper {
	case "CON", "PRN", "AUX", "NUL", "CLOCK$":
		return true
	}
	if len(upper) == 4 {
		prefix := upper[:3]
		suffix := upper[3]
		if (prefix == "COM" || prefix == "LPT") && suffix >= '1' && suffix <= '9' {
			return true
		}
	}
	return false
}

func truncateFilenameBytes(name string, maxBytes int) string {
	if maxBytes <= 0 || len(name) <= maxBytes {
		return name
	}
	ext := filepath.Ext(name)
	base := strings.TrimSuffix(name, ext)
	return truncateFilenameBaseBytes(base, ext, maxBytes) + ext
}

func truncateFilenameBaseBytes(base, ext string, maxBytes int) string {
	extBytes := len(ext)
	if maxBytes <= extBytes {
		return trimStringBytes(base, maxBytes)
	}
	baseLimit := maxBytes - extBytes
	base = trimStringBytes(base, baseLimit)
	base = strings.Trim(base, " .")
	if base == "" {
		base = "image"
	}
	if isWindowsReservedFilename(base) {
		base = trimStringBytes(base+"_", baseLimit)
	}
	return base
}

func trimStringBytes(s string, maxBytes int) string {
	if maxBytes <= 0 {
		return ""
	}
	if len(s) <= maxBytes {
		return s
	}
	for i := range s {
		if i > maxBytes {
			break
		}
		if i == maxBytes {
			return s[:i]
		}
	}
	for len(s) > maxBytes {
		_, size := utf8.DecodeLastRuneInString(s)
		if size <= 0 {
			return ""
		}
		s = s[:len(s)-size]
	}
	return s
}

func atoi(s string) (int, bool) {
	n := 0
	if s == "" {
		return 0, false
	}
	for _, r := range s {
		if r < '0' || r > '9' {
			return 0, false
		}
		n = n*10 + int(r-'0')
	}
	return n, true
}

func naturalLess(a, b string) bool {
	ra, rb := []rune(a), []rune(b)
	for ia, ib := 0, 0; ia < len(ra) && ib < len(rb); {
		if unicode.IsDigit(ra[ia]) && unicode.IsDigit(rb[ib]) {
			ja, jb := ia, ib
			for ja < len(ra) && unicode.IsDigit(ra[ja]) {
				ja++
			}
			for jb < len(rb) && unicode.IsDigit(rb[jb]) {
				jb++
			}
			na, _ := atoi(string(ra[ia:ja]))
			nb, _ := atoi(string(rb[ib:jb]))
			if na != nb {
				return na < nb
			}
			ia, ib = ja, jb
			continue
		}
		if ra[ia] != rb[ib] {
			return ra[ia] < rb[ib]
		}
		ia++
		ib++
	}
	return len(ra) < len(rb)
}

func utf8ValidOrASCII(s string) string {
	if utf8.ValidString(s) {
		return s
	}
	return strings.ToValidUTF8(s, "")
}
