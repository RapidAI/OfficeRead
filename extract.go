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
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"math"
	"mime/quotedprintable"
	"net/url"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"sync"
	"unicode"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/richardlehane/mscfb"
	textencoding "golang.org/x/text/encoding"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/encoding/korean"
	"golang.org/x/text/encoding/simplifiedchinese"
	"golang.org/x/text/encoding/traditionalchinese"
)

const maxOOXMLEmbeddedDepth = 3
const maxRepeatedTextPartBytes = 4096
const maxCompressedMetafileBytes = 256 << 20
const maxSmallDuplicateLegacyImageBytes = 4096
const maxImageFilenameBytes = 180
const maxMarkdownTableRows = 50000
const maxMarkdownTableCols = 1024
const maxMarkdownTableCellBytes = 512 << 10
const maxHiddenResourceMetadataReferenceBytes = 8192
const markdownIndentMarker = "\ue000"
const minMarkdownVisibleShortDigitBytes = 1 << 20
const maxMarkdownVisibleShortDigitBytes = 8 << 20
const minMarkdownVisibleShortDigitLines = 30000
const maxMarkdownVisibleShortDigitLineBytes = 512

type Result struct {
	Text               string
	StructuredMarkdown string
	Images             []Image
}

type Image struct {
	Name string
	Alt  string
	Ext  string
	Data []byte
}

type Options struct {
	ImageDir        string
	IncludeMetadata bool
}

func Extract(filename string, opts Options) (*Result, error) {
	data, err := os.ReadFile(filename)
	if err != nil {
		return nil, err
	}
	var res *Result
	if isZip(data) {
		res, err = extractOOXML(filename, data, opts)
	} else {
		res, err = extractLegacy(filename, data, opts)
	}
	if err != nil {
		return nil, err
	}
	res.Images = finalizeOutputImages(res.Images)
	if opts.ImageDir != "" {
		if err := writeImages(opts.ImageDir, res.Images); err != nil {
			return nil, err
		}
	}
	return res, nil
}

func writeImages(dir string, images []Image) error {
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return err
	}
	images = validOutputImages(images)
	names := imageOutputFilenames(images)
	for i, img := range images {
		if err := os.WriteFile(filepath.Join(dir, names[i]), img.Data, 0o644); err != nil {
			return err
		}
	}
	return nil
}

func validOutputImages(images []Image) []Image {
	if len(images) == 0 {
		return nil
	}
	out := make([]Image, 0, len(images))
	for _, img := range images {
		data, ext, ok := normalizeOOXMLImageData(img.Ext, img.Data)
		if !ok {
			continue
		}
		img.Data = data
		img.Ext = ext
		if img.Name != "" {
			img.Name = imageNameWithExt(img.Name, ext)
		}
		out = append(out, img)
	}
	return out
}

func finalizeOutputImages(images []Image) []Image {
	images = validOutputImages(images)
	names := imageOutputFilenames(images)
	for i := range images {
		if i < len(names) {
			images[i].Name = names[i]
		}
	}
	return images
}

func imageOutputFilenames(images []Image) []string {
	names := make([]string, len(images))
	used := map[string]bool{}
	for i, img := range images {
		name := imageOutputNameCandidate(img.Name)
		if name == "" {
			name = fmt.Sprintf("image-%03d%s", i+1, img.Ext)
		}
		if img.Ext != "" {
			name = imageNameWithExt(name, strings.ToLower(img.Ext))
		}
		name = cleanImageOutputFilename(name, img.Ext, i)
		name = sanitizeFilename(name)
		if filepath.Ext(name) == "" && img.Ext != "" {
			name += img.Ext
		}
		names[i] = uniqueImageFilename(name, used)
	}
	return names
}

func cleanImageOutputFilename(name, ext string, index int) string {
	name = strings.TrimSpace(name)
	if name == "" {
		return name
	}
	rawExt := path.Ext(name)
	base := strings.TrimSpace(strings.TrimSuffix(name, rawExt))
	cleanBase := cleanText(base)
	if cleanBase == "" {
		return name
	}
	cleanBase = stripInlineHiddenOfficeReferences(cleanBase)
	if cleanBase == "" ||
		looksLikeRelationshipIDReference(cleanBase) ||
		looksLikeOfficeRelationshipMetadataReference(cleanBase) ||
		looksLikeOfficeXMLMetadataReference(cleanBase) {
		cleanBase = fmt.Sprintf("image-%03d", index+1)
	}
	if rawExt == "" {
		rawExt = ext
	}
	return cleanBase + strings.ToLower(rawExt)
}

func imageOutputNameCandidate(name string) string {
	name = strings.TrimSpace(name)
	if name == "" {
		return ""
	}
	queue := []string{name}
	seen := map[string]bool{}
	for len(queue) > 0 && len(seen) < 8 {
		cur := strings.TrimSpace(queue[0])
		queue = queue[1:]
		if cur == "" || seen[cur] {
			continue
		}
		seen[cur] = true
		normalized := strings.ReplaceAll(cur, "\\", "/")
		if part := hiddenPackageURIPathCandidate(normalized); part != "" {
			if decoded, err := url.PathUnescape(stripOfficePartPathSuffix(part)); err == nil && decoded != "" {
				part = decoded
			}
			if base := path.Base(stripOfficePartPathSuffix(part)); base != "" && base != "." && base != "/" {
				return base
			}
		}
		if looksLikeOfficePartPath(strings.ToLower(normalized)) {
			part := stripOfficePartPathSuffix(normalized)
			if decoded, err := url.PathUnescape(part); err == nil && decoded != "" {
				part = decoded
			}
			return path.Base(part)
		}
		if base := remoteImagePathBaseName(normalized); base != "" {
			return base
		}
		if base := localImagePathBaseName(normalized); base != "" {
			return base
		}
		if strings.Contains(normalized, "/") {
			if base := supportedImagePathBaseName(normalized); base != "" {
				return base
			}
		}
		if decoded, err := url.PathUnescape(normalized); err == nil && decoded != normalized {
			queue = append(queue, decoded)
		}
	}
	return name
}

func remoteImagePathBaseName(name string) string {
	trimmed := stripOfficePartPathSuffix(strings.TrimSpace(name))
	if trimmed == "" {
		return ""
	}
	lower := strings.ToLower(trimmed)
	switch {
	case strings.HasPrefix(lower, "http://") || strings.HasPrefix(lower, "https://") || strings.HasPrefix(lower, "ftp://"):
		if u, err := url.Parse(trimmed); err == nil && u.Host != "" {
			if base := supportedImagePathBaseName(u.Path); base != "" {
				return base
			}
			if base := supportedImagePathBaseName(u.Opaque); base != "" {
				return base
			}
		}
	case strings.HasPrefix(lower, "cid:") || strings.HasPrefix(lower, "mid:"):
		value := strings.TrimSpace(trimmed[4:])
		if at := strings.IndexByte(value, '@'); at > 0 {
			value = value[:at]
		}
		if base := supportedImagePathBaseName(value); base != "" {
			return base
		}
	}
	return ""
}

func localImagePathBaseName(name string) string {
	trimmed := stripOfficePartPathSuffix(strings.TrimSpace(name))
	if trimmed == "" {
		return ""
	}
	lower := strings.ToLower(trimmed)
	if strings.HasPrefix(lower, "file:/") || strings.Contains(lower, "|file:/") {
		value := trimmed
		if i := strings.LastIndex(strings.ToLower(value), "|file:/"); i >= 0 {
			value = value[i+1:]
		}
		if u, err := url.Parse(value); err == nil {
			candidates := []string{u.Path, u.Opaque}
			for _, candidate := range candidates {
				if base := supportedImagePathBaseName(candidate); base != "" {
					return base
				}
			}
		}
		value = strings.TrimPrefix(value, "file:")
		value = strings.TrimLeft(value, "/")
		if base := supportedImagePathBaseName(value); base != "" {
			return base
		}
	}
	if strings.HasPrefix(trimmed, "//") {
		if base := supportedImagePathBaseName(trimmed); base != "" {
			return base
		}
	}
	if len(trimmed) >= 3 && isASCIILetter(rune(trimmed[0])) && trimmed[1] == ':' && trimmed[2] == '/' {
		if base := supportedImagePathBaseName(trimmed); base != "" {
			return base
		}
	}
	return ""
}

func supportedImagePathBaseName(name string) string {
	name = strings.TrimSpace(strings.ReplaceAll(name, "\\", "/"))
	if decoded, err := url.PathUnescape(name); err == nil && decoded != "" {
		name = decoded
	}
	base := path.Base(strings.TrimRight(name, "/"))
	if base == "" || base == "." || base == "/" || !isSupportedImageExt(path.Ext(base)) {
		return ""
	}
	return base
}

func stripOfficePartPathSuffix(name string) string {
	if i := strings.IndexAny(name, "?#"); i >= 0 {
		name = name[:i]
	}
	return strings.TrimRight(name, ".,;:)]}>")
}

func uniquifyImageNames(images []Image) {
	used := map[string]bool{}
	for i := range images {
		if images[i].Name == "" {
			continue
		}
		images[i].Name = uniqueImageFilename(sanitizeFilename(images[i].Name), used)
	}
}

func (r *Result) Markdown(imageBase string) string {
	if r == nil {
		return ""
	}
	var out strings.Builder
	images := validOutputImages(r.Images)
	text := normalizeMarkdownOutputSpaces(strings.TrimSpace(r.StructuredMarkdown))
	if text != "" {
		text = appendMissingMarkdownText(text, r.Text, images)
	} else {
		text = markdownText(r.Text)
	}
	text = normalizeMarkdownOutputSpaces(text)
	if text != "" {
		out.WriteString(text)
	}
	names := imageOutputFilenames(images)
	placed := placeMarkdownImages(&out, images, names, imageBase)
	if len(names) > 0 {
		if hasUnplacedImages(placed) && out.Len() > 0 {
			out.WriteString("\n\n## Images")
		} else if hasUnplacedImages(placed) {
			out.WriteString("## Images")
		}
	}
	for i, name := range names {
		if placed[i] {
			continue
		}
		out.WriteString("\n\n")
		out.WriteString(markdownImage(images[i], imageBase, name, i))
	}
	return strings.TrimSpace(out.String())
}

func normalizeMarkdownOutputSpaces(s string) string {
	if s == "" {
		return s
	}
	return strings.Map(cleanTextRune, strings.ToValidUTF8(s, ""))
}

func hasUnplacedImages(placed []bool) bool {
	for _, ok := range placed {
		if !ok {
			return true
		}
	}
	return false
}

func placeMarkdownImages(out *strings.Builder, images []Image, names []string, imageBase string) []bool {
	placed := make([]bool, len(names))
	if out.Len() == 0 || len(images) == 0 || len(names) == 0 {
		return placed
	}
	lines := strings.Split(out.String(), "\n")
	placementLines := append([]string(nil), lines...)
	changed := false
	for i, img := range images {
		if i >= len(names) {
			break
		}
		alt := cleanMarkdownImageAltText(img.Alt)
		if alt == "" {
			continue
		}
		var matches []markdownImagePlacementMatch
		inFence := false
		inHTMLComment := false
		htmlBlockEnd := ""
		for lineIndex, line := range placementLines {
			if markdownFenceLine(line) {
				inFence = !inFence
				continue
			}
			trimmed := strings.TrimSpace(line)
			if markdownHTMLCommentStart(trimmed) {
				if !strings.Contains(trimmed, "-->") {
					inHTMLComment = true
				}
				continue
			}
			if inHTMLComment {
				if strings.Contains(trimmed, "-->") {
					inHTMLComment = false
				}
				continue
			}
			if htmlBlockEnd != "" {
				if markdownHTMLBlockEnd(trimmed, htmlBlockEnd) {
					htmlBlockEnd = ""
				}
				continue
			}
			if endTag, ok := markdownHTMLBlockStart(trimmed); ok {
				htmlBlockEnd = endTag
				continue
			}
			if inFence || markdownIndentedCodeLine(line) || markdownHTMLTagLine(trimmed) || strings.HasPrefix(trimmed, "#") {
				continue
			}
			if markdownThematicBreakLine(trimmed) {
				continue
			}
			if strings.Contains(line, "|") {
				for _, cellIndex := range markdownTableImagePlacementCells(line, alt) {
					matches = append(matches, markdownImagePlacementMatch{line: lineIndex, cell: cellIndex})
					if len(matches) > 1 {
						break
					}
				}
				if len(matches) > 1 {
					break
				}
				continue
			}
			if strings.Contains(line, "![") {
				continue
			}
			if markdownVisibleLineText(line) != alt {
				continue
			}
			matches = append(matches, markdownImagePlacementMatch{line: lineIndex, cell: -1})
			if len(matches) > 1 {
				break
			}
		}
		if len(matches) == 1 {
			match := matches[0]
			image := markdownImage(img, imageBase, names[i], i)
			if match.cell >= 0 {
				lines[match.line] = markdownTableImagePlacementLine(lines[match.line], match.cell, image)
			} else {
				lines[match.line] = markdownImagePlacementLine(lines[match.line], image)
			}
			placed[i] = true
			changed = true
		}
	}
	if changed {
		out.Reset()
		out.WriteString(strings.Join(lines, "\n"))
	}
	return placed
}

type markdownImagePlacementMatch struct {
	line int
	cell int
}

func markdownImagePlacementLine(line, image string) string {
	prefix, rest := markdownStructuralLinePrefix(line)
	if strings.TrimSpace(rest) == "" {
		return image
	}
	return line + "\n" + prefix + image
}

func markdownTableImagePlacementCells(line, alt string) []int {
	trimmed := strings.TrimSpace(line)
	if !markdownLikelyTableRow(trimmed) {
		return nil
	}
	parts := splitMarkdownTableRow(strings.Trim(trimmed, "|"))
	matches := make([]int, 0, 1)
	for i, part := range parts {
		if strings.Contains(part, "![") {
			continue
		}
		if markdownVisibleTableCellText(part) == alt {
			matches = append(matches, i)
		}
	}
	return matches
}

func markdownLikelyTableRow(trimmed string) bool {
	if !strings.Contains(trimmed, "|") {
		return false
	}
	if strings.HasPrefix(trimmed, "|") || strings.HasSuffix(trimmed, "|") {
		return true
	}
	return len(splitMarkdownTableRow(trimmed)) >= 3
}

func markdownVisibleTableCellText(part string) string {
	part = strings.TrimSpace(unescapeMarkdownVisibleText(part))
	part = strings.TrimSpace(stripMarkdownFootnoteReferences(part))
	part = strings.TrimSpace(stripMarkdownInlineWrappers(part))
	part = strings.TrimSpace(stripMarkdownInlineFormatting(part))
	part = strings.TrimSpace(unescapeMarkdownInlineFormattingMarkers(part))
	part = strings.TrimSpace(markdownAutolinkVisibleText(part))
	part = strings.TrimSpace(markdownVisibleHTMLText(part))
	part = strings.TrimSpace(stripMarkdownHardLineBreakMarker(part))
	if part == "" || markdownTableSeparatorCell(part) {
		return ""
	}
	return part
}

func markdownTableImagePlacementLine(line string, cell int, image string) string {
	leading := line[:len(line)-len(strings.TrimLeft(line, " \t"))]
	trimmed := strings.TrimSpace(line)
	hasLeadingPipe := strings.HasPrefix(trimmed, "|")
	hasTrailingPipe := strings.HasSuffix(trimmed, "|")
	parts := splitMarkdownTableRow(strings.Trim(trimmed, "|"))
	if cell < 0 || cell >= len(parts) {
		return line
	}
	if strings.Contains(parts[cell], "![") {
		parts[cell] = strings.TrimSpace(parts[cell]) + "<br>" + image
	} else {
		parts[cell] = image
	}
	var out strings.Builder
	out.WriteString(leading)
	if hasLeadingPipe {
		out.WriteString("|")
	}
	for i, part := range parts {
		if i > 0 {
			out.WriteString("|")
		}
		out.WriteByte(' ')
		out.WriteString(strings.TrimSpace(part))
		out.WriteByte(' ')
	}
	if hasTrailingPipe {
		out.WriteString("|")
	}
	return out.String()
}

func markdownStructuralLinePrefix(line string) (string, string) {
	var prefix strings.Builder
	leading := len(line) - len(strings.TrimLeft(line, " \t"))
	if leading > 0 {
		prefix.WriteString(line[:leading])
		line = line[leading:]
	}
	for strings.HasPrefix(strings.TrimLeft(line, " \t"), ">") {
		quoteIndent := len(line) - len(strings.TrimLeft(line, " \t"))
		if quoteIndent > 0 {
			prefix.WriteString(line[:quoteIndent])
			line = line[quoteIndent:]
		}
		prefix.WriteString("> ")
		line = strings.TrimSpace(line[1:])
	}
	if p, rest, ok := markdownListLinePrefix(line); ok {
		prefix.WriteString(p)
		line = rest
	}
	return prefix.String(), line
}

func markdownListLinePrefix(line string) (string, string, bool) {
	if len(line) >= 2 && (line[0] == '-' || line[0] == '*' || line[0] == '+') && unicode.IsSpace(rune(line[1])) {
		marker := line[:2]
		rest := strings.TrimSpace(line[2:])
		if task, taskRest, ok := markdownTaskLinePrefix(rest); ok {
			return marker + task, taskRest, true
		}
		return marker, rest, true
	}
	i := 0
	for i < len(line) && line[i] >= '0' && line[i] <= '9' {
		i++
	}
	if i > 0 && i+1 < len(line) && (line[i] == '.' || line[i] == ')') && unicode.IsSpace(rune(line[i+1])) {
		marker := line[:i+2]
		rest := strings.TrimSpace(line[i+2:])
		if task, taskRest, ok := markdownTaskLinePrefix(rest); ok {
			return marker + task, taskRest, true
		}
		return marker, rest, true
	}
	return "", line, false
}

func markdownTaskLinePrefix(line string) (string, string, bool) {
	if len(line) >= 4 && line[0] == '[' && line[2] == ']' && unicode.IsSpace(rune(line[3])) {
		switch line[1] {
		case ' ', 'x', 'X':
			return line[:4], strings.TrimSpace(line[4:]), true
		}
	}
	return "", line, false
}

func markdownFenceLine(line string) bool {
	trimmed := strings.TrimSpace(line)
	return strings.HasPrefix(trimmed, "```") || strings.HasPrefix(trimmed, "~~~")
}

func markdownThematicBreakLine(trimmed string) bool {
	if trimmed == "" {
		return false
	}
	var marker rune
	count := 0
	for _, r := range trimmed {
		if unicode.IsSpace(r) {
			continue
		}
		if marker == 0 {
			switch r {
			case '-', '*', '_':
				marker = r
			default:
				return false
			}
		}
		if r != marker {
			return false
		}
		count++
	}
	return count >= 3
}

func markdownIndentedCodeLine(line string) bool {
	if strings.HasPrefix(line, "\t") {
		return true
	}
	spaces := 0
	for spaces < len(line) && line[spaces] == ' ' {
		spaces++
	}
	if spaces < 4 {
		return false
	}
	trimmed := strings.TrimSpace(line)
	if trimmed == "" || markdownLineStartsWithListMarker(trimmed) || strings.HasPrefix(trimmed, ">") {
		return false
	}
	return true
}

func markdownHTMLCommentStart(trimmed string) bool {
	return strings.HasPrefix(trimmed, "<!--")
}

func markdownHTMLTagLine(trimmed string) bool {
	if len(trimmed) < 3 || trimmed[0] != '<' {
		return false
	}
	if strings.HasPrefix(trimmed, "<!") || strings.HasPrefix(trimmed, "<?") {
		return true
	}
	if trimmed[1] == '/' {
		return len(trimmed) > 3 && isASCIILetter(rune(trimmed[2]))
	}
	return isASCIILetter(rune(trimmed[1]))
}

func markdownHTMLBlockStart(trimmed string) (string, bool) {
	if len(trimmed) < 3 || trimmed[0] != '<' || trimmed[1] == '/' || strings.HasPrefix(trimmed, "<!") || strings.HasPrefix(trimmed, "<?") {
		return "", false
	}
	i := 1
	for i < len(trimmed) && (isASCIILetter(rune(trimmed[i])) || (trimmed[i] >= '0' && trimmed[i] <= '9')) {
		i++
	}
	if i == 1 {
		return "", false
	}
	tag := strings.ToLower(trimmed[1:i])
	if !markdownHTMLBlockTag(tag) {
		return "", false
	}
	endTag := "</" + tag + ">"
	lower := strings.ToLower(trimmed)
	if strings.Contains(lower, endTag) || strings.HasSuffix(strings.TrimSpace(lower), "/>") {
		return "", false
	}
	return endTag, true
}

func markdownHTMLBlockEnd(trimmed, endTag string) bool {
	return strings.Contains(strings.ToLower(trimmed), endTag)
}

func markdownHTMLBlockTag(tag string) bool {
	switch tag {
	case "address", "article", "aside", "blockquote", "body", "details", "dialog", "div", "dl", "fieldset", "figcaption", "figure", "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6", "head", "header", "hr", "html", "li", "main", "nav", "ol", "p", "pre", "section", "table", "tbody", "td", "tfoot", "th", "thead", "tr", "ul":
		return true
	default:
		return false
	}
}

func markdownImage(img Image, imageBase, name string, index int) string {
	target := markdownImageTarget(imageBase, name)
	alt := markdownImageAlt(img, name, index)
	return "![" + escapeMarkdownImageAlt(alt) + "](" + escapeMarkdownLinkTarget(target) + ")"
}

func appendMissingMarkdownText(markdown, text string, images []Image) string {
	missing := missingMarkdownText(markdown, text, images)
	if missing == "" {
		return strings.TrimSpace(markdown)
	}
	markdown = strings.TrimSpace(markdown)
	if markdown == "" {
		return missing
	}
	return markdown + "\n\n## Additional Text\n\n" + missing
}

func missingMarkdownText(markdown, text string, images []Image) string {
	return missingMarkdownTextWithOptions(markdown, text, images, false, false)
}

func missingMarkdownTextXLS(markdown, text string) string {
	return missingMarkdownTextWithOptions(markdown, text, nil, true, true)
}

func missingMarkdownTextWithOptions(markdown, text string, images []Image, escapedTableOnlyWhenPipe, shortTableExactBeforeMinLen bool) string {
	lines := markdownBackfillSourceLines(markdownBackfillRawLines(text))
	if len(lines) == 0 {
		return ""
	}
	imageAlts := markdownImageAltSet(images)
	candidateLineCache := map[string]string{}
	candidateLine := func(raw string) string {
		if line, ok := candidateLineCache[raw]; ok {
			return line
		}
		line := markdownBackfillCandidateLine(raw)
		candidateLineCache[raw] = line
		return line
	}
	exact := markdownBackfillExactSet(markdown)
	if markdownBackfillExactLinesCoveredWithExact(exact, lines, imageAlts, candidateLine) {
		return ""
	}
	if escapedTableOnlyWhenPipe && shortTableExactBeforeMinLen && shouldTryMarkdownVisibleOnlyFallback(lines, text) {
		if markdownBackfillLinesCoveredWithExactOrVisibleOnly(markdown, exact, lines, imageAlts, candidateLine) {
			return ""
		}
	}
	coverage, containment := markdownBackfillBuildCoverageContainment(markdown)
	containment.shortTableExactBeforeMinLen = shortTableExactBeforeMinLen
	coverageCache := map[string]bool{}
	tableContainsCache := map[string]bool{}
	visibleContainsCache := map[string]bool{}
	coverageContains := func(line string) bool {
		if line == "" {
			return false
		}
		if covered, ok := coverageCache[line]; ok {
			return covered
		}
		covered := markdownBackfillCoverageCoversLine(coverage, line)
		coverageCache[line] = covered
		return covered
	}
	tableTextContains := func(line string) bool {
		if line == "" {
			return false
		}
		if covered, ok := tableContainsCache[line]; ok {
			return covered
		}
		covered := containment.tableTextContainsLine(line)
		tableContainsCache[line] = covered
		return covered
	}
	visibleLineContains := func(line string) bool {
		if line == "" {
			return false
		}
		if covered, ok := visibleContainsCache[line]; ok {
			return covered
		}
		covered := containment.visibleLineContainsLine(line)
		visibleContainsCache[line] = covered
		return covered
	}
	var missing []string
	seenMissing := map[string]bool{}
	lineMissingCache := map[string]bool{}
	collapsedCoveredCache := map[string]bool{}
	for i := 0; i < len(lines); i++ {
		line := candidateLine(lines[i])
		if line == "" {
			continue
		}
		if !markdownBackfillNormalizedLineAllowed(line, imageAlts) {
			lineMissingCache[line] = false
			continue
		}
		if isMarkdownSingleASCIIWordCharLine(line) {
			run := []string{line}
			j := i + 1
			for ; j < len(lines); j++ {
				next := candidateLine(lines[j])
				if !isMarkdownSingleASCIIWordCharLine(next) {
					break
				}
				run = append(run, next)
			}
			if shouldCollapseMarkdownSingleCharacterRun(run) {
				collapsed := strings.Join(run, "")
				covered, ok := collapsedCoveredCache[collapsed]
				if !ok {
					covered = coverageContains(collapsed) || tableTextContains(collapsed) || visibleLineContains(collapsed)
					collapsedCoveredCache[collapsed] = covered
				}
				if covered {
					i = j - 1
					continue
				}
			}
		}
		if shouldAdd, ok := lineMissingCache[line]; ok {
			if shouldAdd && !seenMissing[line] {
				seenMissing[line] = true
				missing = append(missing, line)
			}
			continue
		}
		if seenMissing[line] {
			continue
		}
		variants := []string{line}
		if visibleLine := markdownBackfillVisibleText(line); visibleLine != "" && visibleLine != line {
			variants = append(variants, visibleLine)
		}
		if markdownLine := markdownVisibleLineText(line); markdownLine != "" {
			duplicate := false
			for _, variant := range variants {
				if variant == markdownLine {
					duplicate = true
					break
				}
			}
			if !duplicate {
				variants = append(variants, markdownLine)
			}
		}
		if !escapedTableOnlyWhenPipe || strings.IndexByte(line, '|') >= 0 {
			if escapedVisibleLine := markdownBackfillVisibleText(escapeMarkdownTableCell(line)); escapedVisibleLine != "" {
				duplicate := false
				for _, variant := range variants {
					if variant == escapedVisibleLine {
						duplicate = true
						break
					}
				}
				if !duplicate {
					variants = append(variants, escapedVisibleLine)
				}
			}
		}
		covered := false
		for _, variant := range variants {
			if coverageContains(variant) || tableTextContains(variant) || visibleLineContains(variant) {
				covered = true
				break
			}
		}
		shouldAdd := !covered
		lineMissingCache[line] = shouldAdd
		if !shouldAdd {
			continue
		}
		seenMissing[line] = true
		missing = append(missing, line)
	}
	return markdownText(strings.Join(missing, "\n"))
}

func shouldTryMarkdownVisibleOnlyFallback(lines []string, text string) bool {
	if len(lines) < 30000 {
		return false
	}
	return len(text) <= 512*1024
}

func markdownBackfillLinesCoveredWithExactOrVisibleOnly(markdown string, exact map[string]struct{}, lines []string, imageAlts map[string]bool, candidateLine func(string) string) bool {
	containment := markdownBackfillVisibleContainmentSet(markdown)
	seen := map[string]bool{}
	for _, raw := range lines {
		line := candidateLine(raw)
		if line == "" || seen[line] {
			continue
		}
		seen[line] = true
		if !markdownBackfillNormalizedLineAllowed(line, imageAlts) {
			continue
		}
		if markdownBackfillExactSetContainsLine(exact, line) {
			continue
		}
		if markdownBackfillVisibleOnlyContainsLine(containment, line) {
			continue
		}
		return false
	}
	return true
}

func markdownBackfillVisibleOnlyContainsLine(containment markdownBackfillContainment, line string) bool {
	variants := []string{line}
	if visibleLine := markdownBackfillVisibleText(line); visibleLine != "" && visibleLine != line {
		variants = append(variants, visibleLine)
	}
	if markdownLine := markdownVisibleLineText(line); markdownLine != "" {
		duplicate := false
		for _, variant := range variants {
			if variant == markdownLine {
				duplicate = true
				break
			}
		}
		if !duplicate {
			variants = append(variants, markdownLine)
		}
	}
	if escapedVisibleLine := markdownBackfillVisibleText(escapeMarkdownTableCell(line)); escapedVisibleLine != "" {
		duplicate := false
		for _, variant := range variants {
			if variant == escapedVisibleLine {
				duplicate = true
				break
			}
		}
		if !duplicate {
			variants = append(variants, escapedVisibleLine)
		}
	}
	for _, variant := range variants {
		if containment.visibleLineContainsLine(variant) {
			return true
		}
	}
	return false
}

func markdownBackfillRawLines(s string) []string {
	if text, ok := rtfVisibleText(s); ok {
		s = text
	}
	s = strings.ReplaceAll(s, "\r\n", "\n")
	s = strings.ReplaceAll(s, "\r", "\n")
	return strings.Split(s, "\n")
}

func markdownBackfillSourceLines(lines []string) []string {
	out := make([]string, 0, len(lines))
	blank := false
	for _, line := range lines {
		line = normalizeMarkdownTextLine(line)
		if line == "" {
			if len(out) > 0 {
				blank = true
			}
			continue
		}
		if blank {
			out = append(out, "")
			blank = false
		}
		out = append(out, line)
	}
	return out
}

func markdownBackfillCandidateLine(line string) string {
	line = strings.TrimSpace(line)
	if line == "" {
		return ""
	}
	line = cleanText(line)
	line = stripInlineHiddenOfficeReferences(line)
	return strings.TrimSpace(line)
}

func markdownBackfillExactLinesCovered(markdown string, lines []string, imageAlts map[string]bool, candidateLine func(string) string) bool {
	return markdownBackfillExactLinesCoveredWithExact(markdownBackfillExactSet(markdown), lines, imageAlts, candidateLine)
}

func markdownBackfillExactLinesCoveredWithExact(exact map[string]struct{}, lines []string, imageAlts map[string]bool, candidateLine func(string) string) bool {
	seen := map[string]bool{}
	for _, raw := range lines {
		line := candidateLine(raw)
		if line == "" || seen[line] {
			continue
		}
		seen[line] = true
		if !markdownBackfillNormalizedLineAllowed(line, imageAlts) {
			continue
		}
		if markdownBackfillExactSetContainsLine(exact, line) {
			continue
		}
		return false
	}
	return true
}

func markdownBackfillBuildCoverageContainment(markdown string) (map[string]struct{}, markdownBackfillContainment) {
	referenceDefinitions := markdownReferenceDefinitions(markdown)
	coverage := map[string]struct{}{}
	containment := markdownBackfillContainment{
		tableRawExact:        map[string]struct{}{},
		tableVisibleExact:    map[string]struct{}{},
		tableComparableExact: map[string]struct{}{},
		visibleExact:         map[string]struct{}{},
	}
	inFence := false
	for _, markdownLine := range strings.Split(markdown, "\n") {
		rawLine := markdownLine
		trimmed := strings.TrimSpace(markdownLine)
		if trimmed == "" {
			continue
		}
		if markdownFenceLine(trimmed) {
			inFence = !inFence
			continue
		}
		if inFence || markdownIndentedCodeLine(rawLine) || markdownLooksLikeReferenceDefinition(trimmed) {
			continue
		}
		addMarkdownBackfillCoverageText(coverage, trimmed)
		if markdownLikelyTableRow(trimmed) {
			containment.tableRaw = append(containment.tableRaw, rawLine)
			containment.tableRawExact[rawLine] = struct{}{}
			if len(rawLine) > containment.tableRawMaxLen {
				containment.tableRawMaxLen = len(rawLine)
			}
			visible := markdownBackfillVisibleText(rawLine)
			comparable := markdownBackfillComparableFromVisibleText(visible)
			if visible != "" {
				containment.tableVisible = append(containment.tableVisible, visible)
				containment.tableVisibleExact[visible] = struct{}{}
			}
			if len(visible) > containment.tableVisibleMaxLen {
				containment.tableVisibleMaxLen = len(visible)
			}
			if comparable != "" {
				containment.tableComparable = append(containment.tableComparable, comparable)
				containment.tableComparableExact[comparable] = struct{}{}
			}
			if len(comparable) > containment.tableComparableMaxLen {
				containment.tableComparableMaxLen = len(comparable)
			}
		}
		if !strings.Contains(trimmed, "[") {
			if visible := markdownVisibleLineText(rawLine); visible != "" {
				containment.visibleLines = append(containment.visibleLines, visible)
				containment.visibleExact[visible] = struct{}{}
				if len(visible) > containment.visibleMaxLen {
					containment.visibleMaxLen = len(visible)
				}
			}
		}
		if markdownPlainVisibleLine(trimmed) {
			continue
		}
		lineText := markdownInlineLinkVisibleText(trimmed, referenceDefinitions)
		if visibleText := markdownVisibleLineText(lineText); visibleText != "" {
			addMarkdownBackfillCoverageText(coverage, visibleText)
		}
		for _, cell := range markdownVisibleTableCells(lineText) {
			if cell == "" {
				continue
			}
			addMarkdownBackfillCoverageText(coverage, cell)
		}
	}
	containment.tableRawJoined = strings.Join(containment.tableRaw, "\n")
	containment.tableVisibleJoined = strings.Join(containment.tableVisible, "\n")
	containment.tableComparableJoined = strings.Join(containment.tableComparable, "\n")
	containment.visibleJoined = strings.Join(containment.visibleLines, "\n")
	maybeEnableMarkdownVisibleShortDigits(&containment)
	return coverage, containment
}

func markdownBackfillExactSet(markdown string) map[string]struct{} {
	exact := map[string]struct{}{}
	inFence := false
	referenceDefinitions := markdownReferenceDefinitions(markdown)
	tableCellCache := map[string][]string{}
	for _, visibleLine := range strings.Split(markdown, "\n") {
		rawLine := visibleLine
		visibleLine = strings.TrimSpace(visibleLine)
		if visibleLine == "" {
			continue
		}
		if markdownFenceLine(visibleLine) {
			inFence = !inFence
			continue
		}
		if inFence || markdownIndentedCodeLine(rawLine) || markdownLooksLikeReferenceDefinition(visibleLine) {
			continue
		}
		exact[visibleLine] = struct{}{}
		if markdownPlainVisibleLine(visibleLine) {
			continue
		}
		lineText := markdownInlineLinkVisibleText(visibleLine, referenceDefinitions)
		if visibleText := markdownVisibleLineText(lineText); visibleText != "" {
			exact[visibleText] = struct{}{}
		}
		for _, cell := range markdownVisibleTableCellsWithCache(lineText, tableCellCache) {
			if cell != "" {
				exact[cell] = struct{}{}
			}
		}
	}
	return exact
}

func markdownBackfillExactSetContainsLine(exact map[string]struct{}, line string) bool {
	if _, ok := exact[line]; ok {
		return true
	}
	if visibleLine := markdownBackfillVisibleText(line); visibleLine != "" && visibleLine != line {
		if _, ok := exact[visibleLine]; ok {
			return true
		}
	}
	if markdownLine := markdownVisibleLineText(line); markdownLine != "" && markdownLine != line {
		if _, ok := exact[markdownLine]; ok {
			return true
		}
	}
	if escapedVisibleLine := markdownBackfillVisibleText(escapeMarkdownTableCell(line)); escapedVisibleLine != "" && escapedVisibleLine != line {
		if _, ok := exact[escapedVisibleLine]; ok {
			return true
		}
	}
	return false
}

var markdownImageRE = regexp.MustCompile(`!\[([^\]]*)\]\([^)]+\)`)
var markdownLinkRE = regexp.MustCompile(`\[([^\]]+)\]\([^)]+\)`)
var markdownVisibleHTMLBreakRE = regexp.MustCompile(`(?is)<br\b[^>]*>`)
var markdownVisibleHTMLTagRE = regexp.MustCompile(`(?is)</?[A-Za-z][A-Za-z0-9:-]*(?:\s+[^<>]*)?>`)
var markdownAutoLinkRE = regexp.MustCompile("(?i)<((?:[a-z][a-z0-9+.-]{1,31}:[^\\s<>]*)|(?:[a-z0-9.!#$%&'*+/=?^_{|}~-]+@[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?(?:\\.[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?)+))>")

func markdownBackfillVisibleText(markdown string) string {
	markdown = markdownInlineLinkVisibleText(markdown, markdownReferenceDefinitions(markdown))
	markdown = strings.ReplaceAll(markdown, `\[`, "[")
	markdown = strings.ReplaceAll(markdown, `\]`, "]")
	markdown = strings.ReplaceAll(markdown, `\\`, `\`)
	return markdownText(markdown)
}

func markdownReferenceDefinitions(markdown string) map[string]struct{} {
	defs := make(map[string]struct{})
	inFence := false
	inHTMLComment := false
	htmlEndTag := ""
	for _, line := range strings.Split(markdown, "\n") {
		rawLine := line
		line = strings.TrimSpace(line)
		if markdownFenceLine(line) {
			inFence = !inFence
			continue
		}
		if inFence || markdownIndentedCodeLine(rawLine) {
			continue
		}
		if inHTMLComment {
			if strings.Contains(line, "-->") {
				inHTMLComment = false
			}
			continue
		}
		if htmlEndTag != "" {
			if markdownHTMLBlockEnd(line, htmlEndTag) {
				htmlEndTag = ""
			}
			continue
		}
		if markdownHTMLCommentStart(line) {
			if !strings.Contains(line, "-->") {
				inHTMLComment = true
			}
			continue
		}
		if endTag, ok := markdownHTMLBlockStart(line); ok {
			htmlEndTag = endTag
			continue
		}
		if len(line) < 4 || line[0] != '[' {
			continue
		}
		end, ok := markdownFindUnescaped(line, 1, ']')
		if !ok || end+1 >= len(line) || line[end+1] != ':' {
			continue
		}
		label := strings.TrimSpace(line[1:end])
		if label == "" || strings.HasPrefix(label, "^") {
			continue
		}
		defs[markdownReferenceLabelKey(label)] = struct{}{}
	}
	return defs
}

func markdownReferenceLabelKey(label string) string {
	return strings.ToLower(strings.Join(strings.Fields(label), " "))
}

func markdownInlineLinkVisibleText(markdown string, referenceDefinitions map[string]struct{}) string {
	var out strings.Builder
	for i := 0; i < len(markdown); {
		start := i
		if markdown[i] == '!' && i+1 < len(markdown) && markdown[i+1] == '[' {
			start = i
			i += 2
		} else if markdown[i] == '[' {
			start = i
			i++
		} else {
			out.WriteByte(markdown[i])
			i++
			continue
		}
		labelStart := i
		labelEnd, ok := markdownFindUnescaped(markdown, i, ']')
		if !ok {
			out.WriteString(markdown[start:i])
			continue
		}
		if labelEnd+1 < len(markdown) && markdown[labelEnd+1] == '(' {
			targetEnd, ok := markdownFindLinkTargetEnd(markdown, labelEnd+2)
			if !ok {
				out.WriteString(markdown[start:i])
				continue
			}
			out.WriteString(markdown[labelStart:labelEnd])
			i = targetEnd + 1
			continue
		}
		if labelEnd+1 < len(markdown) && markdown[labelEnd+1] == '[' {
			refEnd, ok := markdownFindUnescaped(markdown, labelEnd+2, ']')
			if !ok {
				out.WriteString(markdown[start:i])
				continue
			}
			ref := strings.TrimSpace(markdown[labelEnd+2 : refEnd])
			if ref == "" {
				ref = markdown[labelStart:labelEnd]
			}
			if _, ok := referenceDefinitions[markdownReferenceLabelKey(ref)]; !ok {
				out.WriteString(markdown[start:i])
				continue
			}
			out.WriteString(markdown[labelStart:labelEnd])
			i = refEnd + 1
			continue
		}
		if _, ok := referenceDefinitions[markdownReferenceLabelKey(markdown[labelStart:labelEnd])]; ok {
			out.WriteString(markdown[labelStart:labelEnd])
			i = labelEnd + 1
			continue
		}
		out.WriteString(markdown[start:i])
	}
	return out.String()
}

func markdownFindUnescaped(s string, start int, target byte) (int, bool) {
	escaped := false
	for i := start; i < len(s); i++ {
		if escaped {
			escaped = false
			continue
		}
		if s[i] == '\\' {
			escaped = true
			continue
		}
		if s[i] == target {
			return i, true
		}
	}
	return 0, false
}

func markdownFindLinkTargetEnd(s string, start int) (int, bool) {
	escaped := false
	depth := 0
	var quote byte
	for i := start; i < len(s); i++ {
		if escaped {
			escaped = false
			continue
		}
		switch s[i] {
		case '\\':
			escaped = true
		case '\'', '"':
			if quote == 0 {
				quote = s[i]
			} else if quote == s[i] {
				quote = 0
			}
		case '(':
			if quote != 0 {
				continue
			}
			depth++
		case ')':
			if quote != 0 {
				continue
			}
			if depth == 0 {
				return i, true
			}
			depth--
		}
	}
	return 0, false
}

func markdownBackfillCoversLine(visibleMarkdown, line string) bool {
	return markdownBackfillCoverageCoversLine(markdownBackfillCoverageSet(visibleMarkdown), line)
}

func markdownBackfillCoverageSet(visibleMarkdown string) map[string]struct{} {
	coverage := map[string]struct{}{}
	addMarkdownBackfillCoverage(coverage, visibleMarkdown, markdownReferenceDefinitions(visibleMarkdown))
	return coverage
}

func addMarkdownBackfillCoverage(coverage map[string]struct{}, visibleMarkdown string, referenceDefinitions map[string]struct{}) {
	inFence := false
	for _, visibleLine := range strings.Split(visibleMarkdown, "\n") {
		rawLine := visibleLine
		visibleLine = strings.TrimSpace(visibleLine)
		if visibleLine == "" {
			continue
		}
		if markdownFenceLine(visibleLine) {
			inFence = !inFence
			continue
		}
		if inFence || markdownIndentedCodeLine(rawLine) || markdownLooksLikeReferenceDefinition(visibleLine) {
			continue
		}
		addMarkdownBackfillCoverageText(coverage, visibleLine)
		if markdownPlainVisibleLine(visibleLine) {
			continue
		}
		lineText := markdownInlineLinkVisibleText(visibleLine, referenceDefinitions)
		if visibleText := markdownVisibleLineText(lineText); visibleText != "" {
			addMarkdownBackfillCoverageText(coverage, visibleText)
		}
		for _, cell := range markdownVisibleTableCells(lineText) {
			if cell != "" {
				addMarkdownBackfillCoverageText(coverage, cell)
			}
		}
	}
}

func addMarkdownBackfillCoverageText(coverage map[string]struct{}, text string) {
	if text == "" {
		return
	}
	coverage[text] = struct{}{}
	if !markdownBackfillComparableMayDiffer(text) {
		return
	}
	if comparable := markdownBackfillComparableText(text); comparable != "" && comparable != text {
		coverage[comparable] = struct{}{}
	}
}

func markdownBackfillComparableMayDiffer(s string) bool {
	if s == "" {
		return false
	}
	if strings.ContainsAny(s, "[]*_~`<>&\\\t\r\n") || strings.Contains(s, "  ") {
		return true
	}
	if strings.HasPrefix(s, "#") || strings.HasPrefix(s, ">") || strings.HasPrefix(s, "[^") {
		return true
	}
	if strings.HasSuffix(s, `\`) && !strings.HasSuffix(s, `\\`) {
		return true
	}
	return markdownLineStartsWithListMarker(s)
}

func markdownBackfillCoverageCoversLine(coverage map[string]struct{}, line string) bool {
	line = strings.TrimSpace(line)
	if line == "" {
		return true
	}
	if _, ok := coverage[line]; ok {
		return true
	}
	if visibleLine := markdownVisibleLineText(line); visibleLine != "" && visibleLine != line {
		if _, ok := coverage[visibleLine]; ok {
			return true
		}
	}
	if comparable := markdownBackfillComparableText(line); comparable != "" && comparable != line {
		if _, ok := coverage[comparable]; ok {
			return true
		}
	}
	alternateLine := stripLegacyStandaloneNoteMarker(line)
	if alternateLine != "" {
		_, ok := coverage[alternateLine]
		return ok
	}
	return false
}

func markdownBackfillTableTextContainsLine(markdown, line string) bool {
	return markdownBackfillContainmentSet(markdown).tableTextContainsLine(line)
}

type markdownBackfillContainment struct {
	tableRaw                    []string
	tableVisible                []string
	tableComparable             []string
	visibleLines                []string
	tableRawMaxLen              int
	tableVisibleMaxLen          int
	tableComparableMaxLen       int
	visibleMaxLen               int
	tableRawJoined              string
	tableVisibleJoined          string
	tableComparableJoined       string
	visibleJoined               string
	tableRawExact               map[string]struct{}
	tableVisibleExact           map[string]struct{}
	tableComparableExact        map[string]struct{}
	visibleExact                map[string]struct{}
	visibleShortDigits          map[string]struct{}
	shortTableExactBeforeMinLen bool
}

func markdownBackfillContainmentSet(markdown string) markdownBackfillContainment {
	c := markdownBackfillContainment{
		tableRawExact:        map[string]struct{}{},
		tableVisibleExact:    map[string]struct{}{},
		tableComparableExact: map[string]struct{}{},
		visibleExact:         map[string]struct{}{},
	}
	inFence := false
	for _, markdownLine := range strings.Split(markdown, "\n") {
		trimmed := strings.TrimSpace(markdownLine)
		if markdownFenceLine(trimmed) {
			inFence = !inFence
			continue
		}
		if inFence || markdownIndentedCodeLine(markdownLine) || markdownLooksLikeReferenceDefinition(trimmed) {
			continue
		}
		if markdownLikelyTableRow(trimmed) {
			c.tableRaw = append(c.tableRaw, markdownLine)
			c.tableRawExact[markdownLine] = struct{}{}
			if len(markdownLine) > c.tableRawMaxLen {
				c.tableRawMaxLen = len(markdownLine)
			}
			visible := markdownBackfillVisibleText(markdownLine)
			comparable := markdownBackfillComparableFromVisibleText(visible)
			if visible != "" {
				c.tableVisible = append(c.tableVisible, visible)
				c.tableVisibleExact[visible] = struct{}{}
			}
			if len(visible) > c.tableVisibleMaxLen {
				c.tableVisibleMaxLen = len(visible)
			}
			if len(comparable) > c.tableComparableMaxLen {
				c.tableComparableMaxLen = len(comparable)
			}
			if comparable != "" {
				c.tableComparable = append(c.tableComparable, comparable)
				c.tableComparableExact[comparable] = struct{}{}
			}
		}
		if !strings.Contains(trimmed, "[") {
			if visible := markdownVisibleLineText(markdownLine); visible != "" {
				c.visibleLines = append(c.visibleLines, visible)
				if len(visible) > c.visibleMaxLen {
					c.visibleMaxLen = len(visible)
				}
				c.visibleExact[visible] = struct{}{}
			}
		}
	}
	c.tableRawJoined = strings.Join(c.tableRaw, "\n")
	c.tableVisibleJoined = strings.Join(c.tableVisible, "\n")
	c.tableComparableJoined = strings.Join(c.tableComparable, "\n")
	c.visibleJoined = strings.Join(c.visibleLines, "\n")
	maybeEnableMarkdownVisibleShortDigits(&c)
	return c
}

func markdownBackfillVisibleContainmentSet(markdown string) markdownBackfillContainment {
	c := markdownBackfillContainment{
		visibleExact: map[string]struct{}{},
	}
	inFence := false
	for _, markdownLine := range strings.Split(markdown, "\n") {
		trimmed := strings.TrimSpace(markdownLine)
		if markdownFenceLine(trimmed) {
			inFence = !inFence
			continue
		}
		if inFence || markdownIndentedCodeLine(markdownLine) || markdownLooksLikeReferenceDefinition(trimmed) {
			continue
		}
		if !strings.Contains(trimmed, "[") {
			if visible := markdownVisibleLineText(markdownLine); visible != "" {
				c.visibleLines = append(c.visibleLines, visible)
				if len(visible) > c.visibleMaxLen {
					c.visibleMaxLen = len(visible)
				}
				c.visibleExact[visible] = struct{}{}
			}
		}
	}
	c.visibleJoined = strings.Join(c.visibleLines, "\n")
	maybeEnableMarkdownVisibleShortDigits(&c)
	return c
}

func (c markdownBackfillContainment) tableTextContainsLine(line string) bool {
	line = strings.TrimSpace(line)
	if c.shortTableExactBeforeMinLen {
		if _, ok := c.tableRawExact[line]; ok {
			return true
		}
		if _, ok := c.tableVisibleExact[line]; ok {
			return true
		}
	}
	if utf8.RuneCountInString(line) < 12 {
		return false
	}
	if _, ok := c.tableRawExact[line]; ok {
		return true
	}
	if _, ok := c.tableVisibleExact[line]; ok {
		return true
	}
	if len(line) <= c.tableRawMaxLen && strings.Contains(c.tableRawJoined, line) {
		return true
	}
	if len(line) <= c.tableVisibleMaxLen && strings.Contains(c.tableVisibleJoined, line) {
		return true
	}
	lineComparable := markdownBackfillComparableText(line)
	if utf8.RuneCountInString(lineComparable) < 12 {
		return false
	}
	if _, ok := c.tableComparableExact[lineComparable]; ok {
		return true
	}
	if len(lineComparable) > c.tableComparableMaxLen {
		return false
	}
	return strings.Contains(c.tableComparableJoined, lineComparable)
}

func markdownBackfillVisibleLineContainsLine(markdown, line string) bool {
	return markdownBackfillContainmentSet(markdown).visibleLineContainsLine(line)
}

func (c markdownBackfillContainment) visibleLineContainsLine(line string) bool {
	line = strings.TrimSpace(line)
	if utf8.RuneCountInString(line) < 4 {
		return false
	}
	if _, ok := c.visibleExact[line]; ok {
		return true
	}
	if len(line) > c.visibleMaxLen {
		return false
	}
	if isMarkdownVisibleShortDigitLine(line) && c.visibleShortDigits != nil {
		_, ok := c.visibleShortDigits[line]
		return ok
	}
	return strings.Contains(c.visibleJoined, line)
}

func maybeEnableMarkdownVisibleShortDigits(c *markdownBackfillContainment) {
	if len(c.visibleJoined) < minMarkdownVisibleShortDigitBytes || len(c.visibleJoined) > maxMarkdownVisibleShortDigitBytes {
		return
	}
	if len(c.visibleLines) < minMarkdownVisibleShortDigitLines || c.visibleMaxLen > maxMarkdownVisibleShortDigitLineBytes {
		return
	}
	shortDigits := collectMarkdownVisibleShortDigitSubstrings(c.visibleJoined)
	if len(shortDigits) == 0 {
		return
	}
	c.visibleShortDigits = shortDigits
}

func collectMarkdownVisibleShortDigitSubstrings(s string) map[string]struct{} {
	var out map[string]struct{}
	start := -1
	for i := 0; i <= len(s); i++ {
		if i < len(s) && s[i] >= '0' && s[i] <= '9' {
			if start < 0 {
				start = i
			}
			continue
		}
		if start >= 0 {
			if runLen := i - start; runLen >= 4 {
				if out == nil {
					out = make(map[string]struct{}, runLen)
				}
				addMarkdownVisibleShortDigitRun(out, s[start:i])
			}
			start = -1
		}
	}
	return out
}

func addMarkdownVisibleShortDigitRun(dst map[string]struct{}, run string) {
	for width := 4; width <= 7 && width <= len(run); width++ {
		end := len(run) - width
		for i := 0; i <= end; i++ {
			dst[run[i:i+width]] = struct{}{}
		}
	}
}

func isMarkdownVisibleShortDigitLine(s string) bool {
	if len(s) < 4 || len(s) > 7 {
		return false
	}
	for i := 0; i < len(s); i++ {
		if s[i] < '0' || s[i] > '9' {
			return false
		}
	}
	return true
}

func markdownBackfillComparableContains(container, line string) bool {
	line = markdownBackfillComparableText(line)
	if utf8.RuneCountInString(line) < 12 {
		return false
	}
	container = markdownBackfillComparableText(container)
	return container != line && strings.Contains(container, line)
}

func markdownBackfillComparableText(s string) string {
	s = strings.TrimSpace(s)
	if s == "" {
		return ""
	}
	s = markdownBackfillVisibleText(s)
	return markdownBackfillComparableFromVisibleText(s)
}

func markdownBackfillComparableFromVisibleText(s string) string {
	s = strings.TrimSpace(s)
	if s == "" {
		return ""
	}
	s = markdownVisibleLineText(s)
	s = stripMarkdownBackfillEscapedDecimalTokens(s)
	return strings.Join(strings.Fields(s), " ")
}

func stripMarkdownBackfillEscapedDecimalTokens(s string) string {
	if !strings.Contains(s, `\0.`) {
		return s
	}
	var out strings.Builder
	changed := false
	for i := 0; i < len(s); {
		if s[i] == '\\' && markdownEscapedDecimalTokenAt(s, i) {
			if out.Len() == 0 || !unicode.IsSpace(rune(out.String()[out.Len()-1])) {
				out.WriteByte(' ')
			}
			i += len(`\0.`)
			for i < len(s) && s[i] >= '0' && s[i] <= '9' {
				i++
			}
			changed = true
			continue
		}
		out.WriteByte(s[i])
		i++
	}
	if !changed {
		return s
	}
	return out.String()
}

func markdownEscapedDecimalTokenAt(s string, pos int) bool {
	if pos > 0 {
		r, _ := utf8.DecodeLastRuneInString(s[:pos])
		if r != utf8.RuneError && !unicode.IsSpace(r) {
			return false
		}
	}
	if pos+3 >= len(s) || s[pos] != '\\' || s[pos+1] != '0' || s[pos+2] != '.' || s[pos+3] < '0' || s[pos+3] > '9' {
		return false
	}
	i := pos + 4
	for i < len(s) && s[i] >= '0' && s[i] <= '9' {
		i++
	}
	if i < len(s) {
		r, _ := utf8.DecodeRuneInString(s[i:])
		if r != utf8.RuneError && !unicode.IsSpace(r) {
			return false
		}
	}
	return true
}

func markdownLooksLikeReferenceDefinition(line string) bool {
	if !strings.HasPrefix(line, "[") {
		return false
	}
	if strings.HasPrefix(line, "[^") {
		return false
	}
	if i := strings.Index(line, "]:"); i > 1 && i < 256 {
		return true
	}
	end, ok := markdownFindUnescaped(line, 1, ']')
	return ok && end+1 < len(line) && line[end+1] == ':'
}

func stripLegacyStandaloneNoteMarker(line string) string {
	for _, marker := range []string{"[footnote]", "[comment]"} {
		if !strings.HasPrefix(line, marker) {
			continue
		}
		text := strings.TrimSpace(strings.TrimPrefix(line, marker))
		if text == "" || strings.Contains(text, "[footnote]") || strings.Contains(text, "[comment]") {
			return ""
		}
		return text
	}
	return ""
}

func markdownVisibleLineText(line string) string {
	line = strings.TrimSpace(line)
	if markdownPlainVisibleLine(line) {
		return line
	}
	if strings.HasPrefix(line, "#") {
		line = strings.TrimSpace(strings.TrimLeft(line, "#"))
		line = strings.TrimSpace(stripMarkdownHeadingClosingHashes(line))
	}
	line = strings.TrimSpace(stripMarkdownBlockquoteMarker(line))
	line = strings.TrimSpace(stripMarkdownListMarker(line))
	line = strings.TrimSpace(stripMarkdownFootnoteDefinitionMarker(line))
	line = strings.TrimSpace(stripMarkdownFootnoteReferences(line))
	line = strings.TrimSpace(unescapeMarkdownVisibleText(line))
	line = strings.TrimSpace(stripMarkdownInlineWrappers(line))
	line = strings.TrimSpace(stripMarkdownInlineFormatting(line))
	line = strings.TrimSpace(unescapeMarkdownInlineFormattingMarkers(line))
	line = strings.TrimSpace(markdownAutolinkVisibleText(line))
	line = strings.TrimSpace(markdownVisibleHTMLText(line))
	line = strings.TrimSpace(stripMarkdownHardLineBreakMarker(line))
	return line
}

func markdownPlainVisibleLine(line string) bool {
	if line == "" {
		return false
	}
	switch line[0] {
	case '#', '>', '[', '-', '*', '+':
		return false
	}
	if line[0] >= '0' && line[0] <= '9' {
		return false
	}
	if strings.HasSuffix(line, `\`) {
		return false
	}
	return !strings.ContainsAny(line, "*_~`[<>&\\|")
}

func markdownAutolinkVisibleText(line string) string {
	if !strings.Contains(line, "<") {
		return line
	}
	return markdownAutoLinkRE.ReplaceAllString(line, "$1")
}

func markdownVisibleHTMLText(line string) string {
	if !strings.Contains(line, "<") && !strings.Contains(line, "&") {
		return line
	}
	line = markdownVisibleHTMLBreakRE.ReplaceAllString(line, " ")
	line = markdownVisibleHTMLTagRE.ReplaceAllString(line, " ")
	return strings.Join(strings.Fields(html.UnescapeString(line)), " ")
}

func stripMarkdownHeadingClosingHashes(line string) string {
	if !strings.HasSuffix(line, "#") {
		return line
	}
	i := len(line)
	for i > 0 && line[i-1] == '#' {
		i--
	}
	if i == len(line) || i == 0 {
		return line
	}
	if line[i-1] != ' ' && line[i-1] != '\t' {
		return line
	}
	return strings.TrimSpace(line[:i])
}

func stripMarkdownHardLineBreakMarker(line string) string {
	if !strings.HasSuffix(line, `\`) || strings.HasSuffix(line, `\\`) {
		return line
	}
	if len(line) == 3 && line[1] == ':' && isASCIILetter(rune(line[0])) {
		return line
	}
	return strings.TrimSpace(line[:len(line)-1])
}

func stripMarkdownInlineWrappers(line string) string {
	for {
		next := stripOneMarkdownInlineWrapper(strings.TrimSpace(line))
		if next == line {
			return line
		}
		line = next
	}
}

func stripOneMarkdownInlineWrapper(line string) string {
	for _, marker := range []string{"***", "___", "**", "__", "*", "_", "~~", "`"} {
		if len(line) > len(marker)*2 && strings.HasPrefix(line, marker) && strings.HasSuffix(line, marker) {
			inner := strings.TrimSpace(line[len(marker) : len(line)-len(marker)])
			if inner != "" {
				return inner
			}
		}
	}
	return line
}

func stripMarkdownInlineFormatting(line string) string {
	if !strings.ContainsAny(line, "*_~`") {
		return line
	}
	for _, marker := range []string{"***", "___", "**", "__", "~~", "`", "*", "_"} {
		line = stripMarkdownInlineFormattingMarker(line, marker)
	}
	return line
}

func stripMarkdownInlineFormattingMarker(line, marker string) string {
	var out strings.Builder
	for i := 0; i < len(line); {
		start := strings.Index(line[i:], marker)
		if start < 0 {
			out.WriteString(line[i:])
			break
		}
		start += i
		if markdownEscapedAt(line, start) || !markdownInlineMarkerBoundary(line, marker, start, true) {
			out.WriteString(line[i : start+len(marker)])
			i = start + len(marker)
			continue
		}
		endSearch := start + len(marker)
		end := -1
		for endSearch < len(line) {
			candidate := strings.Index(line[endSearch:], marker)
			if candidate < 0 {
				break
			}
			candidate += endSearch
			if !markdownEscapedAt(line, candidate) && markdownInlineMarkerBoundary(line, marker, candidate, false) {
				end = candidate
				break
			}
			endSearch = candidate + len(marker)
		}
		if end < 0 || end == start+len(marker) {
			out.WriteString(line[i : start+len(marker)])
			i = start + len(marker)
			continue
		}
		out.WriteString(line[i:start])
		out.WriteString(line[start+len(marker) : end])
		i = end + len(marker)
	}
	return out.String()
}

func markdownEscapedAt(s string, pos int) bool {
	backslashes := 0
	for i := pos - 1; i >= 0 && s[i] == '\\'; i-- {
		backslashes++
	}
	return backslashes%2 == 1
}

func markdownInlineMarkerBoundary(s, marker string, pos int, opening bool) bool {
	if marker == "`" || marker == "~~" {
		return true
	}
	before := byte(0)
	after := byte(0)
	if pos > 0 {
		before = s[pos-1]
	}
	if pos+len(marker) < len(s) {
		after = s[pos+len(marker)]
	}
	if opening {
		if after == 0 || after == ' ' || after == '\t' {
			return false
		}
		if marker == "_" && markdownASCIIAlnum(before) && markdownASCIIAlnum(after) {
			return false
		}
		return true
	}
	if before == 0 || before == ' ' || before == '\t' {
		return false
	}
	if marker == "_" && markdownASCIIAlnum(before) && markdownASCIIAlnum(after) {
		return false
	}
	return true
}

func markdownASCIIAlnum(b byte) bool {
	return (b >= '0' && b <= '9') || (b >= 'A' && b <= 'Z') || (b >= 'a' && b <= 'z')
}

func unescapeMarkdownInlineFormattingMarkers(s string) string {
	s = strings.ReplaceAll(s, `\*`, "*")
	s = strings.ReplaceAll(s, `\_`, "_")
	s = strings.ReplaceAll(s, `\~`, "~")
	s = strings.ReplaceAll(s, "\\`", "`")
	return s
}

func stripMarkdownBlockquoteMarker(line string) string {
	for strings.HasPrefix(line, ">") {
		line = strings.TrimSpace(line[1:])
	}
	return line
}

func stripMarkdownListMarker(line string) string {
	if len(line) >= 2 && (line[0] == '-' || line[0] == '*' || line[0] == '+') && unicode.IsSpace(rune(line[1])) {
		return stripMarkdownTaskMarker(strings.TrimSpace(line[2:]))
	}
	i := 0
	for i < len(line) && line[i] >= '0' && line[i] <= '9' {
		i++
	}
	if i > 0 && i+1 < len(line) && (line[i] == '.' || line[i] == ')') && unicode.IsSpace(rune(line[i+1])) {
		return stripMarkdownTaskMarker(strings.TrimSpace(line[i+2:]))
	}
	return line
}

func stripMarkdownTaskMarker(line string) string {
	if len(line) >= 4 && line[0] == '[' && line[2] == ']' && unicode.IsSpace(rune(line[3])) {
		switch line[1] {
		case ' ', 'x', 'X':
			return strings.TrimSpace(line[4:])
		}
	}
	return line
}

func stripMarkdownFootnoteDefinitionMarker(line string) string {
	if len(line) < 5 || line[0] != '[' || line[1] != '^' {
		return line
	}
	end := strings.IndexByte(line, ']')
	if end <= 2 || end+1 >= len(line) || line[end+1] != ':' {
		return line
	}
	return strings.TrimSpace(line[end+2:])
}

func stripMarkdownFootnoteReferences(line string) string {
	if !strings.Contains(line, "[^") {
		return line
	}
	var out strings.Builder
	changed := false
	for i := 0; i < len(line); {
		if line[i] == '[' && i+2 < len(line) && line[i+1] == '^' && !markdownEscapedAt(line, i) {
			end, ok := markdownFindUnescaped(line, i+2, ']')
			if ok && end > i+2 {
				changed = true
				i = end + 1
				continue
			}
		}
		out.WriteByte(line[i])
		i++
	}
	if !changed {
		return line
	}
	return out.String()
}

func markdownVisibleTableCells(line string) []string {
	return markdownVisibleTableCellsWithCache(line, nil)
}

func markdownVisibleTableCellsWithCache(line string, cellCache map[string][]string) []string {
	trimmed := strings.TrimSpace(line)
	if !markdownLikelyTableRow(trimmed) {
		return nil
	}
	raw := strings.Trim(trimmed, "|")
	parts := splitMarkdownTableRow(raw)
	out := make([]string, 0, len(parts))
	for _, part := range parts {
		out = append(out, markdownVisibleTableCellVariants(part, cellCache)...)
	}
	return out
}

func markdownVisibleTableCellVariants(part string, cellCache map[string][]string) []string {
	cacheable := cellCache != nil && len(part) <= 8
	if cacheable {
		if cached, ok := cellCache[part]; ok {
			return cached
		}
	}
	original := part
	part = strings.TrimSpace(unescapeMarkdownVisibleText(part))
	part = strings.TrimSpace(stripMarkdownFootnoteReferences(part))
	part = strings.TrimSpace(stripMarkdownInlineWrappers(part))
	part = strings.TrimSpace(stripMarkdownInlineFormatting(part))
	part = strings.TrimSpace(unescapeMarkdownInlineFormattingMarkers(part))
	part = strings.TrimSpace(markdownAutolinkVisibleText(part))
	part = strings.TrimSpace(stripMarkdownHardLineBreakMarker(part))
	if part == "" || markdownTableSeparatorCell(part) {
		if cacheable {
			cellCache[original] = nil
		}
		return nil
	}
	out := []string{part}
	if strings.IndexByte(part, '<') >= 0 {
		visiblePart := strings.TrimSpace(markdownVisibleHTMLText(part))
		if visiblePart != "" && visiblePart != part {
			out = append(out, visiblePart)
		}
	}
	if strings.Contains(part, "<br>") {
		for _, segment := range strings.Split(part, "<br>") {
			segment = strings.TrimSpace(segment)
			segment = strings.TrimSpace(stripMarkdownFootnoteReferences(segment))
			segment = strings.TrimSpace(stripMarkdownInlineWrappers(segment))
			segment = strings.TrimSpace(stripMarkdownInlineFormatting(segment))
			segment = strings.TrimSpace(unescapeMarkdownInlineFormattingMarkers(segment))
			segment = strings.TrimSpace(markdownAutolinkVisibleText(segment))
			segment = strings.TrimSpace(markdownVisibleHTMLText(segment))
			segment = strings.TrimSpace(stripMarkdownHardLineBreakMarker(segment))
			if segment != "" && segment != part {
				out = append(out, segment)
			}
		}
	}
	if cacheable {
		cellCache[original] = out
	}
	return out
}

func splitMarkdownTableRow(row string) []string {
	parts := make([]string, 0, strings.Count(row, "|")+1)
	start := 0
	for i := 0; i < len(row); i++ {
		switch row[i] {
		case '\\':
			if i+1 < len(row) {
				i++
			}
		case '|':
			parts = append(parts, row[start:i])
			start = i + 1
		}
	}
	parts = append(parts, row[start:])
	return parts
}

func unescapeMarkdownVisibleText(s string) string {
	s = strings.ReplaceAll(s, `\|`, "|")
	s = strings.ReplaceAll(s, `\[`, "[")
	s = strings.ReplaceAll(s, `\]`, "]")
	s = strings.ReplaceAll(s, `\#`, "#")
	s = strings.ReplaceAll(s, `\-`, "-")
	s = strings.ReplaceAll(s, `\+`, "+")
	s = strings.ReplaceAll(s, `\.`, ".")
	s = strings.ReplaceAll(s, `\(`, "(")
	s = strings.ReplaceAll(s, `\)`, ")")
	s = strings.ReplaceAll(s, `\!`, "!")
	s = strings.ReplaceAll(s, `\\`, `\`)
	return s
}

func markdownTableSeparatorCell(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	for _, r := range s {
		if r != '-' && r != ':' {
			return false
		}
	}
	return true
}

func markdownBackfillLineAllowed(line string, imageAlts map[string]bool) bool {
	if len(line) > maxMarkdownTableCellBytes {
		return false
	}
	key := cleanText(line)
	key = stripInlineHiddenOfficeReferences(key)
	return markdownBackfillNormalizedLineAllowed(strings.TrimSpace(key), imageAlts)
}

func markdownBackfillNormalizedLineAllowed(line string, imageAlts map[string]bool) bool {
	if len(line) > maxMarkdownTableCellBytes {
		return false
	}
	key := strings.TrimSpace(line)
	if key == "" || markdownTableSeparatorCell(key) || markdownBackfillLowInformationLine(key) {
		return false
	}
	if maybeDiscardableHiddenOfficeText(key) && (looksLikeMarkdownBackfillHiddenReference(key) || looksLikeRelationshipIDReference(key) || looksLikeOfficeRelationshipMetadataReference(key) || looksLikeOfficeXMLMetadataReference(key)) {
		return false
	}
	return !imageAlts[key]
}

func markdownBackfillLowInformationLine(line string) bool {
	line = strings.TrimSpace(line)
	if isMarkdownSingleASCIIWordCharLine(line) {
		return true
	}
	if line == "" || len(line) > 8 {
		switch strings.ToLower(line) {
		case "self-employment calculation":
			return true
		default:
			return false
		}
	}
	switch strings.ToLower(line) {
	case "use", "apply", "clear":
		return true
	}
	var digit rune
	count := 0
	for _, r := range line {
		if !unicode.IsDigit(r) {
			return false
		}
		if count == 0 {
			digit = r
		} else if r != digit {
			return false
		}
		count++
	}
	return count >= 3
}

func looksLikeMarkdownBackfillHiddenReference(s string) bool {
	trimmed := strings.TrimSpace(strings.ReplaceAll(s, "\\", "/"))
	if trimmed == "" {
		return false
	}
	if looksLikeInlineHiddenResourceReferencePlain(trimmed) {
		return true
	}
	return looksLikeDecodedOfficePartPath(trimmed)
}

func markdownImageAltSet(images []Image) map[string]bool {
	out := map[string]bool{}
	names := imageOutputFilenames(images)
	for i, img := range images {
		name := img.Name
		if name == "" {
			name = fmt.Sprintf("image-%03d%s", i+1, img.Ext)
		}
		if img.Ext != "" {
			name = imageNameWithExt(name, strings.ToLower(img.Ext))
		}
		outputName := ""
		if i < len(names) {
			outputName = names[i]
		}
		for _, value := range []string{
			img.Alt,
			markdownImageAlt(img, outputName, i),
			name,
			strings.TrimSuffix(name, path.Ext(name)),
			outputName,
			strings.TrimSuffix(outputName, path.Ext(outputName)),
		} {
			value = cleanText(value)
			value = stripInlineHiddenOfficeReferences(value)
			if value != "" {
				out[value] = true
			}
		}
	}
	return out
}

func markdownText(s string) string {
	if text, ok := rtfVisibleText(s); ok {
		s = text
	}
	s = strings.ReplaceAll(s, "\r\n", "\n")
	s = strings.ReplaceAll(s, "\r", "\n")
	lines := strings.Split(s, "\n")
	out := make([]string, 0, len(lines))
	blank := false
	for _, line := range lines {
		line = normalizeMarkdownTextLine(line)
		if line == "" {
			if len(out) > 0 {
				blank = true
			}
			continue
		}
		if blank {
			out = append(out, "")
			blank = false
		}
		out = append(out, line)
	}
	out = collapseMarkdownSingleCharacterLines(out)
	return strings.Trim(strings.Join(out, "\n"), "\n")
}

func markdownBackfillSourceText(s string) string {
	return strings.Trim(strings.Join(markdownBackfillSourceLines(markdownBackfillRawLines(s)), "\n"), "\n")
}

func normalizeMarkdownTextLine(line string) string {
	trimmed := strings.TrimSpace(line)
	if trimmed == "" {
		return ""
	}
	if markdownLineStartsWithListMarker(trimmed) {
		indent := markdownLeadingSpaces(line)
		if indent > 0 {
			if indent > 12 {
				indent = 12
			}
			return strings.Repeat(" ", indent) + trimmed
		}
	}
	return trimmed
}

func markdownLineStartsWithListMarker(line string) bool {
	return stripMarkdownListMarker(line) != line
}

func markdownLeadingSpaces(line string) int {
	n := 0
	for n < len(line) && line[n] == ' ' {
		n++
	}
	return n
}

func collapseMarkdownSingleCharacterLines(lines []string) []string {
	if len(lines) < 3 {
		return lines
	}
	out := make([]string, 0, len(lines))
	var run []string
	flush := func() {
		if shouldCollapseMarkdownSingleCharacterRun(run) {
			out = append(out, strings.Join(run, ""))
		} else {
			out = append(out, run...)
		}
		run = run[:0]
	}
	for _, line := range lines {
		if isMarkdownSingleASCIIWordCharLine(line) {
			run = append(run, line)
			continue
		}
		flush()
		out = append(out, line)
	}
	flush()
	return out
}

func shouldCollapseMarkdownSingleCharacterRun(run []string) bool {
	if len(run) < 3 {
		return false
	}
	if len(run) >= 5 {
		return true
	}
	for _, s := range run {
		if len(s) == 1 && s[0] >= 'a' && s[0] <= 'z' {
			return true
		}
	}
	return false
}

func isMarkdownSingleASCIIWordCharLine(s string) bool {
	if len(s) != 1 {
		return false
	}
	c := s[0]
	return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9')
}

func markdownImageTarget(base, name string) string {
	base = strings.TrimSpace(markdownSlashPath(base))
	name = markdownSlashPath(name)
	if base == "" || base == "." {
		return name
	}
	if u, err := url.Parse(base); err == nil && u.Scheme != "" && u.Host != "" {
		prefix := u.Scheme + "://"
		if u.User != nil {
			prefix += u.User.String() + "@"
		}
		prefix += u.Host
		out := prefix + strings.TrimRight(u.EscapedPath(), "/") + "/" + escapeMarkdownPath(strings.TrimLeft(name, "/"))
		if u.RawQuery != "" {
			out += "?" + u.RawQuery
		}
		if u.Fragment != "" {
			out += "#" + u.EscapedFragment()
		}
		return out
	}
	return strings.TrimRight(base, "/") + "/" + strings.TrimLeft(name, "/")
}

func markdownImageAlt(img Image, name string, index int) string {
	if alt := cleanMarkdownImageAltText(img.Alt); alt != "" {
		return alt
	}
	base := strings.TrimSuffix(name, path.Ext(name))
	base = strings.TrimSpace(strings.ReplaceAll(base, "_", " "))
	if alt := cleanMarkdownImageAltText(base); alt != "" {
		return alt
	}
	return fmt.Sprintf("image %d", index+1)
}

func cleanMarkdownImageAltText(s string) string {
	s = cleanText(s)
	s = stripInlineHiddenOfficeReferences(s)
	if s == "" || looksLikeMarkdownImageAltHiddenReference(s) || looksLikeRelationshipIDReference(s) || looksLikeOfficeRelationshipMetadataReference(s) || looksLikeOfficeXMLMetadataReference(s) || looksLikeBinaryControlFragment(s) {
		return ""
	}
	return s
}

func looksLikeMarkdownImageAltHiddenReference(s string) bool {
	raw := strings.TrimSpace(strings.ReplaceAll(s, "\\", "/"))
	if raw == "" {
		return false
	}
	seen := map[string]bool{}
	queue := []string{raw}
	for len(queue) > 0 && len(seen) < 32 {
		trimmed := strings.TrimSpace(strings.ReplaceAll(queue[0], "\\", "/"))
		queue = queue[1:]
		if trimmed == "" || seen[trimmed] {
			continue
		}
		seen[trimmed] = true
		if looksLikeInlineHiddenResourceReferencePlain(trimmed) {
			return true
		}
		if looksLikeDecodedOfficePartPath(trimmed) {
			return true
		}
		if len(trimmed) >= 3 && isASCIILetter(rune(trimmed[0])) && trimmed[1] == ':' && trimmed[2] == '/' && isSupportedImageExt(path.Ext(trimmed)) {
			return true
		}
		if normalized := hiddenResourceReferenceCandidate(trimmed); normalized != trimmed {
			queue = append(queue, normalized)
		}
		if decoded, err := url.PathUnescape(trimmed); err == nil && decoded != trimmed {
			queue = append(queue, decoded)
		}
		if strings.Contains(trimmed, "&") {
			if unescaped := html.UnescapeString(trimmed); unescaped != trimmed {
				queue = append(queue, unescaped)
			}
		}
	}
	return false
}

func escapeMarkdownImageAlt(s string) string {
	s = strings.Join(strings.Fields(s), " ")
	s = strings.ReplaceAll(s, "\\", "\\\\")
	s = strings.ReplaceAll(s, "[", "\\[")
	s = strings.ReplaceAll(s, "]", "\\]")
	return s
}

func escapeMarkdownLinkTarget(s string) string {
	s = markdownSlashPath(s)
	if u, err := url.Parse(s); err == nil && u.Scheme != "" && u.Host != "" {
		prefix := u.Scheme + "://"
		if u.User != nil {
			prefix += u.User.String() + "@"
		}
		prefix += u.Host
		out := prefix + escapeEscapedMarkdownPath(u.EscapedPath())
		if u.RawQuery != "" {
			out += "?" + strings.ReplaceAll(strings.ReplaceAll(u.RawQuery, "(", "%28"), ")", "%29")
		}
		if u.Fragment != "" {
			out += "#" + escapeMarkdownPathSegment(u.Fragment)
		}
		return out
	}
	return escapeMarkdownPath(s)
}

func markdownSlashPath(s string) string {
	return strings.ReplaceAll(filepath.ToSlash(s), "\\", "/")
}

func escapeEscapedMarkdownPath(s string) string {
	s = strings.ReplaceAll(s, "(", "%28")
	s = strings.ReplaceAll(s, ")", "%29")
	return s
}

func escapeMarkdownPath(s string) string {
	parts := strings.Split(s, "/")
	for i, part := range parts {
		if i == 0 && isWindowsDrivePathSegment(part) {
			parts[i] = part
			continue
		}
		parts[i] = escapeMarkdownPathSegmentPreserveEscapes(part)
	}
	return strings.Join(parts, "/")
}

func isWindowsDrivePathSegment(s string) bool {
	if len(s) != 2 || s[1] != ':' {
		return false
	}
	c := s[0]
	return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')
}

func escapeMarkdownPathSegment(s string) string {
	s = url.PathEscape(s)
	s = strings.ReplaceAll(s, "(", "%28")
	s = strings.ReplaceAll(s, ")", "%29")
	return s
}

func escapeMarkdownPathSegmentPreserveEscapes(s string) string {
	var b strings.Builder
	for i := 0; i < len(s); {
		if s[i] == '%' && i+2 < len(s) && isHexDigit(s[i+1]) && isHexDigit(s[i+2]) {
			b.WriteString(s[i : i+3])
			i += 3
			continue
		}
		r, size := utf8.DecodeRuneInString(s[i:])
		if r == utf8.RuneError && size == 0 {
			break
		}
		b.WriteString(url.PathEscape(string(r)))
		i += size
	}
	out := b.String()
	out = strings.ReplaceAll(out, "(", "%28")
	out = strings.ReplaceAll(out, ")", "%29")
	return out
}

func isHexDigit(c byte) bool {
	return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F')
}

func uniqueImageFilename(name string, used map[string]bool) string {
	name = truncateFilenameBytes(name, maxImageFilenameBytes)
	key := strings.ToLower(name)
	if !used[key] {
		used[key] = true
		return name
	}
	ext := filepath.Ext(name)
	base := strings.TrimSuffix(name, ext)
	for i := 2; ; i++ {
		suffix := fmt.Sprintf("-%d", i)
		candidate := truncateFilenameBaseBytes(base, ext, maxImageFilenameBytes-len(suffix)) + suffix + ext
		key = strings.ToLower(candidate)
		if !used[key] {
			used[key] = true
			return candidate
		}
	}
}

func isZip(data []byte) bool {
	return len(data) >= 4 && bytes.Equal(data[:4], []byte{'P', 'K', 3, 4})
}

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
		texts, err = extractDocxText(files)
	case "pptx":
		texts, err = extractPptxText(files)
	case "xlsx":
		text, xlsxMarkdown, err = extractXlsxText(files)
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
	images, err := extractOOXMLImages(files, kind, opts.IncludeMetadata)
	if err != nil {
		return nil, err
	}
	if kind == "docx" {
		images = append(images, extractDocxAltChunkMHTMLImages(files)...)
		images = append(images, extractDocxAltChunkHTMLDataImages(files)...)
	}
	var embeddedText []string
	var embeddedMarkdown []string
	if depth < maxOOXMLEmbeddedDepth {
		var embeddedImages []Image
		embeddedText, embeddedMarkdown, embeddedImages = extractEmbeddedOfficePackages(files, kind, depth+1, opts)
		texts = append(texts, embeddedText...)
		images = append(images, embeddedImages...)
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
	} else {
		text = joinText(texts)
	}
	return &Result{Text: strings.TrimSpace(text), StructuredMarkdown: structuredMarkdown, Images: images}, nil
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

func extractDocxText(files map[string]*zip.File) ([]string, error) {
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
			if base == "document.xml" || base == "footnotes.xml" || base == "endnotes.xml" || base == "comments.xml" ||
				(isDocxHeaderFooterPart(lower) && (!constrainedHeaderFooter || visibleHeaderFooter[lower])) {
				if base == "footnotes.xml" || base == "endnotes.xml" || base == "comments.xml" {
					continue
				}
				names = append(names, name)
			}
			if docxRelatedTextPart(lower) && (!constrainedRelated || visibleRelated[lower]) {
				names = append(names, name)
			}
		}
		if docxVisibleVMLPart(files, lower) {
			names = append(names, name)
		}
	}
	sort.Strings(names)
	out, err := xmlTextFromFiles(files, names)
	if err != nil {
		return nil, err
	}
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

func extractPptxText(files map[string]*zip.File) ([]string, error) {
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
	if len(names) == 0 {
		return nil, false, nil
	}
	return names, true, nil
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

func extractXlsxText(files map[string]*zip.File) (string, map[string]xlsxWorksheetMarkdownData, error) {
	shared, err := readSharedStrings(files)
	if err != nil {
		return "", nil, err
	}
	var out strings.Builder
	markdown := map[string]xlsxWorksheetMarkdownData{}
	workbookTexts, sheetNames, err := workbookTextAndSheets(files)
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
		if err := appendWorksheetText(&out, b, shared, &md); err != nil {
			return "", nil, err
		}
		markdown[ooxmlPartKey(name)] = md
	}
	if xlsxHasAnyPartPrefix(files, []string{"xl/charts/", "xl/drawings/", "xl/tables/", "xl/pivottables/", "xl/pivotcache/", "xl/slicers/", "xl/slicercaches/"}) {
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
	if xlsxHasAnyPartPrefix(files, []string{"xl/comments", "xl/threadedcomments"}) {
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
	text, _, err := workbookTextAndSheets(files)
	return text, err
}

func workbookTextAndSheets(files map[string]*zip.File) ([]string, []string, error) {
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
					if sheetName != "" {
						out = append(out, sheetName)
					}
					if hasRels {
						target := rels[relID]
						if actual := ooxmlPartName(files, target); actual != "" {
							sheets = append(sheets, actual)
							sheetIndex++
							break
						}
					}
					if len(fallbackParts) > sheetIndex {
						sheets = append(sheets, fallbackParts[sheetIndex])
					}
				}
				sheetIndex++
			case "definedName":
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
			if contentVisible {
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
	choiceSeen bool
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
	return vmlElementHidden(start)
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
	return strings.Contains(font, "symbol") || strings.Contains(font, "wingdings")
}

func fontEncodedSymbolRune(font string, code rune) (rune, bool) {
	switch {
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

func appendWorksheetText(out *strings.Builder, b []byte, shared []string, md *xlsxWorksheetMarkdownData) error {
	if ok, err := appendSimpleInlineWorksheetText(out, b, md); ok || err != nil {
		return err
	}
	if ok, err := appendSharedStringWorksheetText(out, b, shared, md); ok || err != nil {
		return err
	}
	dec := xml.NewDecoder(bytes.NewReader(b))
	var cellType string
	var inV, inT bool
	var inHeaderFooter bool
	var rowHidden bool
	var skipCell bool
	var collectMarkdownRow bool
	var collectMarkdownCell bool
	var markdownRowValues []string
	var hiddenCols []intRange
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
				cellRef := ""
				collectMarkdownCell = false
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
				skipCell = rowHidden || hiddenColumnCell(cellRef, hiddenCols) || columnHidden(cellCol, hiddenCols)
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
			if !skipCell && worksheetElementMayHaveVisibleTextAttributes(t.Name.Local) && !worksheetElementHiddenByRef(t, hiddenCols, hiddenRows) {
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
				value := cleanExcelHeaderFooterText(cur.String())
				if value != "" {
					appendCleanedTextBlock(out, value)
					if md != nil {
						md.headerFooter = append(md.headerFooter, value)
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
							if len(shared[idx]) > maxRepeatedTextPartBytes {
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
					if !skipValue {
						if t.Name.Local == "v" && cellType == "" && plainExcelNumberValue(value) {
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

func extractOOXMLImages(files map[string]*zip.File, kind string, includeMetadata bool) ([]Image, error) {
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
	for name := range files {
		if isOOXMLMediaPart(name, prefix, kind) || (includeMetadata && isOOXMLThumbnail(name)) {
			names = append(names, name)
		}
	}
	visibleMedia, filterMedia := visibleOOXMLMediaParts(files, kind)
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
	if actual := ooxmlPartName(files, part); actual != "" {
		part = actual
	}
	if strings.HasPrefix(ooxmlPartKey(part), "word/media/") {
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
	return refs, nil
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
	if len(hidden) == 0 {
		return visible, true
	}
	for name := range hidden {
		if !visible[name] {
			delete(visible, name)
		}
	}
	return visible, true
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
	refs, err := docxImageRelationshipRefs(b)
	if err != nil || (len(refs.Visible) == 0 && len(refs.Hidden) == 0) {
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
		target := cleanText(xmlAttrValue(start, "Target"))
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
		target := cleanText(xmlAttrValue(start, "Target"))
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
				if streams, streamErr := readOLEStreams(b); streamErr == nil {
					res = extractOfficePackagesFromOLEStreams(name, streams, depth, opts)
					if res != nil {
						break
					}
				}
				for _, img := range extractNonOfficeOLEImages(name, b, len(images)) {
					if img.Name != "" {
						img.Name = uniqueImageFilename(img.Name, usedImageNames)
					}
					images = append(images, img)
				}
				continue
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

func extractLegacy(filename string, data []byte, opts Options) (*Result, error) {
	return extractLegacyWithDepth(filename, data, 0, opts)
}

func extractLegacyWithDepth(filename string, data []byte, depth int, opts Options) (*Result, error) {
	ext := strings.ToLower(filepath.Ext(filename))
	streams, err := readOLEStreams(data)
	streamReadFailed := err != nil && isOLE(data)
	if err != nil {
		streams = nil
	}
	if legacyOLEEncryptedPackage(streams) {
		if ext == ".xls" {
			if parts := biffBoundSheetNames(data, 1252); len(parts) > 0 {
				return &Result{Text: cleanVisibleText(strings.Join(parts, "\n"))}, nil
			}
		}
		return &Result{}, nil
	}
	if !streamReadFailed && inferLegacyExt(streams) == "" && depth < maxOOXMLEmbeddedDepth {
		if res := extractOfficePackagesFromOLEStreams(filename, streams, depth+1, opts); res != nil {
			return res, nil
		}
	}
	if ext == ".xls" && legacyXLSEncrypted(streams) {
		return nil, errors.New("encrypted legacy Excel document is not supported")
	}
	var textParts []string
	var xlsWorkbook []byte
	var xlsWorkbookText []string
	if streamReadFailed {
		if ext == ".xls" {
			textParts = biffBoundSheetNames(data, 1252)
		}
		if len(textParts) == 0 {
			textParts = extractCorruptOLEText(data)
		}
	} else {
		if ext == ".xls" {
			xlsWorkbook = legacyWorkbookBytes(data, streams)
			if len(xlsWorkbook) > 0 {
				xlsWorkbookText = biffText(xlsWorkbook)
				if len(xlsWorkbookText) > 0 {
					textParts = xlsLegacyTextPartsFromWorkbookData(data, streams, opts.IncludeMetadata, xlsWorkbookText)
				}
			}
		}
		if len(textParts) == 0 {
			textParts = extractLegacyTextWithMetadata(filename, data, streams, opts.IncludeMetadata)
		}
	}
	var text string
	if ext == ".xls" && !streamReadFailed {
		text = joinCleanedText(textParts)
	} else {
		text = cleanVisibleText(strings.Join(textParts, "\n"))
	}
	var structuredMarkdown string
	if ext == ".xls" && !streamReadFailed && len(xlsWorkbook) > 0 {
		if len(xlsWorkbookText) > 0 {
			structuredMarkdown = biffMarkdownWithText(xlsWorkbook, xlsWorkbookText)
		} else {
			structuredMarkdown = biffMarkdown(xlsWorkbook)
		}
	} else if ext == ".doc" && !opts.IncludeMetadata && !streamReadFailed {
		structuredMarkdown = legacyWordMarkdown(textParts)
	} else {
		structuredMarkdown = extractLegacyMarkdown(filename, data, streams)
	}
	if structuredMarkdown == "" {
		structuredMarkdown = legacyFallbackMarkdown(filename, textParts)
	}
	images := imagesFromOLEStreams(data, streams)
	return &Result{Text: text, StructuredMarkdown: structuredMarkdown, Images: images}, nil
}

func extractOfficePackagesFromOLEStreams(filename string, streams []oleStream, depth int, opts Options) *Result {
	payloads := embeddedOfficePackagePayloadsFromStreams(streams)
	if len(payloads) == 0 {
		return nil
	}
	var texts []string
	var markdowns []string
	var images []Image
	usedImageNames := map[string]bool{}
	for _, payload := range payloads {
		res, err := extractOfficePackagePayload(payload.name, payload.data, depth, opts)
		if err != nil || res == nil {
			continue
		}
		if res.Text != "" {
			texts = append(texts, res.Text)
		}
		if res.StructuredMarkdown != "" {
			markdowns = append(markdowns, res.StructuredMarkdown)
		}
		for _, img := range res.Images {
			if img.Name != "" {
				img.Name = sanitizeFilename(oleStreamBaseName(payload.name) + "-" + img.Name)
			}
			if img.Name != "" {
				img.Name = uniqueImageFilename(img.Name, usedImageNames)
			}
			images = append(images, img)
		}
	}
	text := joinText(texts)
	markdown := strings.TrimSpace(strings.Join(markdowns, "\n\n"))
	if markdown == "" && text != "" {
		markdown = legacyFallbackMarkdown(filename, texts)
	}
	if text == "" && markdown == "" && len(images) == 0 {
		return nil
	}
	return &Result{Text: text, StructuredMarkdown: markdown, Images: images}
}

type officePackagePayload struct {
	name string
	data []byte
}

func embeddedOfficePackagePayloadsFromStreams(streams []oleStream) []officePackagePayload {
	seen := map[string]bool{}
	var out []officePackagePayload
	for _, s := range streams {
		if !officePackagePayloadStreamName(s.Name, s.Path) {
			continue
		}
		for _, payload := range officePackagePayloadsFromBytes(s.Name, s.Data) {
			key := officePackagePayloadKind(payload.data) + "\x00" + string(payload.data)
			if seen[key] {
				continue
			}
			seen[key] = true
			out = append(out, payload)
		}
	}
	return out
}

func officePackagePayloadStreamName(name, streamPath string) bool {
	for _, value := range []string{name, streamPath} {
		base := strings.ToLower(path.Base(filepath.ToSlash(value)))
		switch base {
		case "package", "contents", "ole10native":
			return true
		}
		if strings.HasPrefix(base, "\x01ole10native") {
			return true
		}
	}
	return false
}

func officePackagePayloadsFromBytes(name string, data []byte) []officePackagePayload {
	var out []officePackagePayload
	add := func(payload []byte) bool {
		if len(payload) == 0 {
			return false
		}
		if zipPayload, ok := normalizeOfficePackageZipPayload(payload); ok {
			out = append(out, officePackagePayload{name: officePackagePayloadName(name, zipPayload), data: zipPayload})
			return true
		}
		if isOLE(payload) {
			streams, err := readOLEStreams(payload)
			if err == nil && inferLegacyExt(streams) != "" {
				out = append(out, officePackagePayload{name: officePackagePayloadName(name, payload), data: payload})
				return true
			}
		}
		return false
	}
	if add(data) {
		return out
	}
	for _, off := range officePayloadMagicOffsets(data, []byte{'P', 'K', 3, 4}) {
		if add(data[off:]) {
			return out
		}
	}
	for _, off := range officePayloadMagicOffsets(data, []byte{0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1}) {
		if add(data[off:]) {
			return out
		}
	}
	return out
}

func normalizeOfficePackageZipPayload(data []byte) ([]byte, bool) {
	if !isZip(data) {
		return nil, false
	}
	if isOOXMLPackage(data) {
		return data, true
	}
	for _, end := range zipPayloadEndOffsets(data) {
		payload := data[:end]
		if isOOXMLPackage(payload) {
			return payload, true
		}
	}
	return nil, false
}

func zipPayloadEndOffsets(data []byte) []int {
	const eocdLen = 22
	var ends []int
	for pos := len(data) - eocdLen; pos >= 0; pos-- {
		if binary.LittleEndian.Uint32(data[pos:]) != 0x06054b50 {
			continue
		}
		commentLen := int(binary.LittleEndian.Uint16(data[pos+20:]))
		end := pos + eocdLen + commentLen
		if end >= pos+eocdLen && end <= len(data) {
			ends = append(ends, end)
		}
	}
	return ends
}

func officePayloadMagicOffsets(data, magic []byte) []int {
	var out []int
	for offset := 0; offset < len(data); {
		i := bytes.Index(data[offset:], magic)
		if i < 0 {
			break
		}
		pos := offset + i
		if pos > 0 {
			out = append(out, pos)
		}
		offset = pos + len(magic)
	}
	return out
}

func officePackagePayloadName(name string, data []byte) string {
	base := path.Base(filepath.ToSlash(name))
	if base == "" || base == "." {
		base = "embedded-office"
	}
	kind := officePackagePayloadKind(data)
	if kind != "" {
		return imageNameWithExt(base, "."+kind)
	}
	return base
}

func officePackagePayloadKind(data []byte) string {
	if isZip(data) {
		zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			return ""
		}
		files := map[string]*zip.File{}
		for _, f := range zr.File {
			files[f.Name] = f
		}
		return ooxmlKind(files)
	}
	if isOLE(data) {
		streams, err := readOLEStreams(data)
		if err != nil {
			return ""
		}
		return strings.TrimPrefix(inferLegacyExt(streams), ".")
	}
	return ""
}

func extractOfficePackagePayload(name string, data []byte, depth int, opts Options) (*Result, error) {
	if isZip(data) && isOOXMLPackage(data) {
		return extractOOXMLWithDepth(name, data, depth, opts)
	}
	if isOLE(data) {
		legacyName, ok := embeddedLegacyFilename(name, data)
		if !ok {
			return nil, nil
		}
		return extractLegacyWithDepth(legacyName, data, depth, opts)
	}
	return nil, nil
}

func extractLegacyMarkdown(filename string, data []byte, streams []oleStream) string {
	switch strings.ToLower(filepath.Ext(filename)) {
	case ".xls":
		workbook := legacyWorkbookBytes(data, streams)
		if len(workbook) == 0 {
			return ""
		}
		return biffMarkdown(workbook)
	case ".doc":
		return legacyWordMarkdown(extractDOCLegacyText(streams))
	case ".ppt":
		if ppt, ok := findLegacyStream(streams, "PowerPoint Document"); ok {
			if md := pptLegacyMarkdown(ppt.Data); md != "" {
				return md
			}
			return legacyTextMarkdown("## Presentation", extractPPTFallbackStrings(ppt.Data))
		}
		if pp40, ok := findLegacyStream(streams, "PP40"); ok {
			return legacyTextMarkdown("## Presentation", extractPPTFallbackStrings(pp40.Data))
		}
	}
	return ""
}

func legacyTextMarkdown(heading string, parts []string) string {
	text := markdownText(joinText(uniqueStrings(parts)))
	if text == "" {
		return ""
	}
	return heading + "\n\n" + text
}

func legacyWordMarkdown(parts []string) string {
	var body, notes, comments []string
	for _, part := range uniqueStrings(parts) {
		for _, line := range strings.Split(part, "\n") {
			line = strings.TrimSpace(line)
			if line == "" {
				continue
			}
			if kind, text, ok := legacyWordStandaloneNoteLine(line); ok {
				switch kind {
				case "note":
					notes = append(notes, text)
				case "comment":
					comments = append(comments, text)
				}
				continue
			}
			body = append(body, line)
		}
	}
	var blocks []string
	if text := markdownText(joinText(uniqueStrings(body))); text != "" {
		blocks = append(blocks, "## Document\n\n"+text)
	}
	if text := markdownText(joinText(uniqueStrings(notes))); text != "" {
		blocks = append(blocks, "## Footnotes and Endnotes\n\n"+text)
	}
	if text := markdownText(joinText(uniqueStrings(comments))); text != "" {
		blocks = append(blocks, "## Comments\n\n"+text)
	}
	return strings.Join(blocks, "\n\n")
}

func legacyWordStandaloneNoteLine(line string) (string, string, bool) {
	for _, marker := range []struct {
		prefix string
		kind   string
	}{
		{prefix: "[footnote]", kind: "note"},
		{prefix: "[comment]", kind: "comment"},
	} {
		if !strings.HasPrefix(line, marker.prefix) {
			continue
		}
		text := strings.TrimSpace(strings.TrimPrefix(line, marker.prefix))
		if text == "" || strings.Contains(text, "[footnote]") || strings.Contains(text, "[comment]") {
			return "", "", false
		}
		return marker.kind, text, true
	}
	return "", "", false
}

func legacyFallbackMarkdown(filename string, parts []string) string {
	switch strings.ToLower(filepath.Ext(filename)) {
	case ".doc":
		return legacyTextMarkdown("## Document", parts)
	case ".ppt":
		return legacyTextMarkdown("## Presentation", parts)
	case ".xls":
		return legacyTextMarkdown("## Workbook", parts)
	default:
		return ""
	}
}

func legacyWorkbookBytes(data []byte, streams []oleStream) []byte {
	for _, s := range streams {
		name := strings.ToLower(s.Name)
		if name == "workbook" || name == "book" {
			return s.Data
		}
	}
	if isBIFFWorkbook(data) {
		return data
	}
	return nil
}

func extractNonOfficeOLEImages(name string, data []byte, startIndex int) []Image {
	streams, err := readOLEStreams(data)
	if err != nil || inferLegacyExt(streams) != "" {
		return nil
	}
	if len(embeddedOfficePackagePayloadsFromStreams(streams)) > 0 {
		return nil
	}
	return nonOfficeOLEImagesFromStreams(name, data, streams, startIndex)
}

func nonOfficeOLEImagesFromStreams(name string, data []byte, streams []oleStream, startIndex int) []Image {
	if inferLegacyExt(streams) != "" {
		return nil
	}
	if len(embeddedOfficePackagePayloadsFromStreams(streams)) > 0 {
		return nil
	}
	images := imagesFromOLEStreams(data, streams)
	for i := range images {
		if images[i].Name != "" {
			images[i].Name = sanitizeFilename(oleStreamBaseName(name) + "-" + images[i].Name)
		} else if images[i].Ext != "" {
			images[i].Name = fmt.Sprintf("embedded-ole-image-%03d%s", startIndex+i+1, images[i].Ext)
		}
	}
	return images
}

func imagesFromOLEStreams(data []byte, streams []oleStream) []Image {
	imageSource := data
	if len(streams) > 0 {
		var buf bytes.Buffer
		for _, s := range streams {
			buf.Write(s.Data)
		}
		imageSource = buf.Bytes()
	}
	images := carveImages(imageSource)
	if len(streams) > 0 {
		images = append(images, legacyNamedImageStreamImages(streams, len(images))...)
		images = append(images, compressedImageStreamImages(streams, len(images))...)
	}
	images = compactGeneratedLegacyImageNames(deduplicateImagesByContent(images))
	uniquifyImageNames(images)
	return images
}

func deduplicateImagesByContent(images []Image) []Image {
	if len(images) < 2 {
		return images
	}
	namedByContent := map[string]bool{}
	for _, img := range images {
		if !isGeneratedLegacyImageName(img.Name, img.Ext) {
			namedByContent[imageContentKey(img)] = true
		}
	}
	seenNamedContent := map[string]bool{}
	repeatedSmall := map[string]int{}
	for _, img := range images {
		if len(img.Data) <= maxSmallDuplicateLegacyImageBytes {
			repeatedSmall[imageContentKey(img)]++
		}
	}
	seenSmall := map[string]bool{}
	out := make([]Image, 0, len(images))
	for _, img := range images {
		key := imageContentKey(img)
		if namedByContent[key] {
			if isGeneratedLegacyImageName(img.Name, img.Ext) {
				continue
			}
			if seenNamedContent[key] {
				continue
			}
			seenNamedContent[key] = true
			out = append(out, img)
			continue
		}
		if len(img.Data) > maxSmallDuplicateLegacyImageBytes {
			out = append(out, img)
			continue
		}
		if repeatedSmall[key] < 3 {
			out = append(out, img)
			continue
		}
		if seenSmall[key] {
			continue
		}
		seenSmall[key] = true
		out = append(out, img)
	}
	return out
}

func imageContentKey(img Image) string {
	return imageContentKeyExt(img.Ext) + "\x00" + string(img.Data)
}

func imageContentKeyExt(ext string) string {
	switch strings.ToLower(ext) {
	case ".jpeg", ".jpe", ".jfif":
		return ".jpg"
	case ".jpx", ".jpf":
		return ".jp2"
	case ".j2c", ".jpc":
		return ".j2k"
	case ".hdp", ".wdp":
		return ".jxr"
	case ".tiff":
		return ".tif"
	default:
		return strings.ToLower(ext)
	}
}

func compactGeneratedLegacyImageNames(images []Image) []Image {
	next := 1
	for i := range images {
		if !isGeneratedLegacyImageName(images[i].Name, images[i].Ext) {
			continue
		}
		images[i].Name = fmt.Sprintf("legacy-image-%03d%s", next, images[i].Ext)
		next++
	}
	return images
}

func isGeneratedLegacyImageName(name, ext string) bool {
	if ext == "" || !strings.HasPrefix(name, "legacy-image-") || path.Ext(name) != ext {
		return false
	}
	base := strings.TrimSuffix(strings.TrimPrefix(name, "legacy-image-"), ext)
	if len(base) != 3 {
		return false
	}
	for _, r := range base {
		if r < '0' || r > '9' {
			return false
		}
	}
	return true
}

func extractCorruptOLEText(data []byte) []string {
	var out []string
	for _, s := range extractBinaryStrings(data) {
		s = cleanVisibleText(s)
		if looksLikePPTFallbackText(s) {
			out = append(out, s)
		}
	}
	return uniqueStrings(out)
}

func legacyXLSEncrypted(streams []oleStream) bool {
	for _, s := range streams {
		name := strings.ToLower(s.Name)
		if (name == "workbook" || name == "book") && biffHasFilePass(s.Data) {
			return true
		}
	}
	return false
}

func legacyOLEEncryptedPackage(streams []oleStream) bool {
	hasPackage := false
	hasProtection := false
	for _, s := range streams {
		name := strings.ToLower(s.Name)
		streamPath := strings.ToLower(s.Path)
		if strings.Contains(name, "encryptedpackage") || strings.Contains(streamPath, "encryptedpackage") {
			hasPackage = true
		}
		if strings.Contains(name, "encryptioninfo") || strings.Contains(streamPath, "encryptioninfo") ||
			strings.Contains(streamPath, "dataspaces") || strings.Contains(streamPath, "strongencryption") ||
			strings.Contains(streamPath, "encryptiontransform") {
			hasProtection = true
		}
	}
	return hasPackage && hasProtection
}

func biffHasFilePass(data []byte) bool {
	for off := 0; off+4 <= len(data); {
		id := binary.LittleEndian.Uint16(data[off:])
		size := int(binary.LittleEndian.Uint16(data[off+2:]))
		off += 4
		if off+size > len(data) {
			return false
		}
		if id == 0x002f {
			return true
		}
		off += size
	}
	return false
}

type oleStream struct {
	Name string
	Path string
	Data []byte
}

func readOLEStreams(data []byte) ([]oleStream, error) {
	if !isOLE(data) {
		return nil, errors.New("not an OLE compound file")
	}
	doc, err := mscfb.New(bytes.NewReader(data))
	if err != nil {
		return nil, err
	}
	var streams []oleStream
	for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
		if entry.FileInfo().IsDir() || entry.Size <= 0 {
			continue
		}
		b, err := io.ReadAll(entry)
		if err != nil {
			return nil, err
		}
		fullPath := path.Join(append(entry.Path, entry.Name)...)
		streams = append(streams, oleStream{Name: entry.Name, Path: fullPath, Data: b})
	}
	return streams, nil
}

func compressedMetafileStreamImages(streams []oleStream, startIndex int) []Image {
	return compressedImageStreamImages(streams, startIndex)
}

func compressedImageStreamImages(streams []oleStream, startIndex int) []Image {
	var images []Image
	for _, s := range streams {
		ext := strings.ToLower(path.Ext(s.Name))
		sourceName := s.Name
		if !isCompressedImageExt(ext) {
			ext = strings.ToLower(path.Ext(s.Path))
			sourceName = oleStreamBaseName(s.Path)
		}
		if !isCompressedImageExt(ext) {
			if !shouldSniffCompressedImageStream(sourceName, s.Path) {
				continue
			}
			data, imgExt, ok := normalizeSniffedCompressedImageData(s.Data)
			if !ok {
				continue
			}
			name := imageNameWithExt(oleStreamBaseName(sourceName), imgExt)
			if name == "" || name == "." {
				name = fmt.Sprintf("legacy-image-%03d%s", startIndex+len(images)+1, imgExt)
			}
			images = append(images, Image{Name: name, Ext: imgExt, Data: data})
			continue
		}
		data, imgExt, ok := normalizeCompressedImageData(ext, s.Data)
		if !ok {
			continue
		}
		name := imageNameWithExt(oleStreamBaseName(sourceName), imgExt)
		if name == "" || name == "." {
			name = fmt.Sprintf("legacy-image-%03d%s", startIndex+len(images)+1, imgExt)
		}
		images = append(images, Image{Name: name, Ext: imgExt, Data: data})
	}
	return images
}

func shouldSniffCompressedImageStream(name, streamPath string) bool {
	for _, value := range []string{name, streamPath} {
		base := strings.ToLower(path.Base(filepath.ToSlash(value)))
		switch {
		case base == "contents":
			return true
		case path.Ext(base) == ".bin":
			return true
		}
	}
	return false
}

func legacyNamedImageStreamImages(streams []oleStream, startIndex int) []Image {
	var images []Image
	for _, s := range streams {
		ext := strings.ToLower(path.Ext(s.Name))
		sourceName := s.Name
		if !isLegacyNamedImageExt(ext) {
			ext = strings.ToLower(path.Ext(s.Path))
			sourceName = oleStreamBaseName(s.Path)
		}
		if !isLegacyNamedImageExt(ext) {
			if !shouldSniffLegacyNamedImageStream(sourceName, s.Path) {
				continue
			}
			data, imageExt, ok := normalizeOOXMLImageData("", s.Data)
			if !ok {
				images = append(images, carvedLegacyNamedStreamImages(sourceName, s.Data, startIndex+len(images))...)
				continue
			}
			name := imageNameWithExt(oleStreamBaseName(sourceName), imageExt)
			if name == "" || name == "." {
				name = fmt.Sprintf("legacy-image-%03d%s", startIndex+len(images)+1, imageExt)
			}
			images = append(images, Image{Name: name, Ext: imageExt, Data: data})
			continue
		}
		data, imageExt, ok := normalizeOOXMLImageData(ext, s.Data)
		if !ok {
			images = append(images, carvedLegacyNamedStreamImages(sourceName, s.Data, startIndex+len(images))...)
			continue
		}
		name := imageNameWithExt(oleStreamBaseName(sourceName), imageExt)
		if name == "" || name == "." {
			name = fmt.Sprintf("legacy-image-%03d%s", startIndex+len(images)+1, imageExt)
		}
		images = append(images, Image{Name: name, Ext: imageExt, Data: data})
	}
	return images
}

func carvedLegacyNamedStreamImages(sourceName string, data []byte, startIndex int) []Image {
	carved := carveImages(data)
	if len(carved) == 0 {
		return nil
	}
	base := oleStreamBaseName(sourceName)
	out := make([]Image, 0, len(carved))
	for _, img := range carved {
		if img.Ext == "" || len(img.Data) == 0 {
			continue
		}
		name := imageNameWithExt(base, img.Ext)
		if len(carved) > 1 {
			stem := strings.TrimSuffix(name, path.Ext(name))
			name = fmt.Sprintf("%s-%03d%s", stem, len(out)+1, img.Ext)
		}
		if name == "" || name == "." {
			name = fmt.Sprintf("legacy-image-%03d%s", startIndex+len(out)+1, img.Ext)
		}
		out = append(out, Image{Name: name, Ext: img.Ext, Data: img.Data})
	}
	return out
}

func oleStreamBaseName(name string) string {
	return path.Base(filepath.ToSlash(name))
}

func shouldSniffLegacyNamedImageStream(name, streamPath string) bool {
	for _, value := range []string{name, streamPath} {
		base := strings.ToLower(path.Base(filepath.ToSlash(value)))
		switch {
		case base == "contents":
			return true
		case path.Ext(base) == ".bin":
			return true
		}
	}
	return false
}

func isLegacyNamedImageExt(ext string) bool {
	return isSupportedImageExt(ext) && !isCompressedImageExt(ext)
}

func isCompressedImageExt(ext string) bool {
	switch strings.ToLower(ext) {
	case ".emz", ".wmz", ".svgz":
		return true
	default:
		return false
	}
}

func isOLE(data []byte) bool {
	return len(data) >= 8 && bytes.Equal(data[:8], []byte{0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1})
}

func extractLegacyText(filename string, data []byte, streams []oleStream) []string {
	return extractLegacyTextWithMetadata(filename, data, streams, true)
}

func extractLegacyTextWithMetadata(filename string, data []byte, streams []oleStream, includeMetadata bool) []string {
	ext := strings.ToLower(filepath.Ext(filename))
	if len(streams) == 0 {
		if ext == ".xls" && isBIFFWorkbook(data) {
			if parts := biffText(data); len(parts) > 0 {
				return uniqueStrings(parts)
			}
		}
		return extractBinaryStrings(data)
	}
	var props []string
	if includeMetadata {
		props = legacyPropertySetText(streams)
	}
	switch ext {
	case ".xls":
		parts := extractXLSLegacyText(streams)
		if len(parts) > 0 {
			sheetNames := biffBoundSheetNamesFromRecords(data, 1252)
			if !includeMetadata && len(sheetNames) == 0 {
				return parts
			}
			parts = append(parts, sheetNames...)
			return uniqueCleanedStrings(append(parts, props...))
		}
		if parts := extractBinaryStrings(data); len(parts) > 0 {
			if stringSliceLooksLikePrinterSettingsDump(parts) {
				return nil
			}
			return uniqueStrings(append(parts, props...))
		}
	case ".doc":
		if _, ok := findLegacyStream(streams, "WordDocument"); ok {
			return uniqueStrings(append(extractDOCLegacyText(streams), props...))
		}
	case ".ppt":
		if ppt, ok := findLegacyStream(streams, "PowerPoint Document"); ok {
			parts := extractPPTLegacyText(streams)
			if len(parts) == 0 {
				parts = extractPPTFallbackStrings(ppt.Data)
			}
			return uniqueStrings(append(parts, props...))
		}
		if pp40, ok := findLegacyStream(streams, "PP40"); ok {
			return uniqueStrings(append(extractPPTFallbackStrings(pp40.Data), props...))
		}
	}
	return uniqueStrings(append(extractSelectedStreamStrings(streams), props...))
}

func extractNamedStreamStrings(streams []oleStream, names ...string) []string {
	wanted := map[string]bool{}
	for _, name := range names {
		wanted[strings.ToLower(name)] = true
	}
	var parts []string
	for _, s := range streams {
		if wanted[strings.ToLower(s.Name)] {
			parts = append(parts, extractBinaryStrings(s.Data)...)
		}
	}
	return uniqueStrings(parts)
}

func extractSelectedStreamStrings(streams []oleStream) []string {
	var parts []string
	for _, s := range streams {
		if shouldSkipFallbackTextStream(s) {
			continue
		}
		parts = append(parts, extractBinaryStrings(s.Data)...)
	}
	return uniqueStrings(parts)
}

func shouldSkipFallbackTextStream(s oleStream) bool {
	for _, value := range []string{s.Name, s.Path} {
		lower := strings.ToLower(filepath.ToSlash(value))
		base := path.Base(lower)
		if strings.HasPrefix(base, "\x05") || strings.Contains(base, "summaryinformation") {
			return true
		}
		cleanBase := strings.TrimLeftFunc(base, func(r rune) bool { return r < 0x20 })
		if looksLikeOLEWrapperStreamName(cleanBase) {
			return true
		}
		switch cleanBase {
		case "ole", "package", "contents":
			return true
		}
		if looksLikeLegacyInternalStoragePath(lower, cleanBase) {
			return true
		}
		if strings.Contains(lower, "/objectpool/") || strings.Contains(lower, "/ole/") || strings.Contains(lower, "/package/") {
			return true
		}
	}
	return false
}

func looksLikeLegacyInternalStoragePath(lower, cleanBase string) bool {
	pathText := "/" + strings.TrimLeft(lower, "/")
	inVBA := strings.Contains(pathText, "/vba/") || strings.Contains(pathText, "/macros/") || strings.Contains(pathText, "/_vba_project")
	inDataSpaces := strings.Contains(pathText, "/microsoft.container.dataspaces/") || strings.Contains(pathText, "/dataspaces/")
	if inVBA || inDataSpaces || strings.Contains(pathText, "/msodatastore/") {
		return true
	}
	switch cleanBase {
	case "encryptedpackage", "encryptioninfo", "strongencryptiondata",
		"_vba_project", "_vba_project_cur",
		"vba", "macros", "msodatastore":
		return true
	default:
		return false
	}
}

func legacyPropertySetText(streams []oleStream) []string {
	var parts []string
	for _, s := range streams {
		if !isPropertySetStreamName(s.Name) && !isPropertySetStreamName(s.Path) {
			continue
		}
		parts = append(parts, propertySetText(s.Data)...)
	}
	return uniqueStrings(parts)
}

func isPropertySetStreamName(name string) bool {
	lower := strings.ToLower(name)
	return strings.Contains(lower, "summaryinformation") || strings.Contains(lower, "documentsummaryinformation")
}

func extractXLSLegacyText(streams []oleStream) []string {
	var workbook []byte
	for _, s := range streams {
		name := strings.ToLower(s.Name)
		if name == "workbook" || name == "book" {
			workbook = s.Data
			break
		}
	}
	if len(workbook) == 0 {
		return nil
	}
	parts := biffText(workbook)
	if len(parts) == 0 {
		parts = extractBinaryStrings(workbook)
		if stringSliceLooksLikePrinterSettingsDump(parts) {
			return nil
		}
		return uniqueStrings(parts)
	}
	return uniqueCleanedStrings(parts)
}

func xlsLegacyTextPartsFromWorkbookData(data []byte, streams []oleStream, includeMetadata bool, workbookText []string) []string {
	parts := uniqueCleanedStrings(workbookText)
	if len(parts) == 0 {
		return nil
	}
	sheetNames := biffBoundSheetNamesFromRecords(data, 1252)
	if !includeMetadata && len(sheetNames) == 0 {
		return parts
	}
	props := []string(nil)
	if includeMetadata {
		props = legacyPropertySetText(streams)
	}
	return uniqueCleanedStrings(append(append(parts, sheetNames...), props...))
}

func extractDOCLegacyText(streams []oleStream) []string {
	word, ok := findLegacyStream(streams, "WordDocument")
	if !ok {
		return nil
	}
	table, _ := findLegacyStream(streams, wordTableStreamName(word.Data))
	return uniqueStrings(wordDocumentText(word.Data, table.Data))
}

func findLegacyStream(streams []oleStream, name string) (oleStream, bool) {
	for _, s := range streams {
		if strings.EqualFold(s.Name, name) {
			return s, true
		}
	}
	return oleStream{}, false
}

func wordTableStreamName(word []byte) string {
	if len(word) >= 0x0c && binary.LittleEndian.Uint16(word[0x0a:])&0x0200 != 0 {
		return "1Table"
	}
	return "0Table"
}

func wordDocumentText(data, table []byte) []string {
	if len(data) < 32 {
		return nil
	}
	switch binary.LittleEndian.Uint16(data) {
	case 0xa5ec:
		if parts := wordPieceTableText(data, table); len(parts) > 0 {
			return parts
		}
		return extractWordTextRange(data)
	case 0xa5dc:
		return extractWordTextRange(data)
	default:
		return nil
	}
}

func extractWordTextRange(data []byte) []string {
	fcMin := int(binary.LittleEndian.Uint32(data[0x18:]))
	fcMac := int(binary.LittleEndian.Uint32(data[0x1c:]))
	if fcMin < 0 || fcMac <= fcMin || fcMac > len(data) {
		return nil
	}
	return extractWordFallbackStrings(data[fcMin:fcMac])
}

func wordPieceTableText(word, table []byte) []string {
	if len(word) < 0x1aa || len(table) == 0 {
		return nil
	}
	fcClx := int(binary.LittleEndian.Uint32(word[0x01a2:]))
	lcbClx := int(binary.LittleEndian.Uint32(word[0x01a6:]))
	if fcClx < 0 || lcbClx <= 0 || fcClx+lcbClx > len(table) {
		return nil
	}
	return parseWordCLXText(word, table[fcClx:fcClx+lcbClx])
}

func parseWordCLXText(word, clx []byte) []string {
	var out []string
	for off := 0; off < len(clx); {
		switch clx[off] {
		case 0x01:
			if off+3 > len(clx) {
				return out
			}
			size := int(binary.LittleEndian.Uint16(clx[off+1:]))
			off += 3 + size
		case 0x02:
			if off+5 > len(clx) {
				return out
			}
			size := int(binary.LittleEndian.Uint32(clx[off+1:]))
			off += 5
			if size < 4 || off+size > len(clx) {
				return out
			}
			out = append(out, parseWordPieceTableText(word, clx[off:off+size])...)
			off += size
		default:
			off++
		}
	}
	return uniqueStrings(out)
}

func parseWordPieceTableText(word, plc []byte) []string {
	if len(plc) < 16 || (len(plc)-4)%12 != 0 {
		return nil
	}
	pieces := (len(plc) - 4) / 12
	cpOff := 0
	pcdOff := (pieces + 1) * 4
	if pcdOff+pieces*8 > len(plc) {
		return nil
	}
	legacyCodePage := wordPieceLegacyCodePage(word, plc, pieces, pcdOff)
	var out []string
	for i := 0; i < pieces; i++ {
		cpStart := binary.LittleEndian.Uint32(plc[cpOff+i*4:])
		cpEnd := binary.LittleEndian.Uint32(plc[cpOff+(i+1)*4:])
		if cpEnd <= cpStart {
			continue
		}
		charCount := int(cpEnd - cpStart)
		pcd := plc[pcdOff+i*8:]
		fcRaw := binary.LittleEndian.Uint32(pcd[2:])
		compressed := fcRaw&0x40000000 != 0
		fc := int(fcRaw & 0x3fffffff)
		var raw []byte
		if compressed {
			fc /= 2
			if fc < 0 || charCount > len(word)-fc {
				continue
			}
			raw = word[fc : fc+charCount]
			addWordText(&out, decodeWordSingleByteText(raw, legacyCodePage))
			continue
		}
		byteCount := charCount * 2
		if fc < 0 || byteCount > len(word)-fc {
			continue
		}
		raw = word[fc : fc+byteCount]
		addWordText(&out, wordPieceUTF16BytesToString(raw, legacyCodePage))
	}
	return uniqueStrings(out)
}

func wordPieceLegacyCodePage(word, plc []byte, pieces, pcdOff int) uint16 {
	var raw []byte
	for i := 0; i < pieces; i++ {
		cpStart := binary.LittleEndian.Uint32(plc[i*4:])
		cpEnd := binary.LittleEndian.Uint32(plc[(i+1)*4:])
		if cpEnd <= cpStart {
			continue
		}
		charCount := int(cpEnd - cpStart)
		pcd := plc[pcdOff+i*8:]
		fcRaw := binary.LittleEndian.Uint32(pcd[2:])
		compressed := fcRaw&0x40000000 != 0
		fc := int(fcRaw & 0x3fffffff)
		if compressed {
			fc /= 2
			if fc >= 0 && charCount <= len(word)-fc {
				raw = append(raw, word[fc:fc+charCount]...)
			}
			continue
		}
		byteCount := charCount * 2
		if fc < 0 || byteCount > len(word)-fc {
			continue
		}
		if low, ok := zeroHighByteTextBytes(word[fc : fc+byteCount]); ok {
			raw = append(raw, low...)
		}
	}
	return legacySingleByteCodePage(raw)
}

func legacySingleByteCodePage(raw []byte) uint16 {
	if len(raw) == 0 {
		return 0
	}
	if utf8.Valid(raw) && hasUTF8Multibyte(raw) {
		return 65001
	}
	if codePage := bestLegacySingleByteCodePage(raw); codePage != 0 {
		return codePage
	}
	return 0
}

func decodeWordSingleByteText(raw []byte, codePage uint16) string {
	if bytes.IndexAny(raw, "\x02\x05\x07\x0b\x0c") >= 0 {
		var out strings.Builder
		start := 0
		flush := func(end int) {
			if end > start {
				out.WriteString(decodeWordSingleByteTextRun(raw[start:end], codePage))
			}
		}
		for i, b := range raw {
			switch b {
			case 0x02:
				flush(i)
				out.WriteString("[footnote] ")
				start = i + 1
			case 0x05:
				flush(i)
				out.WriteString("[comment] ")
				start = i + 1
			case 0x07:
				flush(i)
				out.WriteByte('\n')
				start = i + 1
			case 0x0b, 0x0c:
				flush(i)
				out.WriteByte('\n')
				start = i + 1
			}
		}
		flush(len(raw))
		return out.String()
	}
	return decodeWordSingleByteTextRun(raw, codePage)
}

func decodeWordSingleByteTextRun(raw []byte, codePage uint16) string {
	if codePage != 0 {
		return decodeCodePageBytes(raw, codePage)
	}
	return compressedUnicodeBytesToString(raw)
}

func addWordText(out *[]string, s string) {
	for _, part := range cleanWordVisibleTextParts(s) {
		*out = append(*out, part)
	}
}

func cleanWordVisibleTextParts(s string) []string {
	s = normalizeWordSpecialTextChars(s)
	s = cleanVisibleText(s)
	if s == "" || !looksLikeTextFragment(s) || !looksLikeWordFallbackText(s) {
		return nil
	}
	if !looksLikeBinaryControlFragment(s) {
		return []string{s}
	}
	if !strings.Contains(s, "\n") {
		return nil
	}
	var out []string
	for _, line := range strings.Split(s, "\n") {
		line = strings.TrimSpace(cleanVisibleText(line))
		if line == "" || !looksLikeTextFragment(line) || looksLikeBinaryControlFragment(line) || !looksLikeWordFallbackText(line) {
			continue
		}
		out = append(out, line)
	}
	return out
}

func extractWordFallbackStrings(data []byte) []string {
	data = maskEmbeddedImagesForText(data)
	legacyCodePage := legacySingleByteCodePage(data)
	seen := map[string]bool{}
	var out []string
	add := func(s string) {
		for _, part := range cleanWordVisibleTextParts(s) {
			if len([]rune(part)) < 3 || seen[part] {
				continue
			}
			seen[part] = true
			out = append(out, part)
		}
	}
	for _, s := range wordSingleByteStrings(data, 4, legacyCodePage) {
		add(s)
	}
	if looksLikeDenseSingleByteText(data) {
		return out
	}
	for _, s := range utf16Strings(data, 3) {
		add(s)
	}
	return out
}

func wordSingleByteStrings(data []byte, min int, codePage uint16) []string {
	var out []string
	var cur []byte
	flush := func() {
		if len(cur) >= min {
			out = append(out, decodeWordSingleByteText(cur, codePage))
		}
		cur = cur[:0]
	}
	for _, b := range data {
		if isLegacySingleByteTextByte(b) || isWordSpecialTextByte(b) {
			cur = append(cur, b)
		} else {
			flush()
		}
	}
	flush()
	return out
}

func isWordSpecialTextByte(b byte) bool {
	return b == 0x02 || b == 0x05 || b == 0x07 || b == 0x0b || b == 0x0c
}

func normalizeWordSpecialTextChars(s string) string {
	if !strings.ContainsAny(s, "\x02\x05\x07\x0b\x0c") {
		return s
	}
	var out strings.Builder
	out.Grow(len(s))
	for _, r := range s {
		switch r {
		case '\x02':
			out.WriteString("[footnote] ")
		case '\x05':
			out.WriteString("[comment] ")
		case '\x07':
			out.WriteByte('\n')
		case '\x0b', '\x0c':
			out.WriteByte('\n')
		default:
			out.WriteRune(r)
		}
	}
	return out.String()
}

func looksLikeDenseSingleByteText(data []byte) bool {
	var printable, high, zeros int
	for _, b := range data {
		switch {
		case b == 0:
			zeros++
		case b == '\t' || b == '\n' || b == '\r' || b >= 0x20:
			printable++
			if b >= 0x80 {
				high++
			}
		}
	}
	return printable >= 8 && high >= 4 && high*3 >= printable && zeros*10 < len(data)
}

func looksLikeWordFallbackText(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	lower := strings.ToLower(s)
	if strings.Contains(lower, "http://") || strings.Contains(lower, "https://") || strings.Contains(s, "@") {
		return true
	}
	if strings.ContainsAny(s, " \t\n") {
		return true
	}
	letters, asciiLetters, vowels, digits, marks, nonASCII := 0, 0, 0, 0, 0, 0
	for _, r := range s {
		switch {
		case unicode.IsLetter(r):
			letters++
			if r <= unicode.MaxASCII {
				asciiLetters++
				if strings.ContainsRune("aeiouAEIOU", r) {
					vowels++
				}
			} else {
				nonASCII++
			}
		case unicode.IsDigit(r):
			digits++
		case unicode.IsPunct(r) || unicode.IsSymbol(r):
			marks++
		}
	}
	if nonASCII > 0 && letters > marks+digits {
		return true
	}
	if asciiLetters >= 5 && digits == 0 && marks == 0 && vowels*4 >= asciiLetters {
		return true
	}
	if asciiLetters >= 3 && digits == 0 && marks == 1 {
		last := rune(s[len(s)-1])
		if strings.ContainsRune(".!?:;,", last) && vowels*4 >= asciiLetters {
			return true
		}
	}
	if asciiLetters >= 2 && digits == 0 && marks == 0 && vowels > 0 && len([]rune(s)) <= 4 {
		return true
	}
	return false
}

func extractPPTLegacyText(streams []oleStream) []string {
	s, ok := findLegacyStream(streams, "PowerPoint Document")
	if !ok {
		return nil
	}
	return uniqueStrings(pptRecordText(s.Data))
}

func extractPPTFallbackStrings(data []byte) []string {
	var out []string
	for _, s := range extractBinaryStrings(data) {
		s = strings.TrimSpace(cleanVisibleText(s))
		if looksLikePPTFallbackText(s) {
			out = append(out, s)
		}
	}
	return uniqueStrings(out)
}

func looksLikePPTFallbackText(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" || looksLikeBinaryControlFragment(s) {
		return false
	}
	lower := strings.ToLower(s)
	if strings.Contains(lower, "http://") || strings.Contains(lower, "https://") || strings.Contains(s, "@") {
		return true
	}
	letters, asciiLetters, vowels, digits, marks, spaces, nonASCII := 0, 0, 0, 0, 0, 0, 0
	for _, r := range s {
		switch {
		case unicode.IsLetter(r):
			letters++
			if r <= unicode.MaxASCII {
				asciiLetters++
				if strings.ContainsRune("aeiouAEIOU", r) {
					vowels++
				}
			} else {
				nonASCII++
			}
		case unicode.IsDigit(r):
			digits++
		case unicode.IsPunct(r) || unicode.IsSymbol(r):
			marks++
		case unicode.IsMark(r):
			marks++
		case unicode.IsSpace(r):
			spaces++
		}
	}
	if nonASCII > 0 && letters > marks+digits {
		return looksLikePPTFallbackNonASCIIText(s)
	}
	if spaces > 0 && asciiLetters >= 3 && vowels > 0 && marks <= letters+spaces {
		return vowels*6 >= asciiLetters
	}
	return false
}

func looksLikePPTFallbackNonASCIIText(s string) bool {
	runes := []rune(strings.TrimSpace(s))
	if len(runes) == 0 || looksLikeUnicodeBinaryNoise(s) {
		return false
	}
	var asciiLetters, digits, spaces, marks, cjk, hangul, kana, latin, cyrillic, otherLetters, lowByteControls int
	for _, r := range runes {
		switch {
		case r <= unicode.MaxASCII && unicode.IsLetter(r):
			asciiLetters++
		case unicode.IsDigit(r):
			digits++
		case unicode.IsSpace(r):
			spaces++
		case unicode.IsPunct(r) || unicode.IsSymbol(r) || unicode.IsMark(r):
			marks++
		case unicode.Is(unicode.Han, r):
			cjk++
		case unicode.Is(unicode.Hangul, r):
			hangul++
		case unicode.In(r, unicode.Hiragana, unicode.Katakana):
			kana++
		case unicode.Is(unicode.Latin, r):
			latin++
		case unicode.Is(unicode.Cyrillic, r):
			cyrillic++
		case unicode.IsLetter(r):
			otherLetters++
		}
		if r >= 0x0100 {
			low := r & 0x00ff
			if low <= 0x1f || low == 0x80 || low == 0x81 {
				lowByteControls++
			}
		}
	}
	scriptGroups := 0
	for _, n := range []int{cjk, hangul, kana, latin, cyrillic, otherLetters} {
		if n > 0 {
			scriptGroups++
		}
	}
	if spaces > 0 {
		if lowByteControls > 0 && lowByteControls*3 >= len(runes) {
			return false
		}
		if asciiLetters == 0 && scriptGroups > 1 && kana == 0 {
			return false
		}
		if asciiLetters == 0 && marks+digits > 0 && cjk+hangul+kana == 0 {
			return false
		}
		return cjk+hangul+kana+latin+cyrillic+otherLetters+asciiLetters > 0 && marks <= len(runes)/2
	}
	if asciiLetters > 0 || digits > 0 || marks > 0 {
		return false
	}
	if scriptGroups == 0 {
		return false
	}
	if scriptGroups > 1 && kana == 0 {
		return false
	}
	return cjk+hangul+kana >= len(runes)*2/3
}

func pptRecordText(data []byte) []string {
	var out []string
	pptRecordTextInto(data, 0, &out)
	return uniqueStrings(cleanPPTRecordTextParts(out))
}

type pptMarkdownSection struct {
	heading string
	parts   []string
}

func pptLegacyMarkdown(data []byte) string {
	sections := pptRecordMarkdownSections(data)
	if len(sections) == 0 {
		return legacyTextMarkdown("## Presentation", pptRecordText(data))
	}
	var blocks []string
	for _, section := range sections {
		text := markdownText(joinText(uniqueStrings(cleanPPTRecordTextParts(section.parts))))
		if text == "" {
			continue
		}
		blocks = append(blocks, section.heading+"\n\n"+text)
	}
	if len(blocks) == 0 {
		return legacyTextMarkdown("## Presentation", pptRecordText(data))
	}
	return strings.TrimSpace(strings.Join(blocks, "\n\n"))
}

func pptRecordMarkdownSections(data []byte) []pptMarkdownSection {
	var sections []pptMarkdownSection
	var loose []string
	var notes []string
	pptRecordMarkdownSectionsInto(data, 0, &sections, &loose, &notes)
	if len(loose) > 0 {
		sections = append([]pptMarkdownSection{{heading: "## Presentation", parts: loose}}, sections...)
	}
	if len(notes) > 0 {
		appendPPTMarkdownSection(&sections, pptMarkdownSection{heading: "## Notes and Comments", parts: notes})
	}
	return sections
}

func appendPPTMarkdownSection(sections *[]pptMarkdownSection, section pptMarkdownSection) {
	if len(section.parts) == 0 {
		return
	}
	for i := range *sections {
		if (*sections)[i].heading == section.heading {
			(*sections)[i].parts = append((*sections)[i].parts, section.parts...)
			return
		}
	}
	*sections = append(*sections, section)
}

func pptRecordMarkdownSectionsInto(data []byte, depth int, sections *[]pptMarkdownSection, loose *[]string, notes *[]string) {
	if depth > 32 {
		return
	}
	for off := 0; off+8 <= len(data); {
		options := binary.LittleEndian.Uint16(data[off:])
		recType := binary.LittleEndian.Uint16(data[off+2:])
		size := int(binary.LittleEndian.Uint32(data[off+4:]))
		off += 8
		if size < 0 || size > len(data)-off {
			return
		}
		payload := data[off : off+size]
		if recType == 0x03ee && options&0x000f == 0x000f { // Slide container
			parts, slideNotes := pptRecordMarkdownSlideParts(payload, depth+1)
			if len(parts) > 0 {
				appendPPTMarkdownSection(sections, pptMarkdownSection{
					heading: fmt.Sprintf("## Slide %d", len(*sections)+1),
					parts:   parts,
				})
			}
			if len(slideNotes) > 0 {
				*notes = append(*notes, slideNotes...)
			}
			off += size
			continue
		}
		switch recType {
		case 0x0fa0: // TextCharsAtom
			text := utf16BytesToStringAll(payload)
			if !looksLikeASCIIBytesMisreadAsUTF16(payload, text) {
				addStructuredText(loose, text)
			}
		case 0x0fa8: // TextBytesAtom, compressed Unicode
			addStructuredText(loose, compressedUnicodeBytesToString(payload))
		case 0x0fba: // CString, used by comments and other textual records
			if s, ok := decodePPTCString(payload); ok {
				addStructuredTextIfNotControl(notes, s)
			}
		}
		if options&0x000f == 0x000f {
			pptRecordMarkdownSectionsInto(payload, depth+1, sections, loose, notes)
		}
		off += size
	}
}

func pptRecordMarkdownSlideParts(data []byte, depth int) ([]string, []string) {
	var body, notes []string
	pptRecordMarkdownSlidePartsInto(data, depth, &body, &notes)
	return uniqueStrings(cleanPPTRecordTextParts(body)), uniqueStrings(cleanPPTRecordTextParts(notes))
}

func pptRecordMarkdownSlidePartsInto(data []byte, depth int, body *[]string, notes *[]string) {
	if depth > 32 {
		return
	}
	for off := 0; off+8 <= len(data); {
		options := binary.LittleEndian.Uint16(data[off:])
		recType := binary.LittleEndian.Uint16(data[off+2:])
		size := int(binary.LittleEndian.Uint32(data[off+4:]))
		off += 8
		if size < 0 || size > len(data)-off {
			return
		}
		payload := data[off : off+size]
		switch recType {
		case 0x0fa0: // TextCharsAtom
			text := utf16BytesToStringAll(payload)
			if !looksLikeASCIIBytesMisreadAsUTF16(payload, text) {
				addStructuredText(body, text)
			}
		case 0x0fa8: // TextBytesAtom, compressed Unicode
			addStructuredText(body, compressedUnicodeBytesToString(payload))
		case 0x0fba: // CString, used by comments and other textual records
			if s, ok := decodePPTCString(payload); ok {
				addStructuredTextIfNotControl(notes, s)
			}
		}
		if options&0x000f == 0x000f {
			pptRecordMarkdownSlidePartsInto(payload, depth+1, body, notes)
		}
		off += size
	}
}

func pptRecordTextInto(data []byte, depth int, out *[]string) {
	if depth > 32 {
		return
	}
	for off := 0; off+8 <= len(data); {
		options := binary.LittleEndian.Uint16(data[off:])
		recType := binary.LittleEndian.Uint16(data[off+2:])
		size := int(binary.LittleEndian.Uint32(data[off+4:]))
		off += 8
		if size < 0 || size > len(data)-off {
			return
		}
		payload := data[off : off+size]
		switch recType {
		case 0x0fa0: // TextCharsAtom
			text := utf16BytesToStringAll(payload)
			if !looksLikeASCIIBytesMisreadAsUTF16(payload, text) {
				addStructuredText(out, text)
			}
		case 0x0fa8: // TextBytesAtom, compressed Unicode
			addStructuredText(out, compressedUnicodeBytesToString(payload))
		case 0x0fba: // CString, used by comments and other textual records
			if s, ok := decodePPTCString(payload); ok {
				addStructuredTextIfNotControl(out, s)
			}
		}
		if options&0x000f == 0x000f {
			pptRecordTextInto(payload, depth+1, out)
		}
		off += size
	}
}

func looksLikeASCIIBytesMisreadAsUTF16(raw []byte, decoded string) bool {
	if len(raw) < 4 {
		return false
	}
	var ascii, zeros, high int
	for _, b := range raw {
		if b == 0 {
			zeros++
		}
		if b >= 0x80 {
			high++
		}
		if b == '\t' || b == '\n' || b == '\r' || (b >= 0x20 && b <= 0x7e) {
			ascii++
		}
	}
	if high*8 >= len(raw) {
		return false
	}
	if ascii*4 < len(raw)*3 || zeros*8 >= len(raw) {
		return false
	}
	var cjk, letters int
	for _, r := range decoded {
		if unicode.IsLetter(r) {
			letters++
			if isCJKOrHangul(r) {
				cjk++
			}
		}
	}
	return letters >= 2 && cjk*2 >= letters
}

func isCJKOrHangul(r rune) bool {
	return (r >= 0x3400 && r <= 0x9fff) || (r >= 0xac00 && r <= 0xd7af)
}

func addStructuredText(out *[]string, s string) {
	s = cleanVisibleText(s)
	if strings.Contains(s, "\n") {
		addStructuredTextLines(out, s)
		return
	}
	addStructuredTextFragment(out, s)
}

func addStructuredTextLines(out *[]string, s string) {
	for i, line := range strings.Split(s, "\n") {
		line = strings.TrimSpace(line)
		if i == 0 && looksLikePPTLeadingRecordControlLine(line) {
			continue
		}
		addStructuredTextFragment(out, line)
	}
}

func addStructuredTextFragment(out *[]string, s string) {
	s = strings.TrimSpace(cleanVisibleText(s))
	if s == "" || !looksLikeTextFragment(s) || looksLikeBinaryControlFragment(s) || looksLikePPTEmbeddedObjectLabel(s) || looksLikePPTDesignThemeLabel(s) || looksLikePowerPointPlaceholderPrompt(s) {
		return
	}
	*out = append(*out, s)
}

func addStructuredTextIfNotControl(out *[]string, s string) {
	addStructuredText(out, s)
}

func looksLikePPTLeadingRecordControlLine(s string) bool {
	return strings.TrimSpace(s) == "0"
}

func looksLikePowerPointPlaceholderPrompt(s string) bool {
	normalized := strings.ToLower(spaceRE.ReplaceAllString(strings.TrimSpace(s), " "))
	switch normalized {
	case "click to add title",
		"click to add subtitle",
		"click to add text",
		"click to add body",
		"click to add content",
		"click to add vertical title",
		"click to add vertical text",
		"click to add vertical body",
		"click to add chart",
		"click to add table",
		"click to add diagram",
		"click to add organization chart",
		"click to add media clip",
		"click to add clip art",
		"click to add picture",
		"click to add object",
		"click to add notes",
		"click to add date",
		"click to add footer",
		"click to add slide number",
		"click icon to add chart",
		"click icon to add table",
		"click icon to add smartart graphic",
		"click icon to add picture",
		"click icon to add media clip",
		"double click to add chart",
		"double click to add table",
		"double click to add diagram",
		"double click to add organization chart",
		"double click to add clip art",
		"double click to add picture",
		"double click to add media clip",
		"double click to add object",
		"click to edit master title style",
		"click to edit master subtitle style",
		"click to edit master text styles",
		"click to edit master text style",
		"click to edit master body style",
		"click to edit master body text style",
		"click to edit master body text styles",
		"click to edit master notes style",
		"click to edit master footer style",
		"click to edit master date style",
		"click to edit master slide number style",
		"マスタ タイトルの書式設定",
		"マスタ テキストの書式設定":
		return true
	default:
		return false
	}
}

func cleanPPTRecordTextParts(parts []string) []string {
	hasSubstantial := false
	for _, part := range parts {
		part = strings.TrimSpace(cleanVisibleText(part))
		if part != "" && part != "0" && len([]rune(part)) >= 3 {
			hasSubstantial = true
			break
		}
	}
	out := make([]string, 0, len(parts))
	for _, part := range parts {
		part = strings.TrimSpace(cleanVisibleText(part))
		if part == "" {
			continue
		}
		if hasSubstantial && part == "0" {
			continue
		}
		if hasSubstantial {
			part = stripPPTLeadingRecordControlPrefix(part)
		}
		if part != "" {
			out = append(out, part)
		}
	}
	return out
}

func stripPPTLeadingRecordControlPrefix(s string) string {
	if !strings.HasPrefix(s, "0") {
		return s
	}
	rest := strings.TrimSpace(s[1:])
	if len([]rune(rest)) < 3 || !strings.Contains(rest, " ") {
		return s
	}
	r, _ := utf8.DecodeRuneInString(rest)
	if unicode.IsLetter(r) {
		return rest
	}
	return s
}

func looksLikePPTEmbeddedObjectLabel(s string) bool {
	lower := strings.ToLower(strings.TrimSpace(s))
	switch lower {
	case "equation", "chart", "microsoft graph chart", "microsoft graph 2000 chart",
		"microsoft graph 97 chart", "ms graph chart", "msgraph.chart", "msgraph.chart.8",
		"ms org chart", "ms organization chart 2.0", "organization chart",
		"microsoft equation", "microsoft equation 3.0", "equation.3",
		"package", "package object", "packager shell object",
		"microsoft word document", "microsoft office word document",
		"microsoft excel worksheet", "microsoft office excel worksheet",
		"microsoft powerpoint presentation", "microsoft office powerpoint presentation",
		"adobe acrobat document", "acrobat document", "pdf document",
		"microsoft visio drawing",
		"document", "worksheet", "slide", "presentation":
		return true
	default:
		return looksLikePPTEmbeddedObjectProgID(lower)
	}
}

func looksLikePPTEmbeddedObjectProgID(s string) bool {
	s = strings.Trim(s, " \t\r\n\"'()[]{}<>")
	if s == "" || len(s) > 80 || strings.ContainsAny(s, "\n\r\t") {
		return false
	}
	if strings.Count(s, " ") > 3 {
		return false
	}
	for _, suffix := range []string{".1", ".2", ".3", ".4", ".5", ".6", ".7", ".8", ".9", ".10", ".11", ".12", ".14", ".15", ".16"} {
		s = strings.TrimSuffix(s, suffix)
	}
	switch s {
	case "msgraph.chart", "word.document", "excel.sheet", "excel.chart",
		"powerpoint.slide", "powerpoint.show", "powerpoint.presentation",
		"acroexch.document", "visio.drawing", "equation",
		"forms.textbox", "forms.checkbox", "forms.combobox", "forms.listbox",
		"htmlfile", "package", "packager.shell.object":
		return true
	default:
		return false
	}
}

func decodePPTCString(raw []byte) (string, bool) {
	if len(raw) >= 2 && len(raw)%2 == 0 {
		units := make([]uint16, 0, len(raw)/2)
		for i := 0; i+1 < len(raw); i += 2 {
			units = append(units, binary.LittleEndian.Uint16(raw[i:]))
		}
		text := cleanText(string(utf16.Decode(units)))
		if text != "" && printableRatio(text) > 0.75 && hasUTF16Evidence(raw) && !looksLikeMisalignedUTF16(units) && !looksLikeASCIIBytesMisreadAsUTF16(raw, text) && looksLikeTextFragment(text) {
			return text, true
		}
	}
	text := cleanText(compressedUnicodeBytesToString(raw))
	if text == "" || !looksLikeTextFragment(text) {
		return "", false
	}
	return text, true
}

func propertySetText(data []byte) []string {
	if len(data) < 48 || binary.LittleEndian.Uint16(data) != 0xfffe {
		return nil
	}
	numSets := int(binary.LittleEndian.Uint32(data[24:]))
	if numSets <= 0 || numSets > 8 || len(data) < 28+numSets*20 {
		return nil
	}
	var out []string
	for i := 0; i < numSets; i++ {
		setOffset := int(binary.LittleEndian.Uint32(data[28+i*20+16:]))
		out = append(out, propertySetSectionText(data, setOffset)...)
	}
	return uniqueStrings(out)
}

func propertySetSectionText(data []byte, setOffset int) []string {
	if setOffset < 0 || setOffset+8 > len(data) {
		return nil
	}
	size := int(binary.LittleEndian.Uint32(data[setOffset:]))
	if size < 8 || setOffset+size > len(data) {
		return nil
	}
	section := data[setOffset : setOffset+size]
	count := int(binary.LittleEndian.Uint32(section[4:]))
	if count <= 0 || count > 4096 || 8+count*8 > len(section) {
		return nil
	}
	codePage := uint16(1252)
	codePageExplicit := false
	props := make([]struct {
		id  uint32
		off int
	}, 0, count)
	for i := 0; i < count; i++ {
		entry := section[8+i*8:]
		id := binary.LittleEndian.Uint32(entry)
		off := int(binary.LittleEndian.Uint32(entry[4:]))
		if off < 0 || off+4 > len(section) {
			continue
		}
		props = append(props, struct {
			id  uint32
			off int
		}{id: id, off: off})
		if id == 1 {
			if cp, ok := propertySetCodePage(section[off:]); ok {
				codePage = cp
				codePageExplicit = true
			}
		}
	}
	var out []string
	for _, prop := range props {
		if prop.id == 0 || prop.id == 1 {
			continue
		}
		if value, ok := propertySetStringValue(section[prop.off:], codePage, codePageExplicit); ok {
			value = cleanVisibleText(value)
			if value != "" {
				out = append(out, value)
			}
		}
	}
	return uniqueStrings(out)
}

func propertySetCodePage(data []byte) (uint16, bool) {
	if len(data) < 8 {
		return 0, false
	}
	typ := binary.LittleEndian.Uint16(data)
	switch typ {
	case 0x0002:
		return uint16(binary.LittleEndian.Uint16(data[4:])), true
	case 0x0003:
		return uint16(binary.LittleEndian.Uint32(data[4:])), true
	default:
		return 0, false
	}
}

func propertySetStringValue(data []byte, codePage uint16, codePageExplicit bool) (string, bool) {
	if len(data) < 8 {
		return "", false
	}
	typ := binary.LittleEndian.Uint16(data)
	switch typ {
	case 0x001e:
		return propertySetLPSTRWithCodePage(data[4:], codePage, codePageExplicit)
	case 0x001f:
		return propertySetLPWSTR(data[4:])
	case 0x0008:
		return propertySetBSTR(data[4:])
	default:
		return "", false
	}
}

func propertySetLPSTR(data []byte, codePage uint16) (string, bool) {
	return propertySetLPSTRWithCodePage(data, codePage, true)
}

func propertySetLPSTRWithCodePage(data []byte, codePage uint16, codePageExplicit bool) (string, bool) {
	if len(data) < 4 {
		return "", false
	}
	size := int(binary.LittleEndian.Uint32(data))
	if size <= 0 || size > len(data)-4 {
		return "", false
	}
	raw := data[4 : 4+size]
	raw = bytes.TrimRight(raw, "\x00")
	var text string
	if codePage == 1200 {
		text = utf16BytesToString(raw)
	} else if codePage == 1252 && shouldDecodeWindows1251(raw) {
		text = decodeCodePageBytes(raw, 1251)
	} else if !codePageExplicit && (codePage == 0 || codePage == 1252) {
		text = decodeBestLegacySingleByte(raw)
	} else {
		text = decodeCodePageBytes(raw, codePage)
	}
	text = cleanText(text)
	return text, text != ""
}

func propertySetLPWSTR(data []byte) (string, bool) {
	if len(data) < 4 {
		return "", false
	}
	chars := int(binary.LittleEndian.Uint32(data))
	if chars <= 0 || chars > (len(data)-4)/2 {
		return "", false
	}
	raw := data[4 : 4+chars*2]
	text := cleanText(utf16BytesToString(raw))
	return text, text != ""
}

func propertySetBSTR(data []byte) (string, bool) {
	if len(data) < 4 {
		return "", false
	}
	size := int(binary.LittleEndian.Uint32(data))
	if size <= 0 || size > len(data)-4 {
		return "", false
	}
	raw := data[4 : 4+size]
	text := cleanText(utf16BytesToString(raw))
	return text, text != ""
}

func utf16BytesToString(raw []byte) string {
	if len(raw)%2 == 1 {
		raw = raw[:len(raw)-1]
	}
	u := make([]uint16, 0, len(raw)/2)
	for i := 0; i+1 < len(raw); i += 2 {
		v := binary.LittleEndian.Uint16(raw[i:])
		if v == 0 {
			break
		}
		u = append(u, v)
	}
	return string(utf16.Decode(u))
}

func utf16BytesToStringAll(raw []byte) string {
	if len(raw)%2 == 1 {
		raw = raw[:len(raw)-1]
	}
	u := make([]uint16, 0, len(raw)/2)
	for i := 0; i+1 < len(raw); i += 2 {
		u = append(u, binary.LittleEndian.Uint16(raw[i:]))
	}
	return string(utf16.Decode(u))
}

func wordPieceUTF16BytesToString(raw []byte, legacyCodePage uint16) string {
	text := utf16BytesToStringAll(raw)
	if legacy, ok := decodeZeroHighByteLegacyText(raw, legacyCodePage); ok {
		return legacy
	}
	return text
}

func decodeZeroHighByteLegacyText(raw []byte, legacyCodePage uint16) (string, bool) {
	low, ok := zeroHighByteTextBytes(raw)
	if !ok {
		return "", false
	}
	if legacyCodePage != 0 {
		text := decodeCodePageBytes(low, legacyCodePage)
		return text, text != "" && looksLikeTextFragment(text)
	}
	if !shouldDecodeWindows1251(low) && !(utf8.Valid(low) && hasUTF8Multibyte(low)) {
		return "", false
	}
	text := decodeBestLegacySingleByte(low)
	if text == "" || !looksLikeTextFragment(text) {
		return "", false
	}
	return text, true
}

func zeroHighByteTextBytes(raw []byte) ([]byte, bool) {
	if len(raw) < 2 || len(raw)%2 == 1 {
		return nil, false
	}
	low := make([]byte, 0, len(raw)/2)
	zeroHigh, nonZeroHigh, highLow := 0, 0, 0
	for i := 0; i+1 < len(raw); i += 2 {
		lo, hi := raw[i], raw[i+1]
		low = append(low, lo)
		if hi == 0 {
			zeroHigh++
		} else {
			nonZeroHigh++
		}
		if lo >= 0x80 {
			highLow++
		}
	}
	if zeroHigh == 0 || nonZeroHigh*20 > zeroHigh || highLow == 0 {
		return nil, false
	}
	return low, true
}

func compressedUnicodeBytesToString(raw []byte) string {
	return decodeBestLegacySingleByte(raw)
}

func biffMarkdown(data []byte) string {
	return biffMarkdownWithText(data, nil)
}

func biffMarkdownWithText(data []byte, textParts []string) string {
	tables := biffWorksheetMarkdownTables(data)
	if len(tables) == 0 {
		return ""
	}
	var out []string
	for _, table := range tables {
		rows := compactMarkdownTableRows(table.rows)
		var blocks []string
		if md := markdownTablePrepared(rows); md != "" {
			blocks = append(blocks, md)
		}
		if len(table.headerFooter) > 0 {
			blocks = append(blocks, "### Headers and Footers\n\n"+markdownText(strings.Join(table.headerFooter, "\n")))
		}
		if len(table.comments) > 0 {
			var commentBlocks []string
			for _, comment := range table.comments {
				text := markdownText(comment.text)
				if text == "" {
					continue
				}
				if comment.ref != "" {
					text = "#### " + escapeMarkdownHeading(comment.ref) + "\n\n" + text
				}
				commentBlocks = append(commentBlocks, text)
			}
			if len(commentBlocks) > 0 {
				blocks = append(blocks, "### Comments\n\n"+strings.Join(commentBlocks, "\n\n"))
			}
		}
		heading := escapeMarkdownHeading(table.name)
		if len(blocks) == 0 {
			out = append(out, "## "+heading)
			continue
		}
		out = append(out, "## "+heading+"\n\n"+strings.Join(blocks, "\n\n"))
	}
	if sheetNames := biffWorkbookSheetNamesMarkdownPart(data); sheetNames != "" {
		out = append(out, sheetNames)
	}
	markdown := strings.Join(out, "\n\n")
	if len(textParts) == 0 {
		textParts = biffText(data)
	}
	if extra := missingMarkdownTextXLS(markdown, strings.Join(textParts, "\n")); extra != "" {
		out = append(out, "## Workbook Text\n\n"+extra)
	}
	return strings.Join(out, "\n\n")
}

func biffWorkbookSheetNamesMarkdownPart(data []byte) string {
	names := biffBoundSheetNamesFromRecords(data, 1252)
	if len(names) == 0 {
		return ""
	}
	var lines []string
	seen := map[string]bool{}
	for _, name := range names {
		name = markdownText(name)
		if name == "" || seen[name] {
			continue
		}
		seen[name] = true
		lines = append(lines, "- "+name)
	}
	if len(lines) == 0 {
		return ""
	}
	return "## Workbook Sheets\n\n" + strings.Join(lines, "\n")
}

type biffMarkdownTable struct {
	name         string
	rows         [][]string
	headerFooter []string
	comments     []biffCommentItem
}

type biffSheetInfo struct {
	name   string
	hidden bool
}

type biffPendingFormulaString struct {
	row int
	col int
	ok  bool
}

type biffPendingTXOComment struct {
	row    int
	col    int
	remain int
	text   string
	ok     bool
}

type biffCommentItem struct {
	ref  string
	text string
	row  int
	col  int
}

type biffTextPart struct {
	text string
	row  int
	col  int
	cell bool
	hide bool
}

func biffWorksheetMarkdownTables(data []byte) []biffMarkdownTable {
	var shared []string
	var sheets []biffSheetInfo
	var tables []biffMarkdownTable
	var rows map[int]map[int]string
	var headerFooter []string
	var comments []biffCommentItem
	var hiddenRows map[int]bool
	var hiddenCols []intRange
	var pendingFormulaString biffPendingFormulaString
	var pendingCommentRefs []biffCommentItem
	var pendingTXOComment biffPendingTXOComment
	sheetIndex := 0
	currentName := ""
	currentSheetHidden := false
	inWorksheet := false
	codePage := uint16(1252)
	biff8 := true
	flush := func() {
		pendingFormulaString = biffPendingFormulaString{}
		pendingCommentRefs = nil
		pendingTXOComment = biffPendingTXOComment{}
		if len(rows) == 0 && len(headerFooter) == 0 && len(comments) == 0 {
			if inWorksheet && !currentSheetHidden {
				name := currentName
				if name == "" {
					name = fmt.Sprintf("Sheet %d", len(tables)+1)
				}
				tables = append(tables, biffMarkdownTable{name: name})
			}
			rows = nil
			headerFooter = nil
			comments = nil
			hiddenRows = nil
			hiddenCols = nil
			inWorksheet = false
			return
		}
		tableRows := biffMarkdownRows(rows)
		if len(tableRows) > 0 || len(headerFooter) > 0 || len(comments) > 0 {
			name := currentName
			if name == "" {
				name = fmt.Sprintf("Sheet %d", len(tables)+1)
			}
			tables = append(tables, biffMarkdownTable{name: name, rows: tableRows, headerFooter: append([]string(nil), headerFooter...), comments: append([]biffCommentItem(nil), comments...)})
		}
		rows = nil
		headerFooter = nil
		comments = nil
		hiddenRows = nil
		hiddenCols = nil
		inWorksheet = false
	}
	startSheet := func() {
		flush()
		inWorksheet = true
		hidden := false
		if sheetIndex < len(sheets) {
			currentName = sheets[sheetIndex].name
			hidden = sheets[sheetIndex].hidden
		} else {
			currentName = fmt.Sprintf("Sheet %d", sheetIndex+1)
		}
		sheetIndex++
		currentSheetHidden = hidden
		if hidden {
			rows = nil
			headerFooter = nil
			comments = nil
			hiddenRows = nil
			hiddenCols = nil
			return
		}
		rows = map[int]map[int]string{}
		headerFooter = nil
		comments = nil
		hiddenRows = map[int]bool{}
		hiddenCols = nil
	}
	for off := 0; off+4 <= len(data); {
		id := binary.LittleEndian.Uint16(data[off:])
		size := int(binary.LittleEndian.Uint16(data[off+2:]))
		off += 4
		if off+size > len(data) {
			break
		}
		rec := data[off : off+size]
		if id != 0x0207 {
			pendingFormulaString = biffPendingFormulaString{}
		}
		if id != 0x003c && pendingTXOComment.ok {
			pendingTXOComment = biffPendingTXOComment{}
		}
		switch id {
		case 0x0009, 0x0209, 0x0409, 0x0809:
			biff8 = isBIFF8BOF(id, rec)
			if biffBOFType(rec) == 0x0010 {
				startSheet()
			}
		case 0x000a:
			flush()
		case 0x0042:
			if len(rec) >= 2 {
				codePage = uint16(binary.LittleEndian.Uint16(rec))
			}
		case 0x0014, 0x0015:
			if inWorksheet && !currentSheetHidden {
				if s, ok := parseBIFFHeaderFooterString(rec, biff8, codePage); ok {
					s = cleanExcelHeaderFooterText(s)
					if s != "" {
						headerFooter = append(headerFooter, s)
					}
				}
			}
		case 0x007d:
			if inWorksheet && !currentSheetHidden {
				if r, ok := biffHiddenColumnRange(rec); ok {
					hiddenCols = append(hiddenCols, r)
					biffDeleteHiddenColumnCells(rows, r)
					comments = biffFilterCommentsNotInColumnRange(comments, r)
				}
			}
		case 0x0085:
			if sheet, ok := parseBIFFBoundSheetRecord(rec, biff8, codePage); ok {
				sheets = append(sheets, sheet)
			}
		case 0x00fc:
			if biff8 {
				shared = parseSST(rec)
			}
		case 0x00fd:
			if biff8 && len(rec) >= 10 && len(shared) > 0 && biffCellRecordInMarkdownBounds(rec, hiddenRows, hiddenCols) {
				idx := int(binary.LittleEndian.Uint32(rec[6:]))
				if idx >= 0 && idx < len(shared) {
					biffSetMarkdownCell(rows, rec, shared[idx], hiddenRows, hiddenCols)
				}
			}
		case 0x0204, 0x00d6:
			if len(rec) >= 8 && biffCellRecordInMarkdownBounds(rec, hiddenRows, hiddenCols) {
				var s string
				var ok bool
				if biff8 {
					s, ok = parseXLUnicodeString(rec[6:])
				} else {
					s, ok = parseBIFFLegacyString(rec[6:], codePage)
				}
				if ok {
					biffSetMarkdownCell(rows, rec, s, hiddenRows, hiddenCols)
				}
			}
		case 0x0203:
			if len(rec) >= 14 && biffCellRecordInMarkdownBounds(rec, hiddenRows, hiddenCols) {
				if value, ok := biffNumberDisplayValue(rec); ok {
					biffSetMarkdownCell(rows, rec, value, hiddenRows, hiddenCols)
				}
			}
		case 0x0006:
			if biffCellRecordInMarkdownBounds(rec, hiddenRows, hiddenCols) {
				if value, ok := biffFormulaDisplayValue(rec); ok {
					biffSetMarkdownCell(rows, rec, value, hiddenRows, hiddenCols)
				} else if row, col, ok := biffFormulaStringCell(rec, hiddenRows, hiddenCols); ok {
					pendingFormulaString = biffPendingFormulaString{row: row, col: col, ok: true}
				}
			} else {
				pendingFormulaString = biffPendingFormulaString{}
			}
		case 0x0207:
			if pendingFormulaString.ok {
				if s, ok := parseBIFFCellString(rec, biff8, codePage); ok {
					biffSetMarkdownCellAt(rows, pendingFormulaString.row, pendingFormulaString.col, s, hiddenRows, hiddenCols)
				}
				pendingFormulaString = biffPendingFormulaString{}
			}
		case 0x0205:
			if biffCellRecordInMarkdownBounds(rec, hiddenRows, hiddenCols) {
				if value, ok := biffBoolErrDisplayValue(rec); ok {
					biffSetMarkdownCell(rows, rec, value, hiddenRows, hiddenCols)
				}
			}
		case 0x027e:
			if len(rec) >= 10 && biffCellRecordInMarkdownBounds(rec, hiddenRows, hiddenCols) {
				if value, ok := biffRKDisplayValue(rec[6:]); ok {
					biffSetMarkdownCell(rows, rec, value, hiddenRows, hiddenCols)
				}
			}
		case 0x00bd:
			biffSetMarkdownMulRKCells(rows, rec, hiddenRows, hiddenCols)
		case 0x0208:
			if inWorksheet && !currentSheetHidden {
				if row, ok := biffHiddenRow(rec); ok {
					hiddenRows[row] = true
					delete(rows, row+1)
					comments = biffFilterCommentsNotInRow(comments, row)
				}
			}
		case 0x001c:
			if inWorksheet && !currentSheetHidden {
				if row, col, ok := parseBIFFNoteCell(rec); ok && !biffCellHidden(row, col, hiddenRows, hiddenCols) {
					pendingCommentRefs = append(pendingCommentRefs, biffCommentItem{ref: biffCellRef(row, col), row: row, col: col})
				}
			}
		case 0x01b6:
			if inWorksheet && !currentSheetHidden && len(pendingCommentRefs) > 0 {
				textLen, ok := parseBIFFTXOTextLength(rec)
				comment := pendingCommentRefs[0]
				pendingCommentRefs = pendingCommentRefs[1:]
				if ok && textLen > 0 && !biffCellHidden(comment.row, comment.col, hiddenRows, hiddenCols) {
					pendingTXOComment = biffPendingTXOComment{row: comment.row, col: comment.col, remain: textLen, ok: true}
				}
			}
		case 0x003c:
			if inWorksheet && !currentSheetHidden && pendingTXOComment.ok {
				if text, consumed, ok := parseBIFFTXOContinueTextChunk(rec, pendingTXOComment.remain); ok {
					pendingTXOComment.text += text
					pendingTXOComment.remain -= consumed
				}
				if pendingTXOComment.remain <= 0 {
					text := cleanBIFFCommentText(pendingTXOComment.text)
					if text != "" && !biffCellHidden(pendingTXOComment.row, pendingTXOComment.col, hiddenRows, hiddenCols) {
						comments = append(comments, biffCommentItem{ref: biffCellRef(pendingTXOComment.row, pendingTXOComment.col), text: text, row: pendingTXOComment.row, col: pendingTXOComment.col})
					}
					pendingTXOComment = biffPendingTXOComment{}
				}
			}
		}
		off += size
	}
	flush()
	return tables
}

func biffBOFType(rec []byte) uint16 {
	if len(rec) < 4 {
		return 0
	}
	return binary.LittleEndian.Uint16(rec[2:])
}

func parseBIFFBoundSheetRecord(rec []byte, biff8 bool, codePage uint16) (biffSheetInfo, bool) {
	if len(rec) < 8 {
		return biffSheetInfo{}, false
	}
	var s string
	var ok bool
	if biff8 || looksLikeBIFF8BoundSheetName(rec[6:]) {
		s, ok = parseXLShortUnicodeString(rec[6:])
	} else {
		s, ok = parseBIFFLegacyShortString(rec[6:], codePage)
	}
	if !ok {
		return biffSheetInfo{}, false
	}
	s = cleanBIFFSheetName(s)
	return biffSheetInfo{name: s, hidden: rec[4] != 0}, true
}

func cleanBIFFSheetName(s string) string {
	s = cleanText(s)
	s = stripInlineHiddenOfficeReferences(s)
	if s == "" || looksLikeBinaryControlFragment(s) || looksLikeHiddenResourceReference(s) || looksLikeRelationshipIDReference(s) || looksLikeOfficeRelationshipMetadataReference(s) || looksLikeOfficeXMLMetadataReference(s) {
		return ""
	}
	return s
}

func biffSetMarkdownCell(rows map[int]map[int]string, rec []byte, value string, hiddenRows map[int]bool, hiddenCols []intRange) {
	if rows == nil || len(rec) < 6 {
		return
	}
	row := int(binary.LittleEndian.Uint16(rec[0:]))
	col := int(binary.LittleEndian.Uint16(rec[2:]))
	biffSetMarkdownCellAt(rows, row, col, value, hiddenRows, hiddenCols)
}

func biffSetMarkdownCellAt(rows map[int]map[int]string, row, col int, value string, hiddenRows map[int]bool, hiddenCols []intRange) {
	if rows == nil {
		return
	}
	if row < 0 || col < 0 || col >= maxMarkdownTableCols {
		return
	}
	if biffCellHidden(row, col, hiddenRows, hiddenCols) {
		return
	}
	displayRow := row + 1
	displayCol := col + 1
	if displayRow > maxMarkdownTableRows {
		return
	}
	value = cleanMarkdownTableCellValue(value)
	if value == "" || looksLikeEmbeddedPDFText(value) || looksLikeSpreadsheetFormulaExpression(value) {
		return
	}
	if rows[displayRow] == nil {
		rows[displayRow] = map[int]string{}
	}
	rows[displayRow][displayCol] = prepareMarkdownTableCellValue(value)
}

func biffSetMarkdownMulRKCells(rows map[int]map[int]string, rec []byte, hiddenRows map[int]bool, hiddenCols []intRange) {
	row, firstCol, values, ok := biffMulRKValues(rec)
	if !ok {
		return
	}
	for i, value := range values {
		biffSetMarkdownCellAt(rows, row, firstCol+i, value, hiddenRows, hiddenCols)
	}
}

func biffMarkdownRows(values map[int]map[int]string) [][]string {
	if len(values) == 0 {
		return nil
	}
	rowIndexes := make([]int, 0, len(values))
	for row, cols := range values {
		if row > 0 && len(cols) > 0 {
			rowIndexes = append(rowIndexes, row)
		}
	}
	sort.Ints(rowIndexes)
	if len(rowIndexes) > maxMarkdownTableRows {
		rowIndexes = rowIndexes[:maxMarkdownTableRows]
	}
	out := make([][]string, 0, len(rowIndexes))
	for _, row := range rowIndexes {
		out = append(out, compactWorksheetMarkdownRow(values[row]))
	}
	return out
}

func biffDeleteHiddenColumnCells(rows map[int]map[int]string, r intRange) {
	for row, cols := range rows {
		for col := range cols {
			if col-1 >= r.min && col-1 <= r.max {
				delete(cols, col)
			}
		}
		if len(cols) == 0 {
			delete(rows, row)
		}
	}
}

func biffHiddenRow(rec []byte) (int, bool) {
	if len(rec) < 14 {
		return 0, false
	}
	row := int(binary.LittleEndian.Uint16(rec[0:]))
	grbit := binary.LittleEndian.Uint16(rec[12:])
	return row, grbit&0x0020 != 0
}

func biffHiddenColumnRange(rec []byte) (intRange, bool) {
	if len(rec) < 10 {
		return intRange{}, false
	}
	grbit := binary.LittleEndian.Uint16(rec[8:])
	if grbit&0x0001 == 0 {
		return intRange{}, false
	}
	first := int(binary.LittleEndian.Uint16(rec[0:]))
	last := int(binary.LittleEndian.Uint16(rec[2:]))
	if first < 0 || last < first {
		return intRange{}, false
	}
	return intRange{min: first, max: last}, true
}

func biffCellHidden(row, col int, hiddenRows map[int]bool, hiddenCols []intRange) bool {
	if hiddenRows[row] {
		return true
	}
	for _, r := range hiddenCols {
		if col >= r.min && col <= r.max {
			return true
		}
	}
	return false
}

func biffCellInMarkdownBounds(row, col int, hiddenRows map[int]bool, hiddenCols []intRange) bool {
	if row < 0 || col < 0 || col >= maxMarkdownTableCols {
		return false
	}
	if row+1 > maxMarkdownTableRows {
		return false
	}
	return !biffCellHidden(row, col, hiddenRows, hiddenCols)
}

func biffCellRecordInMarkdownBounds(rec []byte, hiddenRows map[int]bool, hiddenCols []intRange) bool {
	if len(rec) < 4 {
		return false
	}
	row := int(binary.LittleEndian.Uint16(rec[0:]))
	col := int(binary.LittleEndian.Uint16(rec[2:]))
	return biffCellInMarkdownBounds(row, col, hiddenRows, hiddenCols)
}

func looksLikePPTDesignThemeLabel(s string) bool {
	lower := strings.ToLower(strings.TrimSpace(s))
	switch lower {
	case "default design", "office theme", "office テーマ":
		return true
	default:
		return false
	}
}

func parseBIFFNoteCell(rec []byte) (int, int, bool) {
	if len(rec) < 4 {
		return 0, 0, false
	}
	row := int(binary.LittleEndian.Uint16(rec[0:]))
	col := int(binary.LittleEndian.Uint16(rec[2:]))
	if row < 0 || col < 0 || col >= maxMarkdownTableCols {
		return 0, 0, false
	}
	return row, col, true
}

func parseBIFFTXOTextLength(rec []byte) (int, bool) {
	if len(rec) < 12 {
		return 0, false
	}
	textLen := int(binary.LittleEndian.Uint16(rec[10:]))
	if textLen <= 0 || textLen > maxMarkdownTableCellBytes {
		return 0, false
	}
	return textLen, true
}

func parseBIFFTXOContinueTextChunk(rec []byte, maxChars int) (string, int, bool) {
	if len(rec) < 1 || maxChars <= 0 {
		return "", 0, false
	}
	flags := rec[0]
	available := len(rec) - 1
	if flags&0x01 != 0 {
		available /= 2
	}
	cch := maxChars
	if available < cch {
		cch = available
	}
	if cch <= 0 {
		return "", 0, false
	}
	text, n, ok := parseXLUnicodeStringPayload(rec[1:], cch, flags)
	if !ok || n > len(rec)-1 {
		return "", 0, false
	}
	return text, cch, true
}

func cleanBIFFCommentText(text string) string {
	text = cleanVisibleText(text)
	if text == "" || looksLikeEmbeddedPDFText(text) || looksLikeSpreadsheetFormulaExpression(text) {
		return ""
	}
	if strings.Contains(text, "\n") {
		var out []string
		for _, line := range strings.Split(text, "\n") {
			line = cleanVisibleText(line)
			if line == "" || looksLikeEmbeddedPDFText(line) || looksLikeSpreadsheetFormulaExpression(line) {
				continue
			}
			out = append(out, line)
		}
		return cleanText(strings.Join(out, "\n"))
	}
	return text
}

func biffFilterCommentsNotInRow(comments []biffCommentItem, row int) []biffCommentItem {
	if len(comments) == 0 {
		return comments
	}
	kept := comments[:0]
	for _, comment := range comments {
		if comment.row == row {
			continue
		}
		kept = append(kept, comment)
	}
	return kept
}

func biffFilterCommentsNotInColumnRange(comments []biffCommentItem, r intRange) []biffCommentItem {
	if len(comments) == 0 {
		return comments
	}
	kept := comments[:0]
	for _, comment := range comments {
		if comment.col >= r.min && comment.col <= r.max {
			continue
		}
		kept = append(kept, comment)
	}
	return kept
}

func biffCellRef(row, col int) string {
	return biffColumnName(col+1) + strconv.Itoa(row+1)
}

func biffColumnName(col int) string {
	if col <= 0 {
		return ""
	}
	var buf [8]byte
	i := len(buf)
	for col > 0 {
		col--
		i--
		buf[i] = byte('A' + col%26)
		col /= 26
	}
	return string(buf[i:])
}

func biffBoolErrDisplayValue(rec []byte) (string, bool) {
	if len(rec) < 8 {
		return "", false
	}
	value := rec[6]
	isErr := rec[7] != 0
	if !isErr {
		if value == 0 {
			return "FALSE", true
		}
		return "TRUE", true
	}
	return biffErrorDisplayValue(value)
}

func biffErrorDisplayValue(value byte) (string, bool) {
	switch value {
	case 0x00:
		return "#NULL!", true
	case 0x07:
		return "#DIV/0!", true
	case 0x0f:
		return "#VALUE!", true
	case 0x17:
		return "#REF!", true
	case 0x1d:
		return "#NAME?", true
	case 0x24:
		return "#NUM!", true
	case 0x2a:
		return "#N/A", true
	default:
		return "", false
	}
}

func biffFormulaDisplayValue(rec []byte) (string, bool) {
	if len(rec) < 14 {
		return "", false
	}
	raw := rec[6:14]
	if bytes.Equal(raw[6:], []byte{0xff, 0xff}) {
		switch raw[0] {
		case 0x01:
			if raw[2] == 0 {
				return "FALSE", true
			}
			return "TRUE", true
		case 0x02:
			return biffErrorDisplayValue(raw[2])
		default:
			return "", false
		}
	}
	value := math.Float64frombits(binary.LittleEndian.Uint64(raw))
	if math.IsNaN(value) || math.IsInf(value, 0) {
		return "", false
	}
	return strconv.FormatFloat(value, 'f', -1, 64), true
}

func biffNumberDisplayValue(rec []byte) (string, bool) {
	if len(rec) < 14 {
		return "", false
	}
	value := math.Float64frombits(binary.LittleEndian.Uint64(rec[6:]))
	return finiteFloatDisplayValue(value)
}

func biffRKDisplayValue(rec []byte) (string, bool) {
	if len(rec) < 4 {
		return "", false
	}
	return finiteFloatDisplayValue(decodeBIFFRK(binary.LittleEndian.Uint32(rec)))
}

func finiteFloatDisplayValue(value float64) (string, bool) {
	if math.IsNaN(value) || math.IsInf(value, 0) {
		return "", false
	}
	return strconv.FormatFloat(value, 'f', -1, 64), true
}

func biffFormulaStringCell(rec []byte, hiddenRows map[int]bool, hiddenCols []intRange) (int, int, bool) {
	if len(rec) < 14 || !bytes.Equal(rec[12:14], []byte{0xff, 0xff}) || rec[6] != 0x00 {
		return 0, 0, false
	}
	row := int(binary.LittleEndian.Uint16(rec[0:]))
	col := int(binary.LittleEndian.Uint16(rec[2:]))
	if !biffCellInMarkdownBounds(row, col, hiddenRows, hiddenCols) {
		return 0, 0, false
	}
	return row, col, true
}

func parseBIFFCellString(rec []byte, biff8 bool, codePage uint16) (string, bool) {
	if biff8 {
		return parseXLUnicodeString(rec)
	}
	return parseBIFFLegacyString(rec, codePage)
}

func biffMulRKValues(rec []byte) (int, int, []string, bool) {
	if len(rec) < 10 {
		return 0, 0, nil, false
	}
	row := int(binary.LittleEndian.Uint16(rec[0:]))
	firstCol := int(binary.LittleEndian.Uint16(rec[2:]))
	lastCol := int(binary.LittleEndian.Uint16(rec[len(rec)-2:]))
	if firstCol < 0 || lastCol < firstCol {
		return 0, 0, nil, false
	}
	count := lastCol - firstCol + 1
	if count <= 0 || 4+count*6+2 != len(rec) {
		return 0, 0, nil, false
	}
	values := make([]string, 0, count)
	for i := 0; i < count; i++ {
		pos := 4 + i*6
		value, ok := biffRKDisplayValue(rec[pos+2:])
		if !ok {
			value = ""
		}
		values = append(values, value)
	}
	return row, firstCol, values, true
}

func decodeBIFFRK(rk uint32) float64 {
	var v float64
	if rk&0x02 != 0 {
		v = float64(int32(rk) >> 2)
	} else {
		v = math.Float64frombits(uint64(rk&0xfffffffc) << 32)
	}
	if rk&0x01 != 0 {
		v /= 100
	}
	return v
}

func biffText(data []byte) []string {
	return biffVisibleTextFromParts(biffTextParts(data))
}

func biffVisibleTextFromParts(parts []biffTextPart) []string {
	printerSettingsDump := biffTextPartsLookLikePrinterSettingsDump(parts)
	out := make([]string, 0, len(parts))
	for _, part := range parts {
		if part.hide {
			continue
		}
		if printerSettingsDump && looksLikePrinterDriverSettingFragment(part.text) {
			continue
		}
		out = append(out, part.text)
	}
	return out
}

func biffTextParts(data []byte) []biffTextPart {
	var shared []string
	var sheets []biffSheetInfo
	var parts []biffTextPart
	rowPartIndexes := map[int][]int{}
	codePage := uint16(1252)
	biff8 := true
	sheetIndex := 0
	currentSheetHidden := false
	hiddenRows := map[int]bool{}
	var hiddenCols []intRange
	var pendingFormulaString biffPendingFormulaString
	var pendingCommentRefs []biffCommentItem
	var pendingTXOComment biffPendingTXOComment
	seenLargeSharedIndexes := map[int]bool{}
	inWorksheet := false
	appendPart := func(part biffTextPart) {
		parts = append(parts, part)
		if part.cell {
			rowPartIndexes[part.row] = append(rowPartIndexes[part.row], len(parts)-1)
		}
	}
	addText := func(s string) {
		forEachBIFFText(s, func(text string) {
			appendPart(biffTextPart{text: text})
		})
	}
	addCellText := func(rec []byte, s string) {
		row := 0
		if len(rec) >= 2 {
			row = int(binary.LittleEndian.Uint16(rec[0:]))
		}
		col := 0
		if len(rec) >= 4 {
			col = int(binary.LittleEndian.Uint16(rec[2:]))
		}
		forEachBIFFText(s, func(text string) {
			appendPart(biffTextPart{text: text, row: row, col: col, cell: true})
		})
	}
	removeHiddenRowText := func(row int) {
		if len(parts) == 0 {
			return
		}
		for _, idx := range rowPartIndexes[row] {
			if idx >= 0 && idx < len(parts) {
				parts[idx].hide = true
			}
		}
		delete(rowPartIndexes, row)
	}
	removeHiddenColumnText := func(r intRange) {
		if len(parts) == 0 {
			return
		}
		kept := parts[:0]
		for _, part := range parts {
			if part.cell && part.col >= r.min && part.col <= r.max {
				continue
			}
			kept = append(kept, part)
		}
		parts = kept
	}
	addCommentText := func(comment biffCommentItem, s string) {
		if biffCellHidden(comment.row, comment.col, hiddenRows, hiddenCols) {
			return
		}
		forEachBIFFText(s, func(text string) {
			appendPart(biffTextPart{text: text, row: comment.row, col: comment.col, cell: true})
		})
	}
	for off := 0; off+4 <= len(data); {
		id := binary.LittleEndian.Uint16(data[off:])
		size := int(binary.LittleEndian.Uint16(data[off+2:]))
		off += 4
		if off+size > len(data) {
			break
		}
		rec := data[off : off+size]
		if id != 0x0207 {
			pendingFormulaString = biffPendingFormulaString{}
		}
		if id != 0x003c && pendingTXOComment.ok {
			pendingTXOComment = biffPendingTXOComment{}
		}
		switch id {
		case 0x0009, 0x0209, 0x0409, 0x0809:
			biff8 = isBIFF8BOF(id, rec)
			currentSheetHidden = false
			hiddenRows = map[int]bool{}
			rowPartIndexes = map[int][]int{}
			hiddenCols = nil
			pendingFormulaString = biffPendingFormulaString{}
			pendingCommentRefs = nil
			pendingTXOComment = biffPendingTXOComment{}
			inWorksheet = false
			if biffBOFType(rec) == 0x0010 {
				inWorksheet = true
				if sheetIndex < len(sheets) {
					currentSheetHidden = sheets[sheetIndex].hidden
				}
				sheetIndex++
			}
		case 0x000a:
			currentSheetHidden = false
			hiddenRows = map[int]bool{}
			rowPartIndexes = map[int][]int{}
			hiddenCols = nil
			pendingFormulaString = biffPendingFormulaString{}
			pendingCommentRefs = nil
			pendingTXOComment = biffPendingTXOComment{}
			inWorksheet = false
		case 0x0042:
			if len(rec) >= 2 {
				codePage = uint16(binary.LittleEndian.Uint16(rec))
			}
		case 0x00fc:
			if biff8 {
				shared = parseSST(rec)
			}
		case 0x00fd:
			if inWorksheet && !currentSheetHidden && biff8 && len(rec) >= 10 && len(shared) > 0 && !biffBIFFCellRecordHidden(rec, hiddenRows, hiddenCols) {
				idx := int(binary.LittleEndian.Uint32(rec[6:]))
				if idx >= 0 && idx < len(shared) {
					if len(shared[idx]) > maxRepeatedTextPartBytes {
						if seenLargeSharedIndexes[idx] {
							off += size
							continue
						}
						seenLargeSharedIndexes[idx] = true
					}
					addCellText(rec, shared[idx])
				}
			}
		case 0x0204, 0x00d6:
			if inWorksheet && !currentSheetHidden && len(rec) >= 8 && !biffBIFFCellRecordHidden(rec, hiddenRows, hiddenCols) {
				var s string
				var ok bool
				if biff8 {
					s, ok = parseXLUnicodeString(rec[6:])
				} else {
					s, ok = parseBIFFLegacyString(rec[6:], codePage)
				}
				if ok {
					addCellText(rec, s)
				}
			}
		case 0x0006:
			if inWorksheet && !currentSheetHidden && !biffBIFFCellRecordHidden(rec, hiddenRows, hiddenCols) {
				if value, ok := biffFormulaDisplayValue(rec); ok {
					addCellText(rec, value)
				} else if row, col, ok := biffFormulaStringCell(rec, hiddenRows, hiddenCols); ok {
					pendingFormulaString = biffPendingFormulaString{row: row, col: col, ok: true}
				}
			}
		case 0x0207:
			if inWorksheet && !currentSheetHidden && pendingFormulaString.ok {
				if s, ok := parseBIFFCellString(rec, biff8, codePage); ok && !biffCellHidden(pendingFormulaString.row, pendingFormulaString.col, hiddenRows, hiddenCols) {
					forEachBIFFText(s, func(text string) {
						appendPart(biffTextPart{text: text, row: pendingFormulaString.row, col: pendingFormulaString.col, cell: true})
					})
				}
				pendingFormulaString = biffPendingFormulaString{}
			}
		case 0x0203:
			if inWorksheet && !currentSheetHidden && len(rec) >= 14 && !biffBIFFCellRecordHidden(rec, hiddenRows, hiddenCols) {
				if value, ok := finiteFloatDisplayValue(math.Float64frombits(binary.LittleEndian.Uint64(rec[6:]))); ok {
					addCellText(rec, value)
				}
			}
		case 0x0205:
			if inWorksheet && !currentSheetHidden && !biffBIFFCellRecordHidden(rec, hiddenRows, hiddenCols) {
				if value, ok := biffBoolErrDisplayValue(rec); ok {
					addCellText(rec, value)
				}
			}
		case 0x027e:
			if inWorksheet && !currentSheetHidden && len(rec) >= 10 && !biffBIFFCellRecordHidden(rec, hiddenRows, hiddenCols) {
				if value, ok := finiteFloatDisplayValue(decodeBIFFRK(binary.LittleEndian.Uint32(rec[6:]))); ok {
					addCellText(rec, value)
				}
			}
		case 0x00bd:
			if inWorksheet && !currentSheetHidden {
				biffAddMulRKTextParts(&parts, rowPartIndexes, rec, hiddenRows, hiddenCols)
			}
		case 0x0014, 0x0015:
			if inWorksheet && !currentSheetHidden {
				if s, ok := parseBIFFHeaderFooterString(rec, biff8, codePage); ok {
					addText(cleanExcelHeaderFooterText(s))
				}
			}
		case 0x007d:
			if inWorksheet && !currentSheetHidden {
				if r, ok := biffHiddenColumnRange(rec); ok {
					hiddenCols = append(hiddenCols, r)
					removeHiddenColumnText(r)
				}
			}
		case 0x0085:
			if sheet, ok := parseBIFFBoundSheetRecord(rec, biff8, codePage); ok {
				sheets = append(sheets, sheet)
				if !sheet.hidden {
					addText(sheet.name)
				}
			}
		case 0x0208:
			if inWorksheet && !currentSheetHidden {
				if row, ok := biffHiddenRow(rec); ok {
					hiddenRows[row] = true
					removeHiddenRowText(row)
				}
			}
		case 0x001c:
			if inWorksheet && !currentSheetHidden {
				if row, col, ok := parseBIFFNoteCell(rec); ok && !biffCellHidden(row, col, hiddenRows, hiddenCols) {
					pendingCommentRefs = append(pendingCommentRefs, biffCommentItem{ref: biffCellRef(row, col), row: row, col: col})
				}
			}
		case 0x01b6:
			if inWorksheet && !currentSheetHidden && len(pendingCommentRefs) > 0 {
				textLen, ok := parseBIFFTXOTextLength(rec)
				comment := pendingCommentRefs[0]
				pendingCommentRefs = pendingCommentRefs[1:]
				if ok && textLen > 0 && !biffCellHidden(comment.row, comment.col, hiddenRows, hiddenCols) {
					pendingTXOComment = biffPendingTXOComment{row: comment.row, col: comment.col, remain: textLen, ok: true}
				}
			}
		case 0x003c:
			if inWorksheet && !currentSheetHidden && pendingTXOComment.ok {
				comment := biffCommentItem{ref: biffCellRef(pendingTXOComment.row, pendingTXOComment.col), row: pendingTXOComment.row, col: pendingTXOComment.col}
				if text, consumed, ok := parseBIFFTXOContinueTextChunk(rec, pendingTXOComment.remain); ok {
					pendingTXOComment.text += text
					pendingTXOComment.remain -= consumed
				}
				if pendingTXOComment.remain <= 0 {
					addCommentText(comment, cleanBIFFCommentText(pendingTXOComment.text))
					pendingTXOComment = biffPendingTXOComment{}
				}
			}
		}
		off += size
	}
	return parts
}

func biffTextPartsLookLikePrinterSettingsDump(parts []biffTextPart) bool {
	hits := 0
	for _, part := range parts {
		if looksLikePrinterDriverSettingFragment(part.text) {
			hits++
			if hits >= 5 {
				return true
			}
		}
	}
	return false
}

func stringSliceLooksLikePrinterSettingsDump(parts []string) bool {
	hits := 0
	for _, part := range parts {
		if looksLikePrinterDriverSettingFragment(part) {
			hits++
			if hits >= 5 {
				return true
			}
		}
	}
	return false
}

func looksLikePrinterDriverSettingFragment(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	lower := strings.ToLower(s)
	if strings.Contains(lower, "laserjet") || strings.Contains(lower, "printer") {
		return true
	}
	if strings.HasPrefix(lower, "hp") || strings.HasPrefix(lower, "pcl") || strings.HasPrefix(lower, "ps") || strings.HasPrefix(lower, "jr") {
		return true
	}
	switch lower {
	case "inputbin", "formsource", "resdll", "uniresdll", "resolution", "fastres",
		"orientation", "portrait", "duplex", "papersize", "letter", "mediatype",
		"collate", "outputbin", "stapling", "economode", "textasblack", "jpegenable",
		"retchoice", "alternateletterhead", "printqualitygroup", "modeless",
		"evenpagesfirst", "driverrotate", "front_cover", "excel.exe":
		return true
	default:
		return false
	}
}

func biffBIFFCellRecordHidden(rec []byte, hiddenRows map[int]bool, hiddenCols []intRange) bool {
	if len(rec) < 4 {
		return false
	}
	row := int(binary.LittleEndian.Uint16(rec[0:]))
	col := int(binary.LittleEndian.Uint16(rec[2:]))
	return biffCellHidden(row, col, hiddenRows, hiddenCols)
}

func biffAddMulRKText(out *[]string, rec []byte, hiddenRows map[int]bool, hiddenCols []intRange) {
	row, firstCol, count, ok := biffMulRKRange(rec)
	if !ok {
		return
	}
	for i := 0; i < count; i++ {
		if biffCellHidden(row, firstCol+i, hiddenRows, hiddenCols) {
			continue
		}
		pos := 4 + i*6
		if value, ok := biffRKDisplayValue(rec[pos+2:]); ok {
			addBIFFText(out, value)
		}
	}
}

func biffAddMulRKTextParts(parts *[]biffTextPart, rowPartIndexes map[int][]int, rec []byte, hiddenRows map[int]bool, hiddenCols []intRange) {
	row, firstCol, count, ok := biffMulRKRange(rec)
	if !ok {
		return
	}
	for i := 0; i < count; i++ {
		if biffCellHidden(row, firstCol+i, hiddenRows, hiddenCols) {
			continue
		}
		pos := 4 + i*6
		if value, ok := biffRKDisplayValue(rec[pos+2:]); ok {
			forEachBIFFText(value, func(text string) {
				*parts = append(*parts, biffTextPart{text: text, row: row, col: firstCol + i, cell: true})
				rowPartIndexes[row] = append(rowPartIndexes[row], len(*parts)-1)
			})
		}
	}
}

func biffMulRKRange(rec []byte) (int, int, int, bool) {
	if len(rec) < 10 {
		return 0, 0, 0, false
	}
	row := int(binary.LittleEndian.Uint16(rec[0:]))
	firstCol := int(binary.LittleEndian.Uint16(rec[2:]))
	lastCol := int(binary.LittleEndian.Uint16(rec[len(rec)-2:]))
	if firstCol < 0 || lastCol < firstCol {
		return 0, 0, 0, false
	}
	count := lastCol - firstCol + 1
	if count <= 0 || 4+count*6+2 != len(rec) {
		return 0, 0, 0, false
	}
	return row, firstCol, count, true
}

func parseBIFFHeaderFooterString(data []byte, biff8 bool, codePage uint16) (string, bool) {
	if biff8 {
		return parseXLUnicodeString(data)
	}
	return parseBIFFLegacyString(data, codePage)
}

func biffBoundSheetNamesFromRecords(data []byte, codePage uint16) []string {
	var out []string
	biff8 := true
	for off := 0; off+4 <= len(data); {
		id := binary.LittleEndian.Uint16(data[off:])
		size := int(binary.LittleEndian.Uint16(data[off+2:]))
		off += 4
		if size < 0 || off+size > len(data) {
			break
		}
		rec := data[off : off+size]
		switch id {
		case 0x0009, 0x0209, 0x0409, 0x0809:
			biff8 = isBIFF8BOF(id, rec)
		case 0x0042:
			if len(rec) >= 2 {
				codePage = uint16(binary.LittleEndian.Uint16(rec))
			}
		case 0x0085:
			if sheet, ok := parseBIFFBoundSheetRecord(rec, biff8, codePage); ok && !sheet.hidden {
				addBIFFText(&out, sheet.name)
			}
		}
		off += size
	}
	return uniqueStrings(out)
}

func biffBoundSheetNames(data []byte, codePage uint16) []string {
	var out []string
	for off := 0; off+12 <= len(data); off++ {
		i := bytes.IndexByte(data[off:], 0x85)
		if i < 0 {
			break
		}
		off += i
		if off+12 > len(data) {
			break
		}
		if binary.LittleEndian.Uint16(data[off:]) != 0x0085 {
			continue
		}
		size := int(binary.LittleEndian.Uint16(data[off+2:]))
		if size < 7 || size > 255 || off+4+size > len(data) {
			continue
		}
		rec := data[off+4 : off+4+size]
		if rec[4] > 2 || rec[5] > 2 {
			continue
		}
		if rec[4] != 0 {
			continue
		}
		var s string
		var ok bool
		if looksLikeBIFF8BoundSheetName(rec[6:]) {
			s, ok = parseXLShortUnicodeString(rec[6:])
		} else {
			s, ok = parseBIFFLegacyShortString(rec[6:], codePage)
			if ok && countCyrillicLetters(s) == 0 {
				if alt, altOK := parseBIFFLegacyShortString(rec[6:], 1251); altOK && countCyrillicLetters(alt) > countCyrillicLetters(s) {
					s = alt
				}
			}
		}
		if ok {
			addBIFFText(&out, s)
		}
	}
	return uniqueStrings(out)
}

func looksLikeBIFF8BoundSheetName(data []byte) bool {
	if len(data) < 2 {
		return false
	}
	cch := int(data[0])
	flags := data[1]
	if cch <= 0 || flags&0xf0 != 0 {
		return false
	}
	width := 1
	if flags&0x01 != 0 {
		width = 2
	}
	extra := 0
	if flags&0x08 != 0 {
		extra += 2
	}
	if flags&0x04 != 0 {
		extra += 4
	}
	return 2+extra+cch*width <= len(data)
}

func isBIFFWorkbook(data []byte) bool {
	if len(data) < 4 {
		return false
	}
	id := binary.LittleEndian.Uint16(data)
	size := int(binary.LittleEndian.Uint16(data[2:]))
	switch id {
	case 0x0009, 0x0209, 0x0409, 0x0809:
		return size >= 2 && 4+size <= len(data)
	default:
		return false
	}
}

func isBIFF8BOF(id uint16, rec []byte) bool {
	if id != 0x0809 || len(rec) < 2 {
		return false
	}
	return binary.LittleEndian.Uint16(rec) >= 0x0600
}

func addBIFFText(out *[]string, s string) {
	forEachBIFFText(s, func(text string) {
		*out = append(*out, text)
	})
}

func forEachBIFFText(s string, emit func(string)) {
	s = cleanVisibleText(s)
	if s == "" || looksLikeEmbeddedPDFText(s) || looksLikeSpreadsheetFormulaExpression(s) {
		return
	}
	if !strings.Contains(s, "\n") {
		emit(s)
		return
	}
	for _, line := range strings.Split(s, "\n") {
		line = cleanVisibleText(line)
		if line == "" || looksLikeEmbeddedPDFText(line) || looksLikeSpreadsheetFormulaExpression(line) {
			continue
		}
		emit(line)
	}
}

func looksLikeSpreadsheetFormulaExpression(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	if strings.HasPrefix(s, "=") {
		s = strings.TrimSpace(strings.TrimPrefix(s, "="))
	}
	if looksLikeSpreadsheetRangeReference(s) {
		return true
	}
	open := strings.IndexByte(s, '(')
	if open <= 0 || !strings.HasSuffix(s, ")") {
		return false
	}
	fn := s[:open]
	if len(fn) > 40 {
		return false
	}
	hasLetter := false
	for _, r := range fn {
		switch {
		case r >= 'A' && r <= 'Z':
			hasLetter = true
		case r >= '0' && r <= '9':
		case r == '_' || r == '.':
		default:
			return false
		}
	}
	return hasLetter
}

func looksLikeSpreadsheetRangeReference(s string) bool {
	bang := strings.LastIndexByte(s, '!')
	if bang <= 0 || bang >= len(s)-1 {
		return false
	}
	sheet := strings.TrimSpace(s[:bang])
	ref := strings.TrimSpace(s[bang+1:])
	if sheet == "" || ref == "" || strings.ContainsAny(sheet, "\r\n\t") || strings.ContainsAny(ref, "\r\n\t ") {
		return false
	}
	if strings.HasPrefix(sheet, "'") && strings.HasSuffix(sheet, "'") && len(sheet) > 2 {
		sheet = strings.TrimSpace(sheet[1 : len(sheet)-1])
	}
	if sheet == "" {
		return false
	}
	parts := strings.Split(ref, ":")
	if len(parts) > 2 {
		return false
	}
	for _, part := range parts {
		if !looksLikeSpreadsheetCellReference(part) {
			return false
		}
	}
	return true
}

func looksLikeSpreadsheetCellReference(s string) bool {
	s = strings.TrimPrefix(strings.TrimSpace(s), "$")
	if s == "" {
		return false
	}
	i := 0
	for i < len(s) && ((s[i] >= 'A' && s[i] <= 'Z') || (s[i] >= 'a' && s[i] <= 'z')) {
		i++
	}
	if i == 0 || i > 3 || i >= len(s) {
		return false
	}
	s = strings.TrimPrefix(s[i:], "$")
	if s == "" {
		return false
	}
	for _, r := range s {
		if r < '0' || r > '9' {
			return false
		}
	}
	return true
}

func looksLikeEmbeddedPDFText(s string) bool {
	lower := strings.ToLower(strings.TrimSpace(s))
	return strings.Contains(lower, "%pdf-") ||
		strings.Contains(lower, " /filter /flatedecode") ||
		strings.Contains(lower, "/type /page") ||
		strings.Contains(lower, "/procset") ||
		strings.Contains(lower, "endobj") ||
		strings.Contains(lower, "endstream")
}

func parseSST(data []byte) []string {
	if len(data) < 8 {
		return nil
	}
	count := int(binary.LittleEndian.Uint32(data[4:]))
	off := 8
	out := make([]string, 0, count)
	for len(out) < count && off < len(data) {
		s, n, ok := parseXLUnicodeStringWithSize(data[off:])
		if !ok || n <= 0 {
			break
		}
		out = append(out, s)
		off += n
	}
	return out
}

func parseXLUnicodeString(data []byte) (string, bool) {
	s, _, ok := parseXLUnicodeStringWithSize(data)
	return s, ok
}

func parseXLShortUnicodeString(data []byte) (string, bool) {
	if len(data) < 2 {
		return "", false
	}
	cch := int(data[0])
	flags := data[1]
	s, n, ok := parseXLUnicodeStringPayload(data[2:], cch, flags)
	if !ok || n > len(data)-2 {
		return "", false
	}
	return s, true
}

func parseBIFFLegacyString(data []byte, codePage uint16) (string, bool) {
	if len(data) < 2 {
		return "", false
	}
	cch := int(binary.LittleEndian.Uint16(data))
	if cch <= 0 || cch > len(data)-2 {
		return "", false
	}
	text := cleanText(decodeCodePageBytes(data[2:2+cch], codePage))
	return text, text != ""
}

func parseBIFFLegacyShortString(data []byte, codePage uint16) (string, bool) {
	if len(data) < 1 {
		return "", false
	}
	cch := int(data[0])
	if cch <= 0 || cch > len(data)-1 {
		return "", false
	}
	text := cleanText(decodeCodePageBytes(data[1:1+cch], codePage))
	return text, text != ""
}

func parseXLUnicodeStringWithSize(data []byte) (string, int, bool) {
	if len(data) < 3 {
		return "", 0, false
	}
	cch := int(binary.LittleEndian.Uint16(data))
	flags := data[2]
	s, n, ok := parseXLUnicodeStringPayload(data[3:], cch, flags)
	if !ok {
		return "", 0, false
	}
	return s, 3 + n, true
}

func parseXLUnicodeStringPayload(data []byte, cch int, flags byte) (string, int, bool) {
	off := 0
	richRuns := 0
	extSize := 0
	if flags&0x08 != 0 {
		if len(data) < off+2 {
			return "", 0, false
		}
		richRuns = int(binary.LittleEndian.Uint16(data[off:]))
		off += 2
	}
	if flags&0x04 != 0 {
		if len(data) < off+4 {
			return "", 0, false
		}
		extSize = int(binary.LittleEndian.Uint32(data[off:]))
		off += 4
	}
	is16 := flags&0x01 != 0
	byteLen := cch
	if is16 {
		byteLen *= 2
	}
	if len(data) < off+byteLen {
		return "", 0, false
	}
	raw := data[off : off+byteLen]
	off += byteLen
	var s string
	if is16 {
		u := make([]uint16, 0, cch)
		for i := 0; i+1 < len(raw); i += 2 {
			u = append(u, binary.LittleEndian.Uint16(raw[i:]))
		}
		s = string(utf16.Decode(u))
	} else {
		s = biffCompressedUnicodeBytesToString(raw)
	}
	off += richRuns * 4
	off += extSize
	if off > len(data) {
		return "", 0, false
	}
	return cleanText(s), off, true
}

func biffCompressedUnicodeBytesToString(raw []byte) string {
	runes := make([]rune, len(raw))
	for i, b := range raw {
		runes[i] = rune(b)
	}
	return string(runes)
}

func uniqueStrings(in []string) []string {
	seen := map[string]bool{}
	var out []string
	for _, s := range in {
		s = cleanUniqueVisibleString(s)
		if s == "" || seen[s] || looksLikeBinaryControlFragment(s) || looksLikeRelationshipIDReference(s) || looksLikeOfficeRelationshipMetadataReference(s) || looksLikeOfficeXMLMetadataReference(s) || looksLikeUniqueStringHiddenResourceReference(s) || looksLikeEmbeddedPDFText(s) {
			continue
		}
		seen[s] = true
		out = append(out, s)
	}
	return out
}

func cleanUniqueVisibleString(s string) string {
	s = cleanText(s)
	if s == "" {
		return ""
	}
	if strings.Contains(s, "\n") {
		var out []string
		for _, line := range strings.Split(s, "\n") {
			line = cleanUniqueVisibleString(line)
			if line != "" {
				out = append(out, line)
			}
		}
		return cleanText(strings.Join(out, "\n"))
	}
	s = stripInlineHiddenOfficeReferences(s)
	if s == "" || looksLikeBinaryControlFragment(s) || looksLikeRelationshipIDReference(s) || looksLikeOfficeRelationshipMetadataReference(s) || looksLikeOfficeXMLMetadataReference(s) || looksLikeUniqueStringHiddenResourceReference(s) || looksLikeEmbeddedPDFText(s) {
		return ""
	}
	return s
}

func looksLikeUniqueStringHiddenResourceReference(s string) bool {
	trimmed := strings.TrimSpace(strings.ReplaceAll(s, "\\", "/"))
	if trimmed == "" {
		return false
	}
	if looksLikeInlineHiddenResourceReferencePlain(trimmed) {
		return true
	}
	return looksLikeOfficePartPath(strings.ToLower(strings.TrimPrefix(trimmed, "/")))
}

func uniqueCleanedStrings(in []string) []string {
	seen := make(map[string]bool, len(in))
	out := make([]string, 0, len(in))
	for _, s := range in {
		if s == "" || seen[s] {
			continue
		}
		seen[s] = true
		out = append(out, s)
	}
	return out
}

func extractBinaryStrings(data []byte) []string {
	data = maskEmbeddedImagesForText(data)
	data = maskEmbeddedPDFsForText(data)
	data = maskEmbeddedOOXMLZipsForText(data)
	seen := map[string]bool{}
	var out []string
	var add func(string)
	add = func(s string) {
		s = cleanText(s)
		s = stripInlineHiddenOfficeReferences(s)
		if looksLikeEmbeddedPDFText(s) {
			return
		}
		if strings.Contains(s, "\n") {
			for _, line := range strings.Split(s, "\n") {
				add(line)
			}
			return
		}
		if len([]rune(s)) < 3 || seen[s] || !looksLikeTextFragment(s) || looksLikeBinaryControlFragment(s) || looksLikeHiddenResourceReference(s) || looksLikeRelationshipIDReference(s) || looksLikeOfficeRelationshipMetadataReference(s) || looksLikeOfficeXMLMetadataReference(s) {
			return
		}
		seen[s] = true
		out = append(out, s)
	}
	for _, s := range asciiStrings(data, 4) {
		add(s)
	}
	for _, s := range utf16Strings(data, 3) {
		add(s)
	}
	return out
}

func maskEmbeddedOOXMLZipsForText(data []byte) []byte {
	ranges := ooxmlZipByteRanges(data)
	if len(ranges) == 0 {
		return data
	}
	out := append([]byte(nil), data...)
	for _, r := range ranges {
		for i := r.start; i < r.end; i++ {
			out[i] = 0
		}
	}
	return out
}

func ooxmlZipByteRanges(data []byte) []byteRange {
	const zipMagicLen = 4
	var ranges []byteRange
	offset := 0
	for {
		i := bytes.Index(data[offset:], []byte{'P', 'K', 3, 4})
		if i < 0 {
			break
		}
		start := offset + i
		for _, end := range zipPayloadEndOffsets(data[start:]) {
			if end <= 0 || start+end > len(data) {
				continue
			}
			payload := data[start : start+end]
			if isOOXMLPackage(payload) {
				ranges = append(ranges, byteRange{start: embeddedOOXMLPackageMetadataStart(data, start), end: start + end})
				break
			}
		}
		offset = start + zipMagicLen
	}
	return compactByteRanges(ranges)
}

func embeddedOOXMLPackageMetadataStart(data []byte, zipStart int) int {
	const maxMetadataPrefixBytes = 2048
	limit := zipStart - maxMetadataPrefixBytes
	if limit < 0 {
		limit = 0
	}
	pos := zipStart
	for pos > limit {
		tokenEnd := pos
		for tokenEnd > limit && data[tokenEnd-1] == 0 {
			tokenEnd--
		}
		if tokenEnd <= limit {
			break
		}
		tokenStart := tokenEnd
		for tokenStart > limit && data[tokenStart-1] != 0 {
			tokenStart--
		}
		token := strings.TrimSpace(string(data[tokenStart:tokenEnd]))
		if !looksLikeEmbeddedOOXMLPackageMetadataToken(token) {
			break
		}
		pos = tokenStart
		for pos > limit && data[pos-1] == 0 {
			pos--
		}
	}
	return pos
}

func looksLikeEmbeddedOOXMLPackageMetadataToken(s string) bool {
	s = strings.TrimSpace(strings.ReplaceAll(s, "\\", "/"))
	if s == "" {
		return false
	}
	lower := strings.ToLower(s)
	switch path.Base(lower) {
	case "ole10native", "\x01ole10native", "package", "contents", "contents.docx", "contents.xlsx", "contents.pptx":
		return true
	}
	if looksLikeLocalFileURIReference(lower) || hiddenPackageURIPathCandidate(s) != "" {
		return true
	}
	if len(s) >= 3 && ((s[0] >= 'A' && s[0] <= 'Z') || (s[0] >= 'a' && s[0] <= 'z')) && s[1] == ':' && s[2] == '/' {
		return true
	}
	switch strings.ToLower(path.Ext(lower)) {
	case ".docx", ".docm", ".dotx", ".dotm", ".xlsx", ".xlsm", ".xltx", ".xltm", ".pptx", ".pptm", ".ppsx", ".ppsm", ".potx", ".potm":
		return true
	default:
		return false
	}
}

func maskEmbeddedPDFsForText(data []byte) []byte {
	ranges := pdfByteRanges(data)
	if len(ranges) == 0 {
		return data
	}
	out := append([]byte(nil), data...)
	for _, r := range ranges {
		for i := r.start; i < r.end; i++ {
			out[i] = 0
		}
	}
	return out
}

func pdfByteRanges(data []byte) []byteRange {
	var ranges []byteRange
	offset := 0
	for {
		i := bytes.Index(data[offset:], []byte("%PDF-"))
		if i < 0 {
			break
		}
		start := offset + i
		end := len(data)
		if j := bytes.Index(data[start:], []byte("%%EOF")); j >= 0 {
			end = start + j + len("%%EOF")
		}
		if end > start {
			ranges = append(ranges, byteRange{start: start, end: end})
		}
		if end <= start+len("%PDF-") || end >= len(data) {
			break
		}
		offset = end
	}
	return ranges
}

func maskEmbeddedImagesForText(data []byte) []byte {
	ranges := imageByteRanges(data)
	if len(ranges) == 0 {
		return data
	}
	out := append([]byte(nil), data...)
	for _, r := range ranges {
		for i := r.start; i < r.end; i++ {
			out[i] = 0
		}
	}
	return out
}

type byteRange struct {
	start int
	end   int
}

func imageByteRanges(data []byte) []byteRange {
	type sig struct {
		ext   string
		start []byte
	}
	sigs := []sig{
		{".png", []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1a, '\n'}},
		{".jpg", []byte{0xff, 0xd8, 0xff}},
		{".gif", []byte("GIF8")},
		{".bmp", []byte("BM")},
		{".j2k", []byte{0xff, 0x4f, 0xff, 0x51}},
		{".tif", []byte{'I', 'I', 42, 0}},
		{".tif", []byte{'M', 'M', 0, 42}},
		{".tif", []byte{'I', 'I', 43, 0, 8, 0, 0, 0}},
		{".tif", []byte{'M', 'M', 0, 43, 0, 8, 0, 0}},
		{".jxr", []byte{'I', 'I', 0xbc, 0x01}},
		{".webp", []byte("RIFF")},
		{".ico", []byte{0, 0, 1, 0}},
		{".cur", []byte{0, 0, 2, 0}},
		{".pcx", []byte{0x0a}},
	}
	var ranges []byteRange
	for _, s := range sigs {
		offset := 0
		for {
			i := bytes.Index(data[offset:], s.start)
			if i < 0 {
				break
			}
			start := offset + i
			size, ok := imageEndOffset(s.ext, data[start:])
			if !ok {
				offset = start + len(s.start)
				continue
			}
			end := start + size
			if end > start && len(data[start:end]) > 32 && validImageData(s.ext, data[start:end]) {
				ranges = append(ranges, byteRange{start: start, end: end})
			}
			offset = end
		}
	}
	ranges = append(ranges, isoBMFFByteRanges(data)...)
	ranges = append(ranges, jpeg2000ByteRanges(data)...)
	ranges = append(ranges, pictByteRanges(data)...)
	ranges = append(ranges, svgByteRanges(data)...)
	ranges = append(ranges, epsByteRanges(data)...)
	ranges = append(ranges, sizedImageByteRanges(data)...)
	return compactByteRanges(ranges)
}

func compactByteRanges(ranges []byteRange) []byteRange {
	if len(ranges) < 2 {
		return ranges
	}
	sort.Slice(ranges, func(i, j int) bool {
		if ranges[i].start != ranges[j].start {
			return ranges[i].start < ranges[j].start
		}
		return ranges[i].end > ranges[j].end
	})
	kept := ranges[:0]
	for _, r := range ranges {
		if byteRangeContained(r, kept) {
			continue
		}
		kept = append(kept, r)
	}
	return kept
}

func byteRangeContained(r byteRange, kept []byteRange) bool {
	for _, k := range kept {
		if r.start >= k.start && r.end <= k.end {
			return true
		}
	}
	return false
}

func sizedImageByteRanges(data []byte) []byteRange {
	var ranges []byteRange
	for offset := 0; offset+40 <= len(data); offset += 4 {
		size, _, ok := dibDeclaredSize(data[offset:])
		if !ok || offset+size > len(data) {
			continue
		}
		ranges = append(ranges, byteRange{start: offset, end: offset + size})
		offset += size - 4
	}
	for offset := 0; offset+88 <= len(data); offset += 4 {
		if binary.LittleEndian.Uint32(data[offset:]) != 1 {
			continue
		}
		size, ok := emfDeclaredSize(data[offset:])
		if !ok || offset+size > len(data) {
			continue
		}
		ranges = append(ranges, byteRange{start: offset, end: offset + size})
		offset += size - 4
	}
	for offset := 0; offset+18 <= len(data); offset++ {
		size, ok := wmfDeclaredSize(data[offset:])
		if !ok || offset+size > len(data) {
			continue
		}
		ranges = append(ranges, byteRange{start: offset, end: offset + size})
		offset += size - 1
	}
	return ranges
}

func looksLikeTextFragment(s string) bool {
	if s == "" || strings.ContainsRune(s, utf8.RuneError) {
		return false
	}
	var letters, digits, spaces, punctuation, symbols, letterlikeSymbols, total int
	for _, r := range s {
		if r == '\n' || r == '\t' || r == ' ' {
			spaces++
		}
		switch {
		case unicode.IsLetter(r):
			letters++
		case unicode.IsDigit(r):
			digits++
		case isLetterlikeSymbolRune(r):
			letterlikeSymbols++
		case unicode.IsPunct(r):
			punctuation++
		case unicode.IsSymbol(r):
			symbols++
		case unicode.IsSpace(r):
			spaces++
		case !unicode.IsPrint(r):
			return false
		}
		total++
	}
	textual := letters + digits + letterlikeSymbols
	if total == 0 || textual == 0 {
		return false
	}
	if total >= 12 && spaces == 0 && letters+letterlikeSymbols == 0 && punctuation+symbols > digits {
		return false
	}
	if punctuation+symbols > 0 && float64(punctuation+symbols)/float64(total) > 0.55 {
		return false
	}
	return true
}

func looksLikeBinaryControlFragment(s string) bool {
	trimmed := strings.TrimSpace(s)
	nonASCII := containsNonASCIIByte(trimmed)
	switch {
	case strings.EqualFold(trimmed, "powerpoint document"),
		strings.EqualFold(trimmed, "pictures"),
		strings.EqualFold(trimmed, "worddocument"),
		strings.EqualFold(trimmed, "1table"),
		strings.EqualFold(trimmed, "0table"),
		strings.EqualFold(trimmed, "root entry"),
		strings.EqualFold(trimmed, "microsoft office powerpoint"),
		strings.EqualFold(trimmed, "microsoft excel"),
		strings.EqualFold(trimmed, "microsoft excel 2003 worksheet"):
		return true
	}
	if looksLikeOLEClassFragment(trimmed) {
		return true
	}
	if looksLikeOLEIdentifierFragment(trimmed) {
		return true
	}
	if looksLikeOLEWrapperStreamName(trimmed) {
		return true
	}
	if looksLikeOOXMLMarkupNameFragment(trimmed) {
		return true
	}
	if containsASCIIFold(trimmed, "#ppt_") || containsASCIIFold(trimmed, "powerpoint.slide.") {
		return true
	}
	if containsASCIIFold(trimmed, "encryptedpackage") || containsASCIIFold(trimmed, "encryptioninfo") ||
		containsASCIIFold(trimmed, "microsoft.container.dataspaces") || containsASCIIFold(trimmed, "strongencryption") ||
		containsASCIIFold(trimmed, "encryptiontransform") || containsASCIIFold(trimmed, "microsoft enhanced rsa and aes cryptographic provider") {
		return true
	}
	if hasPrefixFold(trimmed, "___ppt") {
		return true
	}
	if hasPrefixFold(trimmed, "%pdf-") || strings.EqualFold(trimmed, "stream") ||
		strings.EqualFold(trimmed, "endstream") || strings.EqualFold(trimmed, "endobj") ||
		hasSuffixFold(trimmed, " obj") {
		return true
	}
	if strings.EqualFold(trimmed, "acrobat document") || hasPrefixFold(trimmed, "acroexch.document") {
		return true
	}
	if strings.HasPrefix(trimmed, "<") && strings.Contains(trimmed, ">") &&
		(strings.Contains(trimmed, "</") || strings.Contains(trimmed, "/>")) {
		return true
	}
	if looksLikeFontTableFragment(trimmed) {
		return true
	}
	if nonASCII {
		if looksLikeLegacyMojibakeControlLine(trimmed) {
			return true
		}
		if looksLikeUnicodeBinaryNoise(trimmed) {
			return true
		}
	}
	if containsLiteralOOXMLTextEscape(trimmed) {
		return false
	}
	if len(trimmed) < 6 || strings.ContainsAny(trimmed, " \t") || nonASCII {
		return false
	}
	if looksLikeLowInformationFragment(trimmed) {
		return true
	}
	return false
}

func looksLikeDiscardableBinaryControlLine(s string) bool {
	trimmed := strings.TrimSpace(s)
	if markdownLikelyTableRow(trimmed) {
		return false
	}
	if markdownCouldStartListMarker(trimmed) {
		if unlisted := strings.TrimSpace(stripMarkdownListMarker(trimmed)); unlisted != "" && unlisted != trimmed {
			return looksLikeBinaryControlFragment(unlisted)
		}
	}
	return looksLikeBinaryControlFragment(trimmed)
}

func markdownCouldStartListMarker(line string) bool {
	if line == "" {
		return false
	}
	switch {
	case line[0] == '-', line[0] == '*', line[0] == '+':
		return true
	case line[0] >= '0' && line[0] <= '9':
		return true
	default:
		return false
	}
}

func containsLiteralOOXMLTextEscape(s string) bool {
	for i := 0; i+7 <= len(s); i++ {
		if s[i] != '_' || (s[i+1] != 'x' && s[i+1] != 'X') || s[i+6] != '_' {
			continue
		}
		if _, ok := parseOOXMLHex4(s[i+2 : i+6]); ok {
			return true
		}
	}
	return false
}

func containsNonASCIIByte(s string) bool {
	for i := 0; i < len(s); i++ {
		if s[i] >= utf8.RuneSelf {
			return true
		}
	}
	return false
}

func looksLikeLegacyMojibakeControlLine(s string) bool {
	if s == "" {
		return false
	}
	noHorizontalSpace := !strings.ContainsAny(s, " \t")
	if looksLikeLegacyRepeatedCyrillicFill(s) {
		return true
	}
	if looksLikeLegacyWordCyrillicControlRun(s) {
		return true
	}
	if looksLikeLegacyPPT95CJKGlyphNoiseLine(s) {
		return true
	}
	if strings.Count(s, "\u745c?") >= 2 && noHorizontalSpace {
		return true
	}
	if strings.Count(s, "я") >= 2 && !strings.ContainsAny(s, " \t") && strings.ContainsAny(s, "0123456789;[$({\\") {
		return true
	}
	if strings.Contains(s, "аяяA0C0E0") || strings.Contains(s, "000°0 2 3 !") {
		return true
	}
	if looksLikeCyrillicEncodingTableNoise(s) {
		return true
	}
	if noHorizontalSpace && stringHasAtLeastRunes(s, 24) {
		if looksLikeLegacyPunctuationEncodingTable(s) {
			return true
		}
		if looksLikeMojibakePunctuationTable(s) {
			return true
		}
	}
	return false
}

func stringHasAtLeastRunes(s string, min int) bool {
	if min <= 0 {
		return true
	}
	count := 0
	for range s {
		count++
		if count >= min {
			return true
		}
	}
	return false
}

func looksLikeLegacyRepeatedCyrillicFill(s string) bool {
	s = strings.TrimSpace(s)
	if len(s) < 6 || strings.ContainsAny(s, " \t") {
		return false
	}
	var unique [4]rune
	uniqueCount := 0
	total := 0
	for _, r := range s {
		if !unicode.Is(unicode.Cyrillic, r) {
			return false
		}
		total++
		seen := false
		for i := 0; i < uniqueCount; i++ {
			if unique[i] == r {
				seen = true
				break
			}
		}
		if !seen {
			if uniqueCount == len(unique) {
				return false
			}
			unique[uniqueCount] = r
			uniqueCount++
		}
	}
	return total >= 6 && uniqueCount <= 3
}

func looksLikeLegacyWordCyrillicControlRun(s string) bool {
	return strings.Contains(s, "胁袏") || strings.Contains(s, "褕屑邪褕") ||
		(strings.Contains(s, "胁") && strings.Contains(s, "?")) ||
		(strings.Contains(s, "Ц") && strings.Contains(s, "ѕ") && strings.Contains(s, "в"))
}

func looksLikeLegacyPPT95CJKGlyphNoiseLine(s string) bool {
	if !strings.ContainsRune(s, '\u8989') && !strings.ContainsRune(s, '\u9689') {
		return false
	}
	for _, r := range s {
		if r <= unicode.MaxASCII && unicode.IsLetter(r) {
			return false
		}
	}
	return true
}

func looksLikeCyrillicEncodingTableNoise(s string) bool {
	var asciiLetters, cyrillic, digits, spaces, marks, symbols int
	for _, r := range s {
		switch {
		case r <= unicode.MaxASCII && unicode.IsLetter(r):
			asciiLetters++
		case unicode.Is(unicode.Cyrillic, r):
			cyrillic++
		case unicode.IsDigit(r):
			digits++
		case unicode.IsSpace(r):
			spaces++
		case unicode.IsPunct(r):
			marks++
		case unicode.IsSymbol(r):
			symbols++
		}
	}
	if cyrillic >= 5 && digits >= 5 && asciiLetters <= 3 && marks+symbols >= 3 {
		return true
	}
	if cyrillic >= 8 && digits >= 3 && asciiLetters <= 3 && spaces <= 2 && marks+symbols >= 2 {
		return true
	}
	if cyrillic >= 5 && digits >= 3 && asciiLetters <= 2 && spaces <= 1 && marks+symbols >= 3 {
		return true
	}
	return false
}

func looksLikeLegacyPunctuationEncodingTable(s string) bool {
	hits := 0
	for _, marker := range []string{"銆併€傦紝", "锛庛兓锛氾紱锛燂紒", "銇併亙銇呫亣", "銈°偅銈ャ偋", "锝★剑锝わ渐"} {
		if strings.Contains(s, marker) {
			hits++
		}
	}
	return hits >= 2
}

func looksLikeOLEClassFragment(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	switch s[0] {
	case 'a', 'A':
		if strings.EqualFold(s, "adobe photoshop image") ||
			strings.EqualFold(s, "adobe acrobat document") ||
			strings.EqualFold(s, "acrobat document") ||
			strings.EqualFold(s, "acroexch.document") ||
			hasPrefixFold(s, "acroexch.document.") ||
			hasPrefixFold(s, "acrobat.document.") {
			return true
		}
	case 'b', 'B':
		if strings.EqualFold(s, "bitmap image") {
			return true
		}
	case 'c', 'C':
		if strings.EqualFold(s, "current user") ||
			strings.EqualFold(s, "cachelastmodifiedfactor.1") ||
			strings.EqualFold(s, "coreldraw") ||
			strings.EqualFold(s, "coreldraw 10.0 graphic") ||
			hasPrefixFold(s, "coreldraw.graphic.") {
			return true
		}
	case 'e', 'E':
		if strings.EqualFold(s, "equation.3") ||
			hasPrefixFold(s, "equation.") ||
			hasPrefixFold(s, "excel.sheet.") ||
			hasPrefixFold(s, "excel.chart.") {
			return true
		}
	case 'f', 'F':
		if hasPrefixFold(s, "forms.") {
			return true
		}
	case 'h', 'H':
		if strings.EqualFold(s, "html document") ||
			strings.EqualFold(s, "htmldocument") ||
			strings.EqualFold(s, "htmlfile") {
			return true
		}
	case 'i', 'I':
		if strings.EqualFold(s, "internet explorer_server") {
			return true
		}
	case 'm', 'M':
		if strings.EqualFold(s, "microsoft equation 3.0") ||
			strings.EqualFold(s, "mathtype equation") ||
			strings.EqualFold(s, "microsoft equation") ||
			strings.EqualFold(s, "microsoft office excel worksheet") ||
			strings.EqualFold(s, "microsoft office excel 2007 worksheet") ||
			strings.EqualFold(s, "microsoft office excel 2007 workbook") ||
			strings.EqualFold(s, "microsoft word document") ||
			strings.EqualFold(s, "microsoft word 97-2003 document") ||
			strings.EqualFold(s, "microsoft word 2007 document") ||
			strings.EqualFold(s, "microsoft office word document") ||
			strings.EqualFold(s, "microsoft office word 97-2003 document") ||
			strings.EqualFold(s, "microsoft office word 2007 document") ||
			strings.EqualFold(s, "microsoft excel worksheet") ||
			strings.EqualFold(s, "microsoft excel 97-2003 worksheet") ||
			strings.EqualFold(s, "microsoft excel 2007 worksheet") ||
			strings.EqualFold(s, "microsoft excel 2007 workbook") ||
			strings.EqualFold(s, "microsoft powerpoint presentation") ||
			strings.EqualFold(s, "microsoft powerpoint 97-2003 presentation") ||
			strings.EqualFold(s, "microsoft powerpoint 2007 presentation") ||
			strings.EqualFold(s, "microsoft office powerpoint presentation") ||
			strings.EqualFold(s, "microsoft office powerpoint 97-2003 presentation") ||
			strings.EqualFold(s, "microsoft office powerpoint 2007 presentation") ||
			strings.EqualFold(s, "microsoft office excel 97-2003 worksheet") ||
			strings.EqualFold(s, "microsoft excel chart") ||
			strings.EqualFold(s, "microsoft graph chart") ||
			strings.EqualFold(s, "microsoft graph 97 chart") ||
			strings.EqualFold(s, "microsoft graph 2000 chart") ||
			strings.EqualFold(s, "microsoft powerpoint slide") ||
			strings.EqualFold(s, "ms org chart") ||
			strings.EqualFold(s, "ms organization chart 2.0") ||
			strings.EqualFold(s, "microsoft photo editor 3.0 photo") ||
			strings.EqualFold(s, "microsoft visio drawing") ||
			strings.EqualFold(s, "media clip") ||
			strings.EqualFold(s, "windows media player") ||
			strings.EqualFold(s, "macromedia flash factory object") ||
			(hasPrefixFold(s, "mathtype ") && hasSuffixFold(s, " equation")) ||
			strings.EqualFold(s, "microsoft forms 2.0 checkbox") ||
			strings.EqualFold(s, "microsoft forms 2.0 combobox") ||
			strings.EqualFold(s, "microsoft forms 2.0 commandbutton") ||
			strings.EqualFold(s, "microsoft forms 2.0 frame") ||
			strings.EqualFold(s, "microsoft forms 2.0 image") ||
			strings.EqualFold(s, "microsoft forms 2.0 label") ||
			strings.EqualFold(s, "microsoft forms 2.0 listbox") ||
			strings.EqualFold(s, "microsoft forms 2.0 multipage") ||
			strings.EqualFold(s, "microsoft forms 2.0 optionbutton") ||
			strings.EqualFold(s, "microsoft forms 2.0 scrollbar") ||
			strings.EqualFold(s, "microsoft forms 2.0 spinbutton") ||
			strings.EqualFold(s, "microsoft forms 2.0 tabstrip") ||
			strings.EqualFold(s, "microsoft forms 2.0 textbox") ||
			strings.EqualFold(s, "microsoft forms 2.0 togglebutton") ||
			hasPrefixFold(s, "ms_clipart_gallery.") ||
			hasPrefixFold(s, "msphotoed.") ||
			hasPrefixFold(s, "mediaplayer.mediaplayer.") ||
			hasPrefixFold(s, "mscomctl.") ||
			hasPrefixFold(s, "mscomctllib.") ||
			hasPrefixFold(s, "mscomct2.") ||
			hasPrefixFold(s, "mscomctl2.") ||
			hasPrefixFold(s, "msforms.") ||
			hasPrefixFold(s, "msgraph.chart.") {
			return true
		}
	case 'o', 'O':
		if strings.EqualFold(s, "outlook.fileattach") ||
			strings.EqualFold(s, "outlook.message") ||
			hasPrefixFold(s, "orgpluswopx.") ||
			hasPrefixFold(s, "outlook.fileattach.") ||
			hasPrefixFold(s, "outlook.message.") {
			return true
		}
	case 'p', 'P':
		if strings.EqualFold(s, "photo editor photo") ||
			strings.EqualFold(s, "package") ||
			strings.EqualFold(s, "package object") ||
			strings.EqualFold(s, "packager shell object") ||
			strings.EqualFold(s, "pdf document") ||
			strings.EqualFold(s, "paint.picture") ||
			strings.EqualFold(s, "powerpoint presentation") ||
			strings.EqualFold(s, "powerpoint slide") ||
			hasPrefixFold(s, "powerpoint.show.") ||
			hasPrefixFold(s, "powerpoint.slide.") ||
			hasPrefixFold(s, "powerpoint.presentation.") ||
			hasPrefixFold(s, "powerpoint.template.") ||
			hasPrefixFold(s, "pdf.document.") ||
			hasPrefixFold(s, "photoshop.image.") {
			return true
		}
	case 'r', 'R':
		if strings.EqualFold(s, "richedit document") ||
			hasPrefixFold(s, "richedit.document.") {
			return true
		}
	case 's', 'S':
		if strings.EqualFold(s, "smartdraw") ||
			strings.EqualFold(s, "smartdraw drawing") ||
			strings.EqualFold(s, "shockwave flash object") ||
			strings.EqualFold(s, "shell explorer") ||
			strings.EqualFold(s, "shell.explorer") ||
			hasPrefixFold(s, "smartdraw.") ||
			hasPrefixFold(s, "shockwaveflash.shockwaveflash.") ||
			hasPrefixFold(s, "shell.explorer.") {
			return true
		}
	case 'v', 'V':
		if hasPrefixFold(s, "visio.drawing.") {
			return true
		}
	case 'w', 'W':
		if strings.EqualFold(s, "wordpad document") ||
			strings.EqualFold(s, "windows media player") ||
			hasPrefixFold(s, "wordpad.document.") ||
			hasPrefixFold(s, "wmplayer.ocx.") ||
			hasPrefixFold(s, "word.document.") {
			return true
		}
	}
	if isLegacyObjectReference(s) {
		return true
	}
	return false
}

func looksLikeOLEIdentifierFragment(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	if looksLikeGUIDString(s) {
		return true
	}
	if value, ok := oleIdentifierAssignmentValue(s); ok {
		return looksLikeGUIDString(value)
	}
	return false
}

func looksLikeOLEWrapperStreamName(s string) bool {
	s = strings.TrimSpace(s)
	s = strings.TrimLeftFunc(s, func(r rune) bool { return r < 0x20 })
	switch {
	case strings.EqualFold(s, "compobj"),
		strings.EqualFold(s, "objinfo"),
		strings.EqualFold(s, "ole10native"):
		return true
	case hasPrefixFold(s, "olepres"):
		for _, r := range s[len("olepres"):] {
			if r < '0' || r > '9' {
				return false
			}
		}
		return len(s) > len("olepres")
	}
	return false
}

func looksLikeOOXMLMarkupNameFragment(s string) bool {
	s = strings.Trim(strings.TrimSpace(s), `"'<>/`)
	switch {
	case strings.EqualFold(s, "relationships"),
		strings.EqualFold(s, "contenttype"),
		strings.EqualFold(s, "partname"),
		strings.EqualFold(s, "targetmode"):
		return true
	default:
		return false
	}
}

func oleIdentifierAssignmentValue(s string) (string, bool) {
	i := strings.IndexAny(s, "=:")
	if i <= 0 {
		return "", false
	}
	key := strings.ToLower(strings.TrimSpace(s[:i]))
	switch key {
	case "clsid", "classid", "guid", "appid":
	default:
		return "", false
	}
	value := strings.TrimSpace(s[i+1:])
	value = strings.Trim(value, `"'`)
	return value, value != ""
}

func looksLikeGUIDString(s string) bool {
	s = strings.TrimSpace(s)
	s = strings.Trim(s, `"'`)
	if len(s) >= 2 && s[0] == '{' && s[len(s)-1] == '}' {
		s = strings.TrimSpace(s[1 : len(s)-1])
	}
	if len(s) != 36 {
		return false
	}
	for i, c := range s {
		switch i {
		case 8, 13, 18, 23:
			if c != '-' {
				return false
			}
		default:
			if !isASCIIHexRune(c) {
				return false
			}
		}
	}
	return true
}

func isASCIIHexRune(r rune) bool {
	return (r >= '0' && r <= '9') || (r >= 'a' && r <= 'f') || (r >= 'A' && r <= 'F')
}

func isLegacyObjectReference(s string) bool {
	bang := strings.IndexByte(s, '!')
	if bang <= 0 || bang+len("!Object 1") > len(s) {
		return false
	}
	for i := 0; i < bang; i++ {
		c := s[i]
		if c >= 'a' && c <= 'z' {
			c -= 'a' - 'A'
		}
		if (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '_' || c == ' ' {
			continue
		}
		return false
	}
	i := bang + 1
	for _, c := range "Object" {
		if i >= len(s) {
			return false
		}
		got := s[i]
		if got >= 'a' && got <= 'z' {
			got -= 'a' - 'A'
		}
		want := byte(c)
		if want >= 'a' && want <= 'z' {
			want -= 'a' - 'A'
		}
		if got != want {
			return false
		}
		i++
	}
	if i >= len(s) || s[i] != ' ' {
		return false
	}
	i++
	if i >= len(s) {
		return false
	}
	for ; i < len(s); i++ {
		if s[i] < '0' || s[i] > '9' {
			return false
		}
	}
	return true
}

func looksLikeFontTableFragment(s string) bool {
	idx := indexASCIIFold(s, "calibri")
	if idx < 0 {
		return false
	}
	if idx == 0 {
		return false
	}
	var nonASCII, letters int
	for _, r := range s[:idx] {
		if unicode.IsLetter(r) {
			letters++
		}
		if r > unicode.MaxASCII {
			nonASCII++
		}
	}
	return nonASCII > 0 && letters == nonASCII
}

func looksLikeUnicodeBinaryNoise(s string) bool {
	trimmed := strings.TrimSpace(s)
	if trimmed == "" {
		return false
	}
	s = trimmed
	runes := []rune(trimmed)
	var asciiLetters, digits, spaces, privateUse, cjk, hangul, kana, cyrillic, latin, otherLetters, symbols, marks, combiningMarks, letterlikeSymbols, rareCJK, suspicious int
	for _, r := range runes {
		switch {
		case r <= unicode.MaxASCII && unicode.IsLetter(r):
			asciiLetters++
		case unicode.IsDigit(r):
			digits++
		case unicode.IsSpace(r):
			spaces++
		case isPrivateUseRune(r):
			privateUse++
			suspicious++
		case unicode.Is(unicode.Han, r):
			cjk++
			if isRareMojibakeCJKRune(r) {
				rareCJK++
				suspicious++
			}
		case unicode.Is(unicode.Hangul, r):
			hangul++
		case unicode.In(r, unicode.Hiragana, unicode.Katakana):
			kana++
		case unicode.Is(unicode.Cyrillic, r):
			cyrillic++
		case isLetterlikeSymbolRune(r):
			letterlikeSymbols++
		case unicode.Is(unicode.Latin, r):
			latin++
		case unicode.IsLetter(r):
			otherLetters++
		case unicode.IsSymbol(r):
			symbols++
			if r > unicode.MaxASCII {
				suspicious++
			}
		case unicode.IsMark(r):
			combiningMarks++
			if r > unicode.MaxASCII {
				suspicious++
			}
		case unicode.IsPunct(r):
			marks++
		}
		if r == '\u7023' || r == '\ufffd' {
			suspicious++
		}
	}
	knownLetters := asciiLetters + cjk + hangul + kana + cyrillic + latin + otherLetters + letterlikeSymbols
	if strings.HasPrefix(s, "嬀") {
		return true
	}
	if looksLikeLegacyRepeatedCyrillicFill(s) {
		return true
	}
	if looksLikeLegacyWordCyrillicControlRun(s) {
		return true
	}
	if cyrillic > 0 && cjk+hangul+kana+latin+otherLetters+letterlikeSymbols+asciiLetters == 0 && privateUse == 0 && combiningMarks == 0 {
		return false
	}
	if len(runes) <= 4 && (runes[0] == 'À' || runes[0] == 'à') && strings.ContainsRune(s, '耀') {
		return true
	}
	if len(runes) >= 2 && len(runes) <= 8 && suspicious > 0 && asciiLetters+digits+kana == 0 && latin <= 1 && cyrillic <= 1 {
		return true
	}
	if looksLikeLegacyPP40ShortUnicodeNoise(runes, spaces) {
		return true
	}
	if looksLikeByteAlignedUnicodeNoise(runes, spaces, asciiLetters, digits, kana, hangul) {
		return true
	}
	if looksLikeSpacedByteAlignedUnicodeNoise(runes, spaces, asciiLetters, digits, kana, hangul) {
		return true
	}
	if looksLikeShortMixedScriptUnicodeNoise(runes, spaces, asciiLetters, digits, kana, hangul, cjk, cyrillic, latin, otherLetters, marks+symbols) {
		return true
	}
	if looksLikeMisreadASCIIWideRunes(runes, spaces, asciiLetters, digits, kana, hangul) {
		return true
	}
	if looksLikeLegacyPPT95UnicodeNoise(runes, spaces, asciiLetters, digits, kana, hangul, cjk, latin, marks+symbols, combiningMarks) {
		return true
	}
	if looksLikeLegacyCJKByteSoup(runes, spaces, asciiLetters, digits, kana, hangul) {
		return true
	}
	if privateUse >= 2 && privateUse*4 >= len(runes) {
		return true
	}
	if len(runes) >= 8 && spaces == 0 && rareCJK > 0 && kana+hangul+cyrillic+letterlikeSymbols == 0 &&
		(latin+digits > 0 || marks+symbols+privateUse > 0) {
		return true
	}
	if len(runes) >= 12 && spaces == 0 && suspicious > 0 && asciiLetters+digits+cyrillic+latin+letterlikeSymbols == 0 && kana <= 1 {
		return true
	}
	if len(runes) >= 16 && spaces == 0 && suspicious >= 2 && asciiLetters+digits+kana+hangul+cyrillic+letterlikeSymbols == 0 {
		return true
	}
	if len(runes) >= 24 && spaces == 0 && cjk*2 >= len(runes) && suspicious > 0 && knownLetters <= cjk+latin && marks+symbols+privateUse > 0 {
		return true
	}
	if len(runes) >= 24 && spaces == 0 && asciiLetters <= 2 && looksLikeMojibakePunctuationTable(s) {
		return true
	}
	if len(runes) >= 20 && spaces == 0 && cjk*2 >= len(runes) && strings.ContainsRune(s, '\u9225') {
		return true
	}
	if looksLikeRepeatedLegacyCJKGlyphNoise(runes, spaces, asciiLetters, digits, kana, hangul) {
		return true
	}
	scriptGroups := 0
	for _, n := range []int{cjk, hangul, kana, cyrillic, latin, otherLetters} {
		if n > 0 {
			scriptGroups++
		}
	}
	if len(runes) >= 24 && spaces == 0 && scriptGroups >= 3 && marks+symbols+privateUse > 0 {
		return true
	}
	return false
}

func looksLikeLegacyShortGlyphSoupLine(s string) bool {
	s = strings.TrimSpace(s)
	total := 0
	if s == "" {
		return false
	}
	var asciiLetters, digits, spaces, cjk, hangul, kana, cyrillic, latin, otherLetters, marksAndSymbols, glyphSoupMarkers int
	var asciiLowBytes, zeroOrControlLowBytes int
	for _, r := range s {
		total++
		if total > 12 {
			return false
		}
		switch {
		case r <= unicode.MaxASCII && unicode.IsLetter(r):
			asciiLetters++
		case unicode.IsDigit(r):
			digits++
		case unicode.IsSpace(r):
			spaces++
		case unicode.Is(unicode.Han, r):
			cjk++
		case unicode.Is(unicode.Hangul, r):
			hangul++
		case unicode.In(r, unicode.Hiragana, unicode.Katakana):
			kana++
		case unicode.Is(unicode.Cyrillic, r):
			cyrillic++
		case unicode.Is(unicode.Latin, r):
			latin++
		case unicode.IsLetter(r):
			otherLetters++
		case unicode.IsPunct(r) || unicode.IsSymbol(r) || unicode.IsMark(r):
			marksAndSymbols++
		}
		if isLegacyShortGlyphSoupMarker(r) {
			glyphSoupMarkers++
		}
		if r >= 0x0100 {
			low := r & 0x00ff
			if low == 0 || low <= 0x1f || low == 0x80 {
				zeroOrControlLowBytes++
			}
			if (low >= 'A' && low <= 'Z') || (low >= 'a' && low <= 'z') || (low >= '0' && low <= '9') || low == ' ' {
				asciiLowBytes++
			}
		}
	}
	if total < 3 {
		return false
	}
	if cyrillic == 0 && asciiLowBytes >= 3 && asciiLowBytes+zeroOrControlLowBytes >= 4 && asciiLetters <= 2 && spaces <= 3 {
		return true
	}
	if kana+cyrillic+otherLetters > 0 {
		return false
	}
	if spaces > 0 || asciiLetters+digits+latin+marksAndSymbols > 0 {
		return false
	}
	letters := cjk + hangul
	if letters != total {
		return false
	}
	if glyphSoupMarkers == 0 {
		return false
	}
	if cjk >= 3 && hangul == 0 {
		return true
	}
	if hangul >= 3 && cjk == 0 {
		return true
	}
	return cjk >= 2 && hangul >= 2
}

func looksLikeLegacyGlyphSoupContinuationLine(s string) bool {
	s = strings.TrimSpace(s)
	total := 0
	if s == "" {
		return false
	}
	for _, r := range s {
		total++
		if total > 16 {
			return false
		}
		if unicode.IsSpace(r) || unicode.IsDigit(r) || (r <= unicode.MaxASCII && unicode.IsLetter(r)) {
			return false
		}
		if unicode.Is(unicode.Han, r) || unicode.Is(unicode.Hangul, r) {
			continue
		}
		return false
	}
	return total >= 1
}

func isLegacyShortGlyphSoupMarker(r rune) bool {
	switch r {
	case '\u3419', '\u35cc', '\u3cef', '\u9fef', '\u9d6c',
		'\u54c5', '\u78ea', '\u7b59', '\u9bfc', '\u9799', '\u9d6f', '\u8c6b', '\u76d6',
		'\u50dd', '\u502f', '\u5087', '\u5edb', '\u7b2d', '\u924b', '\u8e5d', '\u4c4e',
		'\u88d7', '\u6c3d', '\u93a4', '\u82a6', '\u4898', '\u3dcc', '\u98c4', '\u53b3',
		'\ud23f', '\ubfc4', '\ucbdb', '\ubdd7', '\ubfcc', '\ub1f1', '\uacb1', '\ub51d',
		'\ub4b7', '\ub4b4', '\uc9c3', '\ub208', '\ud2d4', '\ub75b', '\uc16e',
		'\uadd1', '\uc8e4', '\ub123', '\ub1fe', '\ube0f', '\ub7ab':
		return true
	default:
		return false
	}
}

func looksLikeLegacyPP40ShortUnicodeNoise(runes []rune, spaces int) bool {
	if len(runes) < 3 || len(runes) > 4 || spaces > 0 {
		return false
	}
	var latinLetters, byteBoundaryRunes int
	for _, r := range runes {
		if unicode.IsDigit(r) || unicode.In(r, unicode.Hiragana, unicode.Katakana) {
			return false
		}
		if r <= unicode.MaxASCII {
			if !unicode.IsLetter(r) {
				return false
			}
			latinLetters++
			continue
		}
		if unicode.Is(unicode.Latin, r) {
			latinLetters++
		}
		if r >= 0x0100 && (r&0x00ff == 0x00 || r&0x00ff == 0x80) {
			byteBoundaryRunes++
		}
	}
	return latinLetters >= 1 && latinLetters <= 2 && byteBoundaryRunes >= 2
}

func looksLikeByteAlignedUnicodeNoise(runes []rune, spaces, asciiLetters, digits, kana, hangul int) bool {
	if len(runes) < 3 || spaces > 0 || digits > 0 || kana > 0 || asciiLetters > 1 {
		return false
	}
	byteLike := 0
	for _, r := range runes {
		if r < 0x0100 {
			continue
		}
		low := r & 0x00ff
		if low == 0x00 || low == 0x80 || low <= 0x1f {
			byteLike++
		}
	}
	if len(runes) <= 4 {
		return byteLike >= 2
	}
	return byteLike*2 >= len(runes) || (len(runes) <= 12 && byteLike >= 3)
}

func looksLikeSpacedByteAlignedUnicodeNoise(runes []rune, spaces, asciiLetters, digits, kana, hangul int) bool {
	if len(runes) < 5 || spaces == 0 || asciiLetters > 0 || kana > 0 || hangul > 0 {
		return false
	}
	byteLike := 0
	nonSpace := 0
	for _, r := range runes {
		if unicode.IsSpace(r) {
			continue
		}
		nonSpace++
		if r >= 0x0100 {
			low := r & 0x00ff
			if low == 0x00 || low == 0x80 || low <= 0x1f {
				byteLike++
			}
		}
	}
	return nonSpace >= 4 && byteLike*2 >= nonSpace
}

func looksLikeShortMixedScriptUnicodeNoise(runes []rune, spaces, asciiLetters, digits, kana, hangul, cjk, cyrillic, latin, otherLetters, marksAndSymbols int) bool {
	if len(runes) < 3 || len(runes) > 32 || spaces > 0 {
		return false
	}
	if kana > 0 && cjk+hangul+cyrillic+latin+otherLetters == 0 {
		return false
	}
	if len(runes) <= 8 && otherLetters >= 2 && asciiLetters <= 1 && digits == 0 && marksAndSymbols == 0 {
		return true
	}
	if len(runes) <= 8 && otherLetters > 0 && cjk > 0 && asciiLetters <= 2 && digits == 0 {
		return true
	}
	if len(runes) <= 16 && hangul >= 2 && cjk >= 2 && kana == 0 && cyrillic == 0 && asciiLetters <= 1 && digits <= 1 {
		return true
	}
	if len(runes) <= 8 && cjk > 0 && latin > 0 && asciiLetters+digits == 0 {
		for _, r := range runes {
			if unicode.IsMark(r) {
				return true
			}
		}
	}
	if len(runes) <= 4 && spaces == 0 && latin > 0 && hangul > 0 && asciiLetters+digits == 0 {
		return true
	}
	scriptGroups := 0
	for _, n := range []int{cjk, hangul, kana, cyrillic, latin, otherLetters} {
		if n > 0 {
			scriptGroups++
		}
	}
	if scriptGroups < 2 {
		return false
	}
	if len(runes) <= 24 && scriptGroups >= 3 && latin > 0 && cjk+hangul > 0 && asciiLetters+digits <= 1 {
		return true
	}
	if len(runes) <= 8 && otherLetters > 0 && marksAndSymbols > 0 && asciiLetters+digits <= 2 {
		return true
	}
	if len(runes) <= 16 && otherLetters > 0 && asciiLetters+digits <= 1 {
		return true
	}
	lowByteControls := 0
	byteAligned := 0
	for _, r := range runes {
		if r < 0x0100 {
			continue
		}
		low := r & 0x00ff
		if low <= 0x1f || low == 0x80 || low == 0x81 {
			lowByteControls++
		}
		if low == 0 || low == 0x80 || low <= 0x1f {
			byteAligned++
		}
	}
	if lowByteControls > 0 && (scriptGroups >= 3 || asciiLetters+digits+marksAndSymbols > 0) {
		return true
	}
	if byteAligned >= 2 && scriptGroups >= 2 && asciiLetters+digits <= 2 {
		return true
	}
	if scriptGroups >= 3 && asciiLetters+digits+marksAndSymbols > 0 {
		return true
	}
	return false
}

func looksLikeMisreadASCIIWideRunes(runes []rune, spaces, asciiLetters, digits, kana, hangul int) bool {
	if len(runes) < 3 || len(runes) > 12 || spaces > 0 || asciiLetters > 0 || digits > 0 || kana+hangul > 0 {
		return false
	}
	asciiLow, zeroLow, cjk := 0, 0, 0
	for _, r := range runes {
		if unicode.Is(unicode.Han, r) {
			cjk++
		}
		if r < 0x0100 {
			continue
		}
		low := r & 0x00ff
		if low == 0 {
			zeroLow++
		}
		if (low >= 'A' && low <= 'Z') || (low >= 'a' && low <= 'z') || (low >= '0' && low <= '9') {
			asciiLow++
		}
	}
	return cjk == len(runes) && asciiLow >= 1 && zeroLow >= 1 && asciiLow+zeroLow >= len(runes)-1
}

func looksLikeLegacyPPT95UnicodeNoise(runes []rune, spaces, asciiLetters, digits, kana, hangul, cjk, latin, marksAndSymbols, combiningMarks int) bool {
	if len(runes) < 3 || spaces > 1 || asciiLetters > 2 {
		return false
	}
	lowByteControls := 0
	numbers := 0
	longestRun := 0
	currentRun := 0
	var prev rune
	for _, r := range runes {
		if unicode.IsNumber(r) {
			numbers++
		}
		if r == prev {
			currentRun++
		} else {
			currentRun = 1
			prev = r
		}
		if currentRun > longestRun {
			longestRun = currentRun
		}
		if r >= 0x0100 {
			low := r & 0x00ff
			if low <= 0x1f || low == 0x80 || low == 0x81 {
				lowByteControls++
			}
		}
	}
	if len(runes) <= 5 {
		if spaces == 0 && cjk >= 2 && marksAndSymbols+digits+numbers > 0 && lowByteControls > 0 && asciiLetters+kana+hangul == 0 {
			return true
		}
		if spaces == 0 && cjk > 0 && hangul > 0 && asciiLetters+digits == 0 && cjk+hangul < len(runes) {
			return true
		}
		return lowByteControls >= 2 && asciiLetters == 0 && digits == 0
	}
	if lowByteControls >= 4 && lowByteControls*2 >= len(runes) {
		return true
	}
	if longestRun >= 6 && (combiningMarks > 0 || cjk+hangul > 0 || lowByteControls > 0) {
		return true
	}
	if combiningMarks >= 4 && combiningMarks*2 >= len(runes) {
		return true
	}
	if len(runes) >= 6 && len(runes) <= 12 && spaces == 0 && cjk > 0 && lowByteControls >= 2 && asciiLetters+digits == 0 {
		return true
	}
	if len(runes) <= 8 && spaces == 0 && cjk > 0 && hangul > 0 && asciiLetters+digits == 0 && cjk+hangul < len(runes) {
		return true
	}
	if len(runes) >= 8 && spaces == 0 && cjk > 0 && hangul > 0 && marksAndSymbols > 0 && asciiLetters+digits == 0 {
		return true
	}
	if len(runes) >= 12 && spaces == 0 && cjk > 0 && latin > 0 && marksAndSymbols > 0 && asciiLetters+digits == 0 {
		return true
	}
	return false
}

func looksLikeLegacyCJKByteSoup(runes []rune, spaces, asciiLetters, digits, kana, hangul int) bool {
	if len(runes) < 3 || spaces > 1 || asciiLetters > 2 || kana+hangul > 1 {
		return false
	}
	markers, privateUse, euro, byteAligned, cjk, letters := 0, 0, 0, 0, 0, 0
	for _, r := range runes {
		if unicode.IsLetter(r) {
			letters++
		}
		if unicode.Is(unicode.Han, r) {
			cjk++
		}
		if isPrivateUseRune(r) {
			privateUse++
			markers++
		}
		if r == '\u20ac' {
			euro++
			markers++
		}
		if r >= 0x0100 {
			low := r & 0x00ff
			if low == 0x00 || low == 0x80 || low <= 0x0f {
				byteAligned++
				markers++
			}
		}
		if isLegacyCJKByteSoupMarker(r) {
			markers++
		}
	}
	if letters == 0 || cjk*2 < letters {
		return false
	}
	if len(runes) <= 5 {
		return (euro > 0 || privateUse > 0 || byteAligned > 0) && markers >= 2
	}
	if len(runes) <= 12 {
		return markers >= 3 && (markers*3 >= len(runes) || privateUse+euro+byteAligned >= 2)
	}
	if privateUse+euro+byteAligned >= 2 && markers*4 >= len(runes) {
		return true
	}
	return markers >= 4 && markers*3 >= len(runes)
}

func isLegacyCJKByteSoupMarker(r rune) bool {
	switch r {
	case '\u4f33', '\u4f77', '\u4f78', '\u4f79', '\u4f7a', '\u4f7b',
		'\u5c88', '\u5d83', '\u5d84', '\u5d8a', '\u5d9e', '\u5dc9',
		'\u60c6', '\u61c8', '\u6d6f', '\u6d9e', '\u6fe1',
		'\u8100', '\u8165', '\u8167', '\u8169', '\u816b', '\u816d', '\u816f', '\u8171', '\u8173', '\u8175', '\u8177', '\u8179', '\u818f',
		'\u8198', '\u81a9', '\u81ab', '\u81ac', '\u81ad', '\u81a7',
		'\u8274', '\u8886', '\u8887', '\u8889', '\u887c',
		'\u91c4', '\u91d4', '\u91d8', '\u91dd', '\u9287', '\u928f',
		'\u9297', '\u929b', '\u929c', '\u92ef', '\u92f7', '\u92ff',
		'\u934a', '\u93c6', '\u93c7', '\u940a':
		return true
	default:
		return false
	}
}

func looksLikeRepeatedLegacyCJKGlyphNoise(runes []rune, spaces, asciiLetters, digits, kana, hangul int) bool {
	if len(runes) >= 3 && len(runes) <= 64 && spaces == 0 && asciiLetters == 0 && digits == 0 && kana+hangul <= 1 {
		markers := 0
		hasRareMarker := false
		for _, r := range runes {
			if isLegacyCJKGlyphNoiseMarker(r) {
				markers++
				if r == '\u8989' || r == '\u9689' || r == '\u9289' {
					hasRareMarker = true
				}
			}
		}
		return markers >= 3 && markers*2 >= len(runes) && hasRareMarker
	}
	if len(runes) < 32 || spaces > 0 || asciiLetters > 0 || digits > 2 || kana > 0 {
		return false
	}
	markers := 0
	longestMarkerRun := 0
	currentMarkerRun := 0
	for _, r := range runes {
		if isLegacyCJKGlyphNoiseMarker(r) {
			markers++
			currentMarkerRun++
			if currentMarkerRun > longestMarkerRun {
				longestMarkerRun = currentMarkerRun
			}
			continue
		}
		currentMarkerRun = 0
	}
	return markers >= 8 && (markers*8 >= len(runes) || longestMarkerRun >= 4)
}

func isLegacyCJKGlyphNoiseMarker(r rune) bool {
	switch r {
	case '\u0101', '\u0401', '\u0404', '\u0504', '\u8900', '\u8913', '\u8983', '\u8989',
		'\u8996', '\u89a1', '\u89a3', '\u89aa', '\u89b6', '\u9289', '\u9689',
		'\ua100', '\ua189', '\ua300', '\ua389', '\uaa00', '\uaa89':
		return true
	default:
		return false
	}
}

func looksLikeMojibakePunctuationTable(s string) bool {
	hits := 0
	for _, marker := range []string{
		"\u951f\ufffd",
		"\u95bf\ufffd",
		"\u95b3\u30e6\u7462\u951f\ufffd",
		"\ufffd\ufffd",
		"銇併亙銇呫亣",
		"銈°偅銈ャ偋",
	} {
		if strings.Contains(s, marker) {
			hits++
		}
	}
	return hits >= 2
}

func isLetterlikeSymbolRune(r rune) bool {
	return r >= 0x1d400 && r <= 0x1d7ff
}

func isRareMojibakeCJKRune(r rune) bool {
	switch r {
	case '\u704f', '\u845a', '\u808a', '\u8197', '\u887b', '\u8880', '\u8641', '\u8651', '\u91d4',
		'\u9293', '\u92f1', '\u941e', '\u9597', '\u95f7', '\u93c7', '\u93c6', '\u93c9':
		return true
	default:
		return false
	}
}

func isPrivateUseRune(r rune) bool {
	return (r >= 0xe000 && r <= 0xf8ff) ||
		(r >= 0xf0000 && r <= 0xffffd) ||
		(r >= 0x100000 && r <= 0x10fffd)
}

func looksLikeLowInformationFragment(s string) bool {
	s = strings.TrimSpace(s)
	if len(s) < 6 {
		return false
	}
	letters, digits, marks, vowels, total := 0, 0, 0, 0, 0
	longestRun, currentRun := 0, 0
	var prev byte
	var unique [3]byte
	uniqueCount := 0
	for i := 0; i < len(s); i++ {
		b := s[i]
		if b >= utf8.RuneSelf || isASCIISpaceByte(b) {
			return false
		}
		total++
		if b == prev {
			currentRun++
		} else {
			currentRun = 1
			prev = b
		}
		if currentRun > longestRun {
			longestRun = currentRun
		}
		seen := false
		for i := 0; i < uniqueCount; i++ {
			if unique[i] == b {
				seen = true
				break
			}
		}
		if !seen && uniqueCount < len(unique) {
			unique[uniqueCount] = b
			uniqueCount++
		}
		switch {
		case (b >= 'A' && b <= 'Z') || (b >= 'a' && b <= 'z'):
			letters++
			switch b {
			case 'a', 'e', 'i', 'o', 'u', 'A', 'E', 'I', 'O', 'U':
				vowels++
			}
		case b >= '0' && b <= '9':
			digits++
		case b >= 0x21 && b <= 0x7e:
			marks++
		}
	}
	if total < 6 {
		return false
	}
	if longestRun >= 3 {
		return true
	}
	if uniqueCount <= 2 {
		return true
	}
	if vowels == 0 && marks > 0 && marks*3 >= total {
		return true
	}
	if vowels == 0 && digits == 0 && letters >= total*3/4 && total >= 8 {
		return true
	}
	return false
}

func isASCIISpaceByte(b byte) bool {
	switch b {
	case ' ', '\t', '\n', '\r', '\f', '\v':
		return true
	default:
		return false
	}
}

func asciiStrings(data []byte, min int) []string {
	var out []string
	var cur []byte
	flush := func() {
		if len(cur) >= min {
			out = append(out, string(cur))
		}
		cur = cur[:0]
	}
	for _, b := range data {
		if b >= 0x20 && b <= 0x7e || b == '\t' || b == '\n' || b == '\r' {
			cur = append(cur, b)
		} else {
			flush()
		}
	}
	flush()
	return out
}

func singleByteStrings(data []byte, min int, codePage uint16) []string {
	var out []string
	var cur []byte
	flush := func() {
		if len(cur) >= min {
			if codePage == 0 {
				out = append(out, decodeBestLegacySingleByte(cur))
			} else {
				out = append(out, decodeCodePageBytes(cur, codePage))
			}
		}
		cur = cur[:0]
	}
	for _, b := range data {
		if isLegacySingleByteTextByte(b) {
			cur = append(cur, b)
		} else {
			flush()
		}
	}
	flush()
	return out
}

func isLegacySingleByteTextByte(b byte) bool {
	return b == '\t' || b == '\n' || b == '\r' || b >= 0x20
}

func decodeBestLegacySingleByte(raw []byte) string {
	if utf8.Valid(raw) && hasUTF8Multibyte(raw) {
		return string(raw)
	}
	if codePage := bestLegacySingleByteCodePage(raw); codePage != 0 {
		return decodeCodePageBytes(raw, codePage)
	}
	return decodeCodePageBytes(raw, 1252)
}

func hasUTF8Multibyte(raw []byte) bool {
	for len(raw) > 0 {
		r, size := utf8.DecodeRune(raw)
		if r == utf8.RuneError && size == 1 {
			return false
		}
		if size > 1 {
			return true
		}
		raw = raw[size:]
	}
	return false
}

func shouldDecodeWindows1251(raw []byte) bool {
	var printable, cyrillicRange int
	for _, b := range raw {
		if b == '\t' || b == '\n' || b == '\r' || b >= 0x20 {
			printable++
		}
		if b >= 0xc0 || b == 0xa8 || b == 0xb8 {
			cyrillicRange++
		}
	}
	return cyrillicRange >= 4 && cyrillicRange*3 >= printable
}

func bestLegacySingleByteCodePage(raw []byte) uint16 {
	gbkText, gbkOK := decodeLikelyGBKText(raw)
	big5Text, big5OK := decodeLikelyBig5Text(raw)
	shiftJISText, shiftJISOK := decodeLikelyShiftJISText(raw)
	eucKRText, eucKROK := decodeLikelyEUCKRText(raw)
	cyrillicOK := shouldDecodeWindows1251(raw)
	if eucKROK {
		eucKRScore := legacyKoreanCandidateScore(eucKRText)
		gbkScore := -1
		if gbkOK {
			gbkScore = legacyCJKCandidateScore(gbkText, false)
		}
		big5Score := -1
		if big5OK {
			big5Score = legacyCJKCandidateScore(big5Text, false)
		}
		shiftJISScore := -1
		if shiftJISOK {
			shiftJISScore = legacyCJKCandidateScore(shiftJISText, true)
		}
		cyrillicScore := -1
		if cyrillicOK {
			cyrillicScore = countRunesInScript(decodeCodePageBytes(raw, 1251), unicode.Cyrillic)
		}
		if eucKRScore >= gbkScore+2 && eucKRScore >= big5Score+2 && eucKRScore >= shiftJISScore+2 && eucKRScore >= cyrillicScore {
			return 949
		}
	}
	if shiftJISOK {
		if !gbkOK && !big5OK && !eucKROK && !cyrillicOK {
			return 932
		}
		shiftJISScore := legacyCJKCandidateScore(shiftJISText, true)
		gbkScore := -1
		if gbkOK {
			gbkScore = legacyCJKCandidateScore(gbkText, false)
		}
		big5Score := -1
		if big5OK {
			big5Score = legacyCJKCandidateScore(big5Text, false)
		}
		cyrillicScore := -1
		if cyrillicOK {
			cyrillicScore = countRunesInScript(decodeCodePageBytes(raw, 1251), unicode.Cyrillic)
		}
		if shiftJISScore >= gbkScore+2 && shiftJISScore >= big5Score+2 && shiftJISScore >= cyrillicScore {
			return 932
		}
	}
	if gbkOK && big5OK && cyrillicOK {
		cyrillicText := decodeCodePageBytes(raw, 1251)
		cyrillicCount := countRunesInScript(cyrillicText, unicode.Cyrillic)
		gbkCJK := countRunesInScript(gbkText, unicode.Han)
		big5CJK := countRunesInScript(big5Text, unicode.Han)
		if cyrillicCount >= gbkCJK*2 && cyrillicCount >= big5CJK*2 {
			return 1251
		}
	}
	if big5OK && cyrillicOK && !gbkOK {
		big5CJK := countRunesInScript(big5Text, unicode.Han)
		cyrillicText := decodeCodePageBytes(raw, 1251)
		cyrillicCount := countRunesInScript(cyrillicText, unicode.Cyrillic)
		if big5CJK >= cyrillicCount || (big5CJK >= 2 && countSymbolRunes(cyrillicText) > 0 && big5CJK*2 >= cyrillicCount) {
			return 950
		}
		return 1251
	}
	if gbkOK && cyrillicOK && !big5OK {
		gbkCJK := countRunesInScript(gbkText, unicode.Han)
		cyrillicText := decodeCodePageBytes(raw, 1251)
		cyrillicCount := countRunesInScript(cyrillicText, unicode.Cyrillic)
		if gbkCJK >= cyrillicCount || (gbkCJK >= 2 && countSymbolRunes(cyrillicText) > 0 && gbkCJK*2 >= cyrillicCount) {
			return 936
		}
		return 1251
	}
	if gbkOK && big5OK {
		gbkScore := legacyChineseCandidateScore(gbkText, false)
		big5Score := legacyChineseCandidateScore(big5Text, true)
		if big5Score >= gbkScore+2 {
			return 950
		}
		if gbkScore >= big5Score {
			return 936
		}
	}
	if gbkOK {
		return 936
	}
	if big5OK {
		return 950
	}
	if cyrillicOK {
		return 1251
	}
	return 0
}

func legacyKoreanCandidateScore(text string) int {
	hangul := countRunesInScript(text, unicode.Hangul)
	cjk := countRunesInScript(text, unicode.Han)
	return hangul*3 + cjk
}

func legacyCJKCandidateScore(text string, preferKana bool) int {
	cjk := countRunesInScript(text, unicode.Han)
	kana := countRunesInScripts(text, unicode.Hiragana, unicode.Katakana)
	score := cjk + kana
	if preferKana {
		score += kana * 2
	}
	return score
}

func legacyChineseCandidateScore(text string, traditional bool) int {
	score := legacyCJKCandidateScore(text, false)
	if traditional {
		score += countLikelyTraditionalChineseRunes(text) * 2
	} else {
		score += countLikelySimplifiedChineseRunes(text) * 2
	}
	return score
}

func countLikelySimplifiedChineseRunes(text string) int {
	count := 0
	for _, r := range text {
		if strings.ContainsRune("这为来个时见说汉语体测试听读写国学门间龙后发经", r) {
			count++
		}
	}
	return count
}

func countLikelyTraditionalChineseRunes(text string) int {
	count := 0
	for _, r := range text {
		if strings.ContainsRune("這為來個時見說漢語體測試聽讀寫國學門間龍後發經", r) {
			count++
		}
	}
	return count
}

func decodeLikelyGBKText(raw []byte) (string, bool) {
	if gbkDoubleBytePairCount(raw) < 2 {
		return "", false
	}
	text := decodeTextEncodingBytes(raw, simplifiedchinese.GBK)
	if text == "" || strings.ContainsRune(text, utf8.RuneError) || !looksLikeTextFragment(text) {
		return "", false
	}
	var cjk, asciiLetters, cyrillic, kana, hangul, otherLetters int
	for _, r := range text {
		switch {
		case unicode.Is(unicode.Han, r):
			cjk++
		case r <= unicode.MaxASCII && unicode.IsLetter(r):
			asciiLetters++
		case unicode.Is(unicode.Cyrillic, r):
			cyrillic++
		case unicode.In(r, unicode.Hiragana, unicode.Katakana):
			kana++
		case unicode.Is(unicode.Hangul, r):
			hangul++
		case unicode.IsLetter(r):
			otherLetters++
		}
	}
	nonASCIILetters := cjk + cyrillic + kana + hangul + otherLetters
	if cjk < 2 || cjk*2 < nonASCIILetters {
		return "", false
	}
	if asciiLetters == 0 && cyrillic+kana+hangul+otherLetters > cjk/2 {
		return "", false
	}
	return text, true
}

func decodeLikelyBig5Text(raw []byte) (string, bool) {
	if big5DoubleBytePairCount(raw) < 2 {
		return "", false
	}
	text := decodeTextEncodingBytes(raw, traditionalchinese.Big5)
	if text == "" || strings.ContainsRune(text, utf8.RuneError) || !looksLikeTextFragment(text) {
		return "", false
	}
	var cjk, asciiLetters, cyrillic, kana, hangul, otherLetters int
	for _, r := range text {
		switch {
		case unicode.Is(unicode.Han, r):
			cjk++
		case r <= unicode.MaxASCII && unicode.IsLetter(r):
			asciiLetters++
		case unicode.Is(unicode.Cyrillic, r):
			cyrillic++
		case unicode.In(r, unicode.Hiragana, unicode.Katakana):
			kana++
		case unicode.Is(unicode.Hangul, r):
			hangul++
		case unicode.IsLetter(r):
			otherLetters++
		}
	}
	nonASCIILetters := cjk + cyrillic + kana + hangul + otherLetters
	if cjk < 2 || cjk*2 < nonASCIILetters {
		return "", false
	}
	if asciiLetters == 0 && cyrillic+kana+hangul+otherLetters > cjk/2 {
		return "", false
	}
	return text, true
}

func decodeLikelyEUCKRText(raw []byte) (string, bool) {
	if eucKRDoubleBytePairCount(raw) < 2 {
		return "", false
	}
	text := decodeTextEncodingBytes(raw, korean.EUCKR)
	if text == "" || strings.ContainsRune(text, utf8.RuneError) || !looksLikeTextFragment(text) {
		return "", false
	}
	var hangul, hangulSyllables, hangulCompatJamo, cjk, kana, asciiLetters, cyrillic, otherLetters int
	for _, r := range text {
		switch {
		case unicode.Is(unicode.Hangul, r):
			hangul++
			if r >= 0xac00 && r <= 0xd7a3 {
				hangulSyllables++
			} else if r >= 0x3130 && r <= 0x318f {
				hangulCompatJamo++
			}
		case unicode.Is(unicode.Han, r):
			cjk++
		case unicode.In(r, unicode.Hiragana, unicode.Katakana):
			kana++
		case r <= unicode.MaxASCII && unicode.IsLetter(r):
			asciiLetters++
		case unicode.Is(unicode.Cyrillic, r):
			cyrillic++
		case unicode.IsLetter(r):
			otherLetters++
		}
	}
	if hangul < 2 {
		return "", false
	}
	if hangulSyllables < 2 || hangulCompatJamo > hangulSyllables/2 {
		return "", false
	}
	if cjk > 0 {
		return "", false
	}
	if hangul*2 < cjk+kana+otherLetters {
		return "", false
	}
	if asciiLetters == 0 && cjk+kana+cyrillic+otherLetters > hangul {
		return "", false
	}
	return text, true
}

func decodeLikelyShiftJISText(raw []byte) (string, bool) {
	if shiftJISDoubleBytePairCount(raw) < 2 {
		return "", false
	}
	text := decodeTextEncodingBytes(raw, japanese.ShiftJIS)
	if text == "" || strings.ContainsRune(text, utf8.RuneError) || !looksLikeTextFragment(text) {
		return "", false
	}
	var cjk, kana, asciiLetters, cyrillic, hangul, otherLetters int
	for _, r := range text {
		switch {
		case unicode.Is(unicode.Han, r):
			cjk++
		case unicode.In(r, unicode.Hiragana, unicode.Katakana):
			kana++
		case r <= unicode.MaxASCII && unicode.IsLetter(r):
			asciiLetters++
		case unicode.Is(unicode.Cyrillic, r):
			cyrillic++
		case unicode.Is(unicode.Hangul, r):
			hangul++
		case unicode.IsLetter(r):
			otherLetters++
		}
	}
	if kana < 2 && !(kana > 0 && cjk >= 2) {
		return "", false
	}
	if asciiLetters == 0 && cyrillic+hangul+otherLetters > cjk+kana {
		return "", false
	}
	return text, true
}

func gbkDoubleBytePairCount(raw []byte) int {
	count := 0
	for i := 0; i+1 < len(raw); i++ {
		lead, trail := raw[i], raw[i+1]
		if lead >= 0x81 && lead <= 0xfe && trail >= 0x40 && trail <= 0xfe && trail != 0x7f {
			count++
			i++
		}
	}
	return count
}

func eucKRDoubleBytePairCount(raw []byte) int {
	count := 0
	for i := 0; i+1 < len(raw); i++ {
		lead, trail := raw[i], raw[i+1]
		if lead >= 0x81 && lead <= 0xfe && trail >= 0x41 && trail <= 0xfe && trail != 0x7f {
			count++
			i++
		}
	}
	return count
}

func big5DoubleBytePairCount(raw []byte) int {
	count := 0
	for i := 0; i+1 < len(raw); i++ {
		lead, trail := raw[i], raw[i+1]
		if lead >= 0x81 && lead <= 0xfe && ((trail >= 0x40 && trail <= 0x7e) || (trail >= 0xa1 && trail <= 0xfe)) {
			count++
			i++
		}
	}
	return count
}

func shiftJISDoubleBytePairCount(raw []byte) int {
	count := 0
	for i := 0; i+1 < len(raw); i++ {
		lead, trail := raw[i], raw[i+1]
		if ((lead >= 0x81 && lead <= 0x9f) || (lead >= 0xe0 && lead <= 0xfc)) &&
			((trail >= 0x40 && trail <= 0x7e) || (trail >= 0x80 && trail <= 0xfc)) {
			count++
			i++
		}
	}
	return count
}

func countRunesInScript(s string, table *unicode.RangeTable) int {
	count := 0
	for _, r := range s {
		if unicode.Is(table, r) {
			count++
		}
	}
	return count
}

func countRunesInScripts(s string, tables ...*unicode.RangeTable) int {
	count := 0
	for _, r := range s {
		for _, table := range tables {
			if unicode.Is(table, r) {
				count++
				break
			}
		}
	}
	return count
}

func countSymbolRunes(s string) int {
	count := 0
	for _, r := range s {
		if unicode.IsSymbol(r) {
			count++
		}
	}
	return count
}

func decodeCodePageBytes(raw []byte, codePage uint16) string {
	switch codePage {
	case 1200:
		return utf16BytesToString(raw)
	case 932, 943:
		return preferValidUTF8IfCodePageLooksMojibake(raw, decodeTextEncodingBytes(raw, japanese.ShiftJIS))
	case 936:
		return preferValidUTF8IfCodePageLooksMojibake(raw, decodeTextEncodingBytes(raw, simplifiedchinese.GBK))
	case 949:
		return preferValidUTF8IfCodePageLooksMojibake(raw, decodeTextEncodingBytes(raw, korean.EUCKR))
	case 950:
		return preferValidUTF8IfCodePageLooksMojibake(raw, decodeTextEncodingBytes(raw, traditionalchinese.Big5))
	case 1250:
		return decodeTextEncodingBytes(raw, charmap.Windows1250)
	case 1251:
		return decodeTextEncodingBytes(raw, charmap.Windows1251)
	case 1253:
		return decodeTextEncodingBytes(raw, charmap.Windows1253)
	case 1254:
		return decodeTextEncodingBytes(raw, charmap.Windows1254)
	case 1255:
		return decodeTextEncodingBytes(raw, charmap.Windows1255)
	case 1256:
		return decodeTextEncodingBytes(raw, charmap.Windows1256)
	case 1257:
		return decodeTextEncodingBytes(raw, charmap.Windows1257)
	case 1258:
		return decodeTextEncodingBytes(raw, charmap.Windows1258)
	case 10000:
		return decodeTextEncodingBytes(raw, charmap.Macintosh)
	case 54936:
		return decodeTextEncodingBytes(raw, simplifiedchinese.GB18030)
	case 65001:
		if utf8.Valid(raw) {
			return string(raw)
		}
	}
	if codePage != 1252 && utf8.Valid(raw) {
		return string(raw)
	}
	var b strings.Builder
	for _, c := range raw {
		if r, ok := codePageByteRune(c, codePage); ok {
			b.WriteRune(r)
		}
	}
	return b.String()
}

func preferValidUTF8IfCodePageLooksMojibake(raw []byte, decoded string) string {
	if decoded == "" || !utf8.Valid(raw) || !hasUTF8Multibyte(raw) {
		return decoded
	}
	utf8Text := string(raw)
	if mojibakeScore(decoded) >= mojibakeScore(utf8Text)+2 && looksLikeTextFragment(utf8Text) {
		return utf8Text
	}
	return decoded
}

func mojibakeScore(s string) int {
	score := 0
	for _, r := range s {
		if r == utf8.RuneError || isPrivateUseRune(r) {
			score += 2
		}
		if isRareMojibakeCJKRune(r) {
			score++
		}
	}
	for _, marker := range []string{"锟", "娴ｆ粍", "閸掑棝", "濮楃喕", "缁夋槒"} {
		if strings.Contains(s, marker) {
			score += 2
		}
	}
	return score
}

func decodeTextEncodingBytes(raw []byte, enc textencoding.Encoding) string {
	decoded, err := enc.NewDecoder().Bytes(raw)
	if err != nil {
		return ""
	}
	return string(decoded)
}

func codePageByteRune(b byte, codePage uint16) (rune, bool) {
	switch b {
	case '\t', '\n', '\r':
		return rune(b), true
	}
	if b >= 0x20 && b <= 0x7e {
		return rune(b), true
	}
	if codePage == 1251 {
		return windows1251ByteRune(b)
	}
	if codePage != 0 && codePage != 1252 {
		return windows1252ByteRune(b)
	}
	return windows1252ByteRune(b)
}

func windows1251ByteRune(b byte) (rune, bool) {
	if b == 0xa0 {
		return ' ', true
	}
	if b >= 0xc0 {
		return rune(0x0410 + int(b) - 0xc0), true
	}
	table := [...]rune{
		0x0402, 0x0403, 0x201a, 0x0453, 0x201e, 0x2026, 0x2020, 0x2021,
		0x20ac, 0x2030, 0x0409, 0x2039, 0x040a, 0x040c, 0x040b, 0x040f,
		0x0452, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
		0, 0x2122, 0x0459, 0x203a, 0x045a, 0x045c, 0x045b, 0x045f,
		0, 0x040e, 0x045e, 0x0408, 0x00a4, 0x0490, 0x00a6, 0x00a7,
		0x0401, 0x00a9, 0x0404, 0x00ab, 0x00ac, 0, 0x00ae, 0x0407,
		0x00b0, 0x00b1, 0x0406, 0x0456, 0x0491, 0x00b5, 0x00b6, 0x00b7,
		0x0451, 0x2116, 0x0454, 0x00bb, 0x0458, 0x0405, 0x0455, 0x0457,
	}
	if b < 0x80 {
		return 0, false
	}
	r := table[b-0x80]
	return r, r != 0
}

func windows1252ByteRune(b byte) (rune, bool) {
	if b == 0xa0 {
		return ' ', true
	}
	if b >= 0xa1 {
		return rune(b), true
	}
	if b < 0x80 {
		return 0, false
	}
	table := [...]rune{
		0x20ac, 0, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021,
		0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0, 0x017d, 0,
		0, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
		0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0, 0x017e, 0x0178,
	}
	r := table[b-0x80]
	return r, r != 0
}

func utf16Strings(data []byte, min int) []string {
	var out []string
	for start := 0; start < 2; start++ {
		var cur []uint16
		var raw []byte
		zeroCount, asciiPrintableCount := 0, 0
		flush := func() {
			if len(cur) >= min {
				r := utf16.Decode(cur)
				s := string(r)
				if printableRatio(s) > 0.75 && hasUTF16EvidenceCounts(zeroCount, asciiPrintableCount, len(raw)) && !looksLikeMisalignedUTF16(cur) && !looksLikeASCIIBytesMisreadAsUTF16(raw, s) {
					out = append(out, s)
				}
			}
			cur = cur[:0]
			raw = raw[:0]
			zeroCount, asciiPrintableCount = 0, 0
		}
		for i := start; i+1 < len(data); i += 2 {
			v := binary.LittleEndian.Uint16(data[i:])
			if len(cur) >= min && pairedASCIIUTF16Unit(v) && hasUTF16EvidenceCounts(zeroCount, asciiPrintableCount, len(raw)) && nextPairedASCIIUTF16Units(data, i, 2) {
				flush()
			}
			if v == '\t' || v == '\n' || v == '\r' || (v >= 0x20 && v < 0xd800) {
				cur = append(cur, v)
				b0, b1 := data[i], data[i+1]
				raw = append(raw, b0, b1)
				if b0 == 0 {
					zeroCount++
				}
				if b1 == 0 {
					zeroCount++
				}
				if b0 >= 0x20 && b0 <= 0x7e {
					asciiPrintableCount++
				}
				if b1 >= 0x20 && b1 <= 0x7e {
					asciiPrintableCount++
				}
			} else {
				flush()
			}
		}
		flush()
	}
	return out
}

func pairedASCIIUTF16Unit(v uint16) bool {
	lo, hi := byte(v), byte(v>>8)
	return lo >= 0x20 && lo <= 0x7e && hi >= 0x20 && hi <= 0x7e
}

func nextPairedASCIIUTF16Units(data []byte, offset, count int) bool {
	for n := 0; n < count; n++ {
		i := offset + n*2
		if i+1 >= len(data) || !pairedASCIIUTF16Unit(binary.LittleEndian.Uint16(data[i:])) {
			return false
		}
	}
	return true
}

func looksLikeMisalignedUTF16(units []uint16) bool {
	if len(units) < 4 {
		return false
	}
	swappedASCII, pairedASCII, normalASCII, nonZero := 0, 0, 0, 0
	for _, v := range units {
		if v == 0 {
			continue
		}
		nonZero++
		lo, hi := byte(v), byte(v>>8)
		if hi == 0 && lo >= 0x20 && lo <= 0x7e {
			normalASCII++
		}
		if lo == 0 && hi >= 0x20 && hi <= 0x7e {
			swappedASCII++
		}
		if lo >= 0x20 && lo <= 0x7e && hi >= 0x20 && hi <= 0x7e {
			pairedASCII++
		}
	}
	if nonZero == 0 {
		return false
	}
	if float64(swappedASCII)/float64(nonZero) > 0.35 {
		return true
	}
	return normalASCII == 0 && float64(pairedASCII)/float64(nonZero) > 0.55
}

func hasUTF16Evidence(raw []byte) bool {
	if len(raw) == 0 {
		return false
	}
	zeros, asciiPrintable := 0, 0
	for _, b := range raw {
		if b == 0 {
			zeros++
		}
		if b >= 0x20 && b <= 0x7e {
			asciiPrintable++
		}
	}
	if zeros > 0 {
		return true
	}
	return float64(asciiPrintable)/float64(len(raw)) < 0.70
}

func hasUTF16EvidenceCounts(zeros, asciiPrintable, total int) bool {
	if total == 0 {
		return false
	}
	if zeros > 0 {
		return true
	}
	return float64(asciiPrintable)/float64(total) < 0.70
}

func printableRatio(s string) float64 {
	if s == "" {
		return 0
	}
	total, printable := 0, 0
	for _, r := range s {
		total++
		if unicode.IsPrint(r) || unicode.IsSpace(r) {
			printable++
		}
	}
	return float64(printable) / float64(total)
}

func carveImages(data []byte) []Image {
	type sig struct {
		ext   string
		start []byte
	}
	sigs := []sig{
		{".png", []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1a, '\n'}},
		{".jpg", []byte{0xff, 0xd8, 0xff}},
		{".gif", []byte("GIF8")},
		{".j2k", []byte{0xff, 0x4f, 0xff, 0x51}},
		{".tif", []byte{'I', 'I', 42, 0}},
		{".tif", []byte{'M', 'M', 0, 42}},
		{".tif", []byte{'I', 'I', 43, 0, 8, 0, 0, 0}},
		{".tif", []byte{'M', 'M', 0, 43, 0, 8, 0, 0}},
		{".jxr", []byte{'I', 'I', 0xbc, 0x01}},
		{".webp", []byte("RIFF")},
		{".ico", []byte{0, 0, 1, 0}},
		{".cur", []byte{0, 0, 2, 0}},
		{".pcx", []byte{0x0a}},
	}
	var candidates []imageCandidate
	for _, s := range sigs {
		offset := 0
		for {
			i := bytes.Index(data[offset:], s.start)
			if i < 0 {
				break
			}
			start := offset + i
			size, ok := imageEndOffset(s.ext, data[start:])
			if !ok {
				offset = start + len(s.start)
				continue
			}
			end := start + size
			img, ok := normalizeImageData(s.ext, data[start:end])
			if ok && len(img) > 32 {
				candidates = append(candidates, imageCandidate{start: start, end: end, ext: s.ext, data: append([]byte(nil), img...)})
			}
			offset = end
		}
	}
	candidates = append(candidates, carveISOImages(data)...)
	candidates = append(candidates, carveJPEG2000Images(data)...)
	candidates = append(candidates, carvePICTImages(data)...)
	candidates = append(candidates, carveSVGImages(data)...)
	candidates = append(candidates, carveEPSImages(data)...)
	sort.Slice(candidates, func(i, j int) bool {
		if candidates[i].start != candidates[j].start {
			return candidates[i].start < candidates[j].start
		}
		return candidates[i].end > candidates[j].end
	})
	var images []Image
	var kept []imageCandidate
	for _, c := range candidates {
		if imageCandidateContained(c, kept) {
			continue
		}
		kept = append(kept, c)
		images = append(images, Image{
			Name: fmt.Sprintf("legacy-image-%03d%s", len(images)+1, c.ext),
			Ext:  c.ext,
			Data: c.data,
		})
	}
	images = append(images, carveSizedImages(data, len(images), kept)...)
	return images
}

type imageCandidate struct {
	start int
	end   int
	ext   string
	data  []byte
}

func imageCandidateContained(c imageCandidate, kept []imageCandidate) bool {
	for _, k := range kept {
		if c.start >= k.start && c.end <= k.end {
			return true
		}
	}
	return false
}

func carveSizedImages(data []byte, startIndex int, kept []imageCandidate) []Image {
	var images []Image
	for offset := 0; offset+14 <= len(data); offset++ {
		if data[offset] == 'B' && data[offset+1] == 'M' {
			size := int(binary.LittleEndian.Uint32(data[offset+2:]))
			if size >= 14 && imageCandidateContained(imageCandidate{start: offset, end: offset + size}, kept) {
				offset += size - 1
				continue
			}
			if size >= 14 && offset+size <= len(data) && validImageData(".bmp", data[offset:offset+size]) {
				img := append([]byte(nil), data[offset:offset+size]...)
				images = append(images, Image{
					Name: fmt.Sprintf("legacy-image-%03d.bmp", startIndex+len(images)+1),
					Ext:  ".bmp",
					Data: img,
				})
				offset += size - 1
			}
		}
	}
	for offset := 0; offset+40 <= len(data); offset += 4 {
		dibSize, _, dibSized := dibDeclaredSize(data[offset:])
		if dibSized && imageCandidateContained(imageCandidate{start: offset, end: offset + dibSize}, kept) {
			offset += dibSize - 4
			continue
		}
		img, ok := dibToBMP(data[offset:])
		if !ok {
			continue
		}
		images = append(images, Image{
			Name: fmt.Sprintf("legacy-image-%03d.bmp", startIndex+len(images)+1),
			Ext:  ".bmp",
			Data: img,
		})
		offset += len(img) - 14 - 4
	}
	for offset := 0; offset+88 <= len(data); offset += 4 {
		if binary.LittleEndian.Uint32(data[offset:]) != 1 {
			continue
		}
		size, ok := emfDeclaredSize(data[offset:])
		if ok && imageCandidateContained(imageCandidate{start: offset, end: offset + size}, kept) {
			offset += size - 4
			continue
		}
		if !ok || offset+size > len(data) {
			continue
		}
		img := append([]byte(nil), data[offset:offset+size]...)
		images = append(images, Image{
			Name: fmt.Sprintf("legacy-image-%03d.emf", startIndex+len(images)+1),
			Ext:  ".emf",
			Data: img,
		})
		offset += size - 4
	}
	for offset := 0; offset+18 <= len(data); offset++ {
		if !possibleWMFHeader(data[offset:]) {
			continue
		}
		size, ok := wmfDeclaredSize(data[offset:])
		if ok && imageCandidateContained(imageCandidate{start: offset, end: offset + size}, kept) {
			offset += size - 1
			continue
		}
		if !ok || offset+size > len(data) {
			continue
		}
		img := append([]byte(nil), data[offset:offset+size]...)
		images = append(images, Image{
			Name: fmt.Sprintf("legacy-image-%03d.wmf", startIndex+len(images)+1),
			Ext:  ".wmf",
			Data: img,
		})
		offset += size - 1
	}
	return images
}

func possibleWMFHeader(b []byte) bool {
	if len(b) < 18 {
		return false
	}
	if len(b) >= 40 && binary.LittleEndian.Uint32(b) == 0x9ac6cdd7 {
		return true
	}
	mtType := binary.LittleEndian.Uint16(b)
	if mtType != 1 && mtType != 2 {
		return false
	}
	if binary.LittleEndian.Uint16(b[2:]) != 9 {
		return false
	}
	version := binary.LittleEndian.Uint16(b[4:])
	return version == 0x0100 || version == 0x0300
}

var spaceRE = regexp.MustCompile(`[ \t\r\f\v]+`)
var blankLineRE = regexp.MustCompile(`\n{3,}`)
var mhtmlWhitespaceAfterPathSeparatorRE = regexp.MustCompile(`/[ \t\r\n]+`)
var mhtmlWhitespaceBeforePathSeparatorRE = regexp.MustCompile(`[ \t\r\n]+/`)
var wordHyperlinkFieldRE = regexp.MustCompile(`\b(?:HYPERLINK|INCLUDEPICTURE)\s+(?:"[^"]*"|\S+)(?:\s+\\[A-Za-z@#*]+(?:"[^"]*")?(?:\s+"[^"]*")?)*`)
var wordStyleRefFieldRE = regexp.MustCompile(`\bSTYLEREF(?:\s+\\[A-Za-z@#*]+)*(?:\s*"[^"]*")?(?:\s+\\[A-Za-z@#*]+(?:"[^"]*")?(?:\s+"[^"]*")?)*`)
var wordNamedFieldRE = regexp.MustCompile(`\b(?:PAGEREF|REF|NOTEREF|MERGEFIELD|DOCPROPERTY|DOCVARIABLE|STYLEREF)\s+(?:"[^"]*"|\S+)(?:\s+\\[A-Za-z@#*]+(?:"[^"]*")?(?:\s+"[^"]*")?)*`)
var wordLinkFieldRE = regexp.MustCompile(`\bLINK\s+(?:"[^"]+"|[A-Za-z0-9_.]+)(?:\s+(?:"[^"]*"|[A-Za-z]:\\\S+|\\\\\S+|https?://\S+))*(?:\s+\\[A-Za-z@#*]+(?:\s+\d+)*)*`)
var wordIncludeTextFieldRE = regexp.MustCompile(`\b(?:INCLUDETEXT|RD)\s+(?:"[^"]*"|\S+)(?:\s+\\[A-Za-z@#*]+(?:\s+(?:"[^"]*"|[A-Za-z0-9_.:-]+))?)*`)
var wordEmbedFieldRE = regexp.MustCompile(`\bEMBED\s+(?:"[^"]+"|[A-Za-z0-9_]*[_.][A-Za-z0-9_.]*|Equation)(?:\s+(?:"[^"]*"|\\[A-Za-z@#*]+|\S+))*`)
var wordMacroButtonFieldRE = regexp.MustCompile(`\bMACROBUTTON\s+\S+\s*`)
var wordTemplateFieldRE = regexp.MustCompile(`\bTEMPLATE\b(?:\s+\\\*\s*\w+|\s+\\[A-Za-z@#*]+(?:"[^"]*")?)*(?:\s+[^\s"<>|?*\\/]+\.(?:dotm|dotx|dot))?`)
var wordFormattedSimpleFieldRE = regexp.MustCompile(`\b(?:AUTHOR|CREATEDATE|DATE|TIME|FILENAME|FILESIZE|EDITTIME|PAGE|NUMPAGES|SECTIONPAGES|SUBJECT|KEYWORDS|COMMENTS|LASTSAVEDBY|PRINTDATE|SAVEDATE|TEMPLATE|USERNAME|USERINITIALS|SHAPE)\b(?:\s+\\\*\s*\w+)+`)
var wordSimpleFieldRE = regexp.MustCompile(`\b(?:AUTHOR|CREATEDATE|DATE|TIME|FILENAME|FILESIZE|EDITTIME|PAGE|NUMPAGES|SECTIONPAGES|SUBJECT|KEYWORDS|COMMENTS|LASTSAVEDBY|PRINTDATE|SAVEDATE|TEMPLATE|USERNAME|USERINITIALS|SHAPE|FORMTEXT|FORMCHECKBOX)\b(?:\s+\\[A-Za-z@#*]+(?:"[^"]*")?(?:\s+"[^"]*")?)*`)
var wordSymbolFieldRE = regexp.MustCompile(`\bSYMBOL\s+\d+(?:\s+\\[A-Za-z]+\s*(?:"[^"]*"|\S+)?)*`)
var wordTOCFieldRE = regexp.MustCompile(`\bTOC(?:\s+\\[A-Za-z@#*]+(?:"[^"]*")?(?:\s+"[^"]*")?)+`)
var wordTOCBookmarkFieldRE = regexp.MustCompile(`\bTOC\s+"[^"]*"\s*`)
var wordInternalBookmarkRE = regexp.MustCompile(`"?__RefHeading__\d+_\d+"?\s*`)
var wordTOCInternalBookmarkRE = regexp.MustCompile(`"?_Toc\d+"?\s*(?:\\h)?\s*`)
var wordSeqFieldRE = regexp.MustCompile(`\bSEQ\s+\w+(?:\s+\\\*\s*\w+|\s+\\[A-Za-z]+(?:\s+\d+)*)*`)
var wordIndexEntryFieldRE = regexp.MustCompile(`\b(?:XE|TC|TA)\s+"[^"]*"(?:\s+\\[A-Za-z]+\s*(?:"[^"]*"|\S+)?)*`)
var wordPromptFieldRE = regexp.MustCompile(`\bASK\s+(?:"[^"]*"|\S+)(?:\s+"[^"]*")?(?:\s+\\[A-Za-z]+(?:\s+"[^"]*"|\s+\S+)?)?|\bFILLIN(?:\s+"[^"]*"|\s+\S+)?(?:\s+\\[A-Za-z]+(?:\s+"[^"]*"|\s+\S+)?)?`)
var wordSetFieldRE = regexp.MustCompile(`\bSET\s+(?:"[^"]*"|\S+)\s+(?:"[^"]*"|\S+)`)
var wordIfFieldRE = regexp.MustCompile(`\bIF\s+(?:\{[^}]*\}|"[^"]*"|\S+)\s*(?:=|<>|>=|<=|>|<)\s*(?:\{[^}]*\}|"[^"]*"|\S+)(?:\s+"[^"]*"){1,2}(?:\s+\\[A-Za-z@#*]+(?:"[^"]*")?)*`)
var wordBibliographyFieldRE = regexp.MustCompile(`\b(?:CITATION|BIBLIOGRAPHY)\b(?:\s+(?:"[^"]*"|[^\s\\]+))?(?:\s+\\[A-Za-z@#*]+(?:\s+(?:"[^"]*"|[^\s\\]+))?)*`)
var wordDatabaseFieldRE = regexp.MustCompile(`\bDATABASE\b(?:\s+\\[A-Za-z@#*]+(?:\s+(?:"[^"]*"|\S+))?)*`)
var wordAdvanceFieldRE = regexp.MustCompile(`\bADVANCE\b(?:\s+\\[A-Za-z]+\s*-?\d+)+`)
var wordPrivateAddinFieldRE = regexp.MustCompile(`\b(?:ADDIN|PRIVATE)\b(?:\s+(?:"[^"]*"|\{[^}]*\}|\\[A-Za-z@#*]+(?:\s+(?:"[^"]*"|[^\s\\]+))?))*`)
var wordPicturePathFieldRE = regexp.MustCompile(`(?i)"?(?:(?:[A-Za-z]:\\|https?://|file://|://|\.\.?[\\/]|[^"\s\\/:\n]+[\\/])[^"\n]*|[^"\s\\/:\n]+)\.(?:gif|jpe?g|jpe|jfif|png|bmp|dib|wm[fz]|em[fz]|svgz?|eps|ps|tiff?|webp|ico|pcx|tga|pct|pict?|heic|heif|avif|wdp|jxr|hdp|jp2|jpx|jpf|j2[ck]|jpc)"?\s+(?:\\\*\s*)?MERGEFORMATINET\b`)
var wordFormatSwitchRE = regexp.MustCompile(`\\\*\s*(?:Upper|Lower|FirstCap|Caps|MERGEFORMAT)\b`)
var wordMergeFormatRE = regexp.MustCompile(`(?:\\\*\s*)?\bMERGEFORMAT(?:INET)?\b`)
var orphanWordFieldTokenRE = regexp.MustCompile(`\b(?:PAGEREF|REF|NOTEREF|MERGEFIELD|DOCPROPERTY|DOCVARIABLE|STYLEREF|AUTHOR|CREATEDATE|DATE|TIME|FILENAME|FILESIZE|EDITTIME|PAGE|NUMPAGES|SECTIONPAGES|FORMTEXT|FORMCHECKBOX|KEYWORDS|COMMENTS|LASTSAVEDBY|PRINTDATE|SAVEDATE|TEMPLATE|USERNAME|USERINITIALS|LINK|INCLUDETEXT|RD|SYMBOL|QUOTE|AUTOTEXT|AUTOTEXTLIST|LISTNUM|AUTONUM|AUTONUMLGL|AUTONUMOUT|ASK|FILLIN|SET|IF|CITATION|BIBLIOGRAPHY|DATABASE|ADVANCE|ADDIN|PRIVATE)\b`)
var inlineHiddenOfficeAssignmentRE = regexp.MustCompile(`(?i)\b(?:Id|Target|TargetMode|Type|ContentType|PartName|r:embed|r:link|r:id|embed|link|href|src|xmlns(?::[A-Za-z_][\w.-]*)?|mc:Ignorable|xsi:schemaLocation|schemaLocation)\s*=\s*(?:"[^"\n]*"|'[^'\n]*'|[^\s<>\n]+)\s*/?>?`)
var inlineHiddenOfficeColonAssignmentRE = regexp.MustCompile(`(?i)\b(?:Id|Target|TargetMode|Type|ContentType|PartName|r:embed|r:link|r:id|embed|link|href|src|Content-Location|Content-ID|Content-Type|Content-Transfer-Encoding|Content-Disposition|Content-Description|Content-Base|MIME-Version|mc:Ignorable|xsi:schemaLocation|schemaLocation)\s*:\s*(?:"[^"\n]*"|'[^'\n]*'|[^\s<>\n]+|<[^<>\n]+>)`)
var inlineHiddenMIMEParameterizedHeaderRE = regexp.MustCompile(`(?i)\bContent-Disposition\s*:\s*(?:inline|attachment|form-data)(?:\s*;\s*[A-Za-z0-9_-]+\s*=\s*(?:"[^"\n]*"|'[^'\n]*'|[^\s;]+))*`)
var inlineHiddenOfficeURLReferenceRE = regexp.MustCompile(`(?i)\burl\(\s*(?:"[^"\n]*"|'[^'\n]*'|[^)\n]*)\s*\)`)
var inlineWrappedHiddenOfficeReferenceRE = regexp.MustCompile(`(?:&lt;|<|\[|\(|\{)([^<>\[\](){}\s]{1,512})(?:&gt;|>|\]|\)|\})`)

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
	if len(b) < 16 || !bytes.Equal(b[10:14], []byte{0x00, 0x11, 0x02, 0xff}) {
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
