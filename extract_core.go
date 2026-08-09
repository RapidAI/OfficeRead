package officeread

import (
	"bytes"
	"fmt"
	"html"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"net/url"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"strings"
	"time"
	"unicode"
	"unicode/utf8"
)

const maxOOXMLEmbeddedDepth = 3
const maxRepeatedTextPartBytes = 4096
const maxCompressedMetafileBytes = 256 << 20
const maxSmallDuplicateLegacyImageBytes = 4096
const maxImageFilenameBytes = 180
const maxMarkdownTableRows = 50000
const maxMarkdownTableCols = 1024
const maxMarkdownTableCellBytes = 512 << 10
const maxBIFFStoredTableRows = 50000
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
	ImageDir           string
	IncludeMetadata    bool
	StrictOfficeImages bool
	// StrictOfficeContent limits OOXML text to what Office's primary document
	// content API exposes, excluding cached drawing/chart data.
	StrictOfficeContent bool
	// OfficeFieldTime is an optional reference clock for evaluating Word DATE
	// fields in strict Office-content mode. COM compatibility tests provide the
	// moment at which Word produced its Content.Text baseline, so a picture that
	// includes seconds remains deterministic across the two processes.
	OfficeFieldTime time.Time
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
	coverageMarkdown := markdown
	// Legacy Word's structured Markdown deliberately preserves readable note
	// anchors, while its COM-aligned Result.Text omits them. Limit this
	// normalization to legacy note sections; doing it for every document can
	// change ordinary Markdown link and citation coverage.
	if (strings.Contains(markdown, "## Footnotes and Endnotes") || strings.Contains(markdown, "## Comments")) &&
		(strings.Contains(markdown, "[footnote]") || strings.Contains(markdown, "[comment]")) {
		coverageMarkdown = strings.ReplaceAll(coverageMarkdown, "[footnote]", "")
		coverageMarkdown = strings.ReplaceAll(coverageMarkdown, "[comment]", "")
		text = strings.ReplaceAll(text, "[footnote]", "")
		text = strings.ReplaceAll(text, "[comment]", "")
	}
	missing := missingMarkdownText(coverageMarkdown, text, images)
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
