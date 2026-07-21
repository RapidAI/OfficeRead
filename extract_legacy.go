package officeread

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"github.com/richardlehane/mscfb"
	textencoding "golang.org/x/text/encoding"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/encoding/korean"
	"golang.org/x/text/encoding/simplifiedchinese"
	"golang.org/x/text/encoding/traditionalchinese"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"math"
	"path"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"unicode"
	"unicode/utf16"
	"unicode/utf8"
)

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
