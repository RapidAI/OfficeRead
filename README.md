# officeread

`officeread` is a pure-Go Office document reader and command-line extractor. It extracts visible text, structured Markdown, and embedded images from Word, PowerPoint, and Excel files, making Office documents easier to preview, archive, or feed into AI助手 workflows.

中文文档见下方：[中文说明](#中文说明)。

## English

### What It Does

- Reads `.docx`, `.doc`, `.pptx`, `.ppt`, `.xlsx`, and `.xls` files.
- Extracts cleaned, user-visible text while filtering hidden runs, deleted revisions, field instructions, relationship noise, binary strings, and most non-content metadata.
- Produces structured Markdown for OOXML files and legacy `.xls` tables.
- Extracts embedded images and can write them to disk with stable, safe filenames.
- Keeps image references in Markdown aligned with the actual output filenames.
- Optionally includes document metadata with `IncludeMetadata` or `-metadata`.
- Handles embedded Office packages recursively, with a depth limit of 3.
- Avoids Microsoft Office, LibreOffice, COM automation, or external document converters.

### Supported Formats

| Family | Extensions | Extracted content |
| --- | --- | --- |
| Word | `.docx`, `.doc` | Body text, tables, headers/footers, footnotes, comments, visible form fields, visible images |
| PowerPoint | `.pptx`, `.ppt` | Slides, notes, visible master/layout text, images |
| Excel | `.xlsx`, `.xls` | Visible worksheets, cell text, tables, comments, headers/footers, chart/pivot labels, images |

### How It Works

`officeread` chooses the parser from the file container:

1. **OOXML files** (`.docx`, `.pptx`, `.xlsx`) are ZIP packages. The library reads the package parts directly, follows relationships, parses XML streams, shared strings, drawing parts, chart parts, notes, comments, and media references, then builds plain text, Markdown, and image outputs.
2. **Legacy Office files** (`.doc`, `.ppt`, `.xls`) are OLE Compound File containers. The library reads OLE streams in Go with `mscfb`, parses useful text streams and BIFF records where applicable, extracts legacy images, and falls back to safe text recovery for damaged or unusual streams.
3. **Markdown generation** normalizes text, renders tables, places images near matching text when possible, and appends any unplaced images under an `Images` section.
4. **Image handling** normalizes common Office image payloads, including PNG, JPEG, GIF, BMP/DIB, TIFF, SVG/SVGZ, EMF, WMF, WebP, ICO, PCX, PICT, EPS/PS, JPEG 2000, JPEG XR, HEIC/HEIF, and AVIF.

The goal is not to recreate the full Office object model. The goal is practical extraction of content a user would normally see.

### Advantages

- **Pure Go deployment**: easy to embed in services, workers, CLIs, and batch jobs without installing Office software.
- **Broad Office coverage**: supports both modern OOXML and Office 97-2003 binary formats.
- **AI-friendly output**: Markdown and cleaned text are designed for AI助手 experiences, previews, and downstream AI/NLP workflows.
- **Safer image output**: generated filenames are sanitized, deduplicated, extension-normalized, and URL-escaped in Markdown.
- **Resilient parsing**: tests include normal samples, malformed packages, encrypted files, XML bombs, fuzz/crash regressions, hidden content, metadata, embedded objects, and images.
- **Metadata control**: metadata is excluded by default so output stays close to visible content, but it can be enabled when needed.

### Install

The current module path is local:

```bash
go get officeread
```

When publishing this as a third-party module, change `go.mod` to the real repository path, for example `github.com/yourname/officeread`.

For local development:

```bash
go test ./...
go run ./cmd/officeread -markdown -images images sample.xlsx
```

### Command Line

```bash
go run ./cmd/officeread [options] file
```

| Option | Description |
| --- | --- |
| `-images dir` | Write extracted images to `dir` |
| `-metadata` | Include document properties, text relationships, custom XML, and related metadata |
| `-markdown` | Print Markdown instead of plain text; defaults `-images` to `images` when omitted |
| `-text-only` | Print only text/Markdown and suppress image count on stderr |

Examples:

```bash
go run ./cmd/officeread sample.docx
go run ./cmd/officeread -images out sample.pptx
go run ./cmd/officeread -markdown -images images sample.xlsx > sample.md
go run ./cmd/officeread -metadata sample.docx
```

Batch extraction helper:

```bash
go run ./cmd/extracttest -out extract-output testdata/samples
```

Compatibility report helper:

```bash
go run ./cmd/compatcheck -json compat-report.json -csv compat-report.csv testdata/samples
```

### API

```go
func Extract(filename string, opts Options) (*Result, error)
```

```go
type Options struct {
	ImageDir        string
	IncludeMetadata bool
}

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
```

Example:

```go
package main

import (
	"fmt"
	"log"

	"officeread"
)

func main() {
	res, err := officeread.Extract("sample.xlsx", officeread.Options{
		ImageDir: "images",
	})
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println(res.Text)
	fmt.Println("images:", len(res.Images))
	fmt.Println(res.Markdown("images"))
}
```

Use `Result.Markdown(imageBase)` to render full Markdown with image references:

```go
md := res.Markdown("images")
// ![alt](images/image1.png)

md = res.Markdown("https://cdn.example.com/assets")
// ![alt](https://cdn.example.com/assets/image1.png)
```

### Limits

- Password-protected Office documents are not decrypted.
- Layout, styling, formulas, macros, and the full Office object model are not fully reconstructed.
- Legacy `.xls` Markdown is simplified from BIFF table records.
- Unknown OOXML package types may fall back to XML text extraction.
- Markdown table output has safety limits to avoid pathological output from abnormal files.

### Tests

```bash
go test ./...
```

## 中文说明

`officeread` 是一个用 Go 编写的 Office 文档读取库，同时提供命令行工具。它可以从 Word、PowerPoint、Excel 文件中提取可见文本、结构化 Markdown 和嵌入图片，适合文档预览、AI助手、AI/RAG 预处理、批量归档、兼容性扫描等场景。

项目不依赖 Microsoft Office、LibreOffice、COM 自动化或外部转换器；现代 Office 文件直接读取 OOXML ZIP 包，Office 97-2003 二进制文件通过 OLE Compound File 流在纯 Go 中解析。

### 功能

- 支持 `.docx`、`.doc`、`.pptx`、`.ppt`、`.xlsx`、`.xls`。
- 提取清洗后的用户可见文本，过滤隐藏文本、删除修订、字段指令、资源路径、二进制噪声等非正文内容。
- 为 OOXML 文件和 legacy `.xls` 表格生成更适合阅读与 AI 输入的 Markdown。
- 提取嵌入图片，并可写入指定目录。
- Markdown 图片引用使用与落盘文件一致的安全文件名，自动处理重名、空格、URL 转义和危险文件名字符。
- 支持读取嵌入的 Office 包，递归深度限制为 3 层，避免无限嵌套。
- 元数据默认不进入正文；可通过 `IncludeMetadata` / `-metadata` 显式包含文档属性、关系中的文本链接、自定义 XML 等。
- 对损坏、加密、畸形包和 fuzz 样本做了防崩溃处理；不支持解密受密码保护的文档。

### 支持格式

| 类型 | 扩展名 | 提取内容 |
| --- | --- | --- |
| Word | `.docx`, `.doc` | 正文、表格、页眉页脚、脚注、批注、可见表单字段、可见图片等 |
| PowerPoint | `.pptx`, `.ppt` | 幻灯片、备注、母版/布局中的可见文本、图片等 |
| Excel | `.xlsx`, `.xls` | 可见工作表、单元格文本、表格、批注、页眉页脚、图表/数据透视表相关标签、图片等 |

### 原理

`officeread` 会根据文件容器选择解析路径：

1. **OOXML 文件**（`.docx`、`.pptx`、`.xlsx`）本质上是 ZIP 包。库会直接读取包内 XML、关系文件、共享字符串、绘图、图表、备注、批注和媒体部件，综合生成纯文本、Markdown 和图片输出。
2. **Legacy Office 文件**（`.doc`、`.ppt`、`.xls`）本质上是 OLE Compound File。库使用 `mscfb` 读取 OLE 流，解析可用的文本流和 BIFF 记录，提取 legacy 图片，并对损坏或非常规流做安全文本恢复。
3. **Markdown 生成**会规范化文本、渲染表格、尽量把图片放回相关文本附近；未能定位的图片会统一追加到 `Images` 小节。
4. **图片处理**会识别和规范化 Office 中常见图片载荷，包括 PNG、JPEG、GIF、BMP/DIB、TIFF、SVG/SVGZ、EMF、WMF、WebP、ICO、PCX、PICT、EPS/PS、JPEG 2000、JPEG XR、HEIC/HEIF、AVIF 等。

本库的目标不是完整还原 Office 对象模型，而是稳定提取“用户通常能看到的内容”。

### 优势

- **纯 Go 部署**：适合嵌入服务、队列 worker、CLI 和批处理任务，不需要安装 Office 软件。
- **格式覆盖广**：同时覆盖现代 OOXML 与 Office 97-2003 二进制格式。
- **面向 AI助手**：输出清洗文本与 Markdown，便于接入 AI助手、生成预览、RAG 入库和后续 NLP 处理。
- **图片输出安全稳定**：图片文件名会被清理、去重、规范扩展名，并在 Markdown 中正确转义。
- **解析韧性强**：测试覆盖正常样本、畸形包、加密文件、XML bomb、fuzz/crash 回归、隐藏内容、元数据、嵌入对象和图片。
- **元数据可控**：默认只输出接近用户可见内容的正文，需要审计或归档时再显式打开元数据。

### 安装

当前模块名为 `officeread`：

```bash
go get officeread
```

如果要作为第三方依赖发布，建议先将 `go.mod` 中的 module path 改为实际仓库地址，例如 `github.com/yourname/officeread`。

本仓库内开发或试用：

```bash
go test ./...
go run ./cmd/officeread -markdown -images images sample.xlsx
```

### 命令行工具

主命令位于 `cmd/officeread`。

```bash
go run ./cmd/officeread [选项] file
```

| 参数 | 说明 |
| --- | --- |
| `-images dir` | 将提取到的图片写入 `dir` 目录 |
| `-metadata` | 在文本中包含文档属性、文本关系、自定义 XML 等元数据 |
| `-markdown` | 输出 Markdown，而不是纯文本；如果未指定 `-images`，默认使用 `images` 目录 |
| `-text-only` | 只输出文本/Markdown，不在 stderr 输出图片数量 |

示例：

```bash
go run ./cmd/officeread sample.docx
go run ./cmd/officeread -images out sample.pptx
go run ./cmd/officeread -markdown -images images sample.xlsx > sample.md
go run ./cmd/officeread -metadata sample.docx
```

批量测试/导出工具：

```bash
go run ./cmd/extracttest -out extract-output testdata/samples
```

兼容性扫描工具：

```bash
go run ./cmd/compatcheck -json compat-report.json -csv compat-report.csv testdata/samples
```

### API 文档

```go
func Extract(filename string, opts Options) (*Result, error)
```

读取指定 Office 文件并返回提取结果。

```go
type Options struct {
	ImageDir        string
	IncludeMetadata bool
}

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
```

字段说明：

| 字段 | 说明 |
| --- | --- |
| `Options.ImageDir` | 非空时，将 `Result.Images` 中的图片写入该目录，目录会自动创建 |
| `Options.IncludeMetadata` | 为 `true` 时额外提取文档属性、可见文本链接、自定义 XML、legacy OLE 属性等元数据 |
| `Result.Text` | 清洗后的纯文本 |
| `Result.StructuredMarkdown` | 内部生成的结构化 Markdown 片段，通常使用 `Markdown(imageBase)` 获取完整输出 |
| `Result.Images` | 提取到的图片列表，即使未设置 `ImageDir`，图片数据也保留在内存中 |
| `Image.Name` | 安全化后的输出文件名 |
| `Image.Alt` | 从图片描述、标题或关系中提取的替代文本 |
| `Image.Ext` | 标准化后的扩展名，例如 `.png`、`.jpg`、`.svg` |
| `Image.Data` | 图片二进制内容 |

示例：

```go
package main

import (
	"fmt"
	"log"

	"officeread"
)

func main() {
	res, err := officeread.Extract("sample.xlsx", officeread.Options{
		ImageDir: "images",
	})
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println(res.Text)
	fmt.Println("images:", len(res.Images))
	fmt.Println(res.Markdown("images"))
}
```

`Result.Markdown(imageBase)` 返回完整 Markdown，并用 `imageBase` 控制图片链接前缀：

```go
md := res.Markdown("images")
// ![alt](images/image1.png)

md = res.Markdown("https://cdn.example.com/assets")
// ![alt](https://cdn.example.com/assets/image1.png)
```

### 错误与限制

- 本库不解密受密码保护的 Office 文档。
- 文本提取目标是“用户可见文本”，不是完整 Office 对象模型解析器；布局、样式、公式计算结果和宏不会被完整还原。
- `.xls` Markdown 支持来自 BIFF 表格解析，复杂格式会被简化为 Markdown 表格。
- OOXML 中未知类型包会退化为 XML 文本提取。
- 表格 Markdown 有安全上限，避免异常文件产生过大输出。

### 测试

```bash
go test ./...
```

测试集包含大量正常样本和负向样本，覆盖六种目标扩展名、OOXML/legacy Office、图片、嵌入对象、元数据、隐藏内容、畸形包、XML bomb、加密文件和 fuzz/crash 回归样本。
