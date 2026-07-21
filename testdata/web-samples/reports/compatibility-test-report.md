# officeread 在线样本文档兼容性测试报告

测试日期：2026-06-24  
测试目标：从公开网络来源下载本程序支持的 Office 文档格式，每种格式抽取 1000 个样本，批量调用 `officeread.Extract` 验证兼容性。

## 1. 测试范围

本轮覆盖 README、测试和工具链中确认支持的 6 类格式：

| 格式 | 样本数 | 样本目录大小 |
|---|---:|---:|
| `.doc` | 1000 | 482.96 MB |
| `.docx` | 1000 | 174.89 MB |
| `.ppt` | 1000 | 1112.19 MB |
| `.pptx` | 1008，测试取前 1000 | 2675.23 MB |
| `.xls` | 1000 | 1101.16 MB |
| `.xlsx` | 1000 | 198.11 MB |

实际兼容性测试总量为 6000 个文件。

## 2. 样本来源

样本均来自公开可访问的在线语料或开源项目测试数据：

- GOVDOCS1 / Digital Corpora S3 by-type ZIP：
  - `https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/doc.zip`
  - `https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/ppt-part*.zip`
  - `https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/xls.zip`
  - `https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/docx.zip`
  - `https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/pptx.zip`
  - `https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/xlsx.zip`
- MSX-13 URL list：
  - `https://roussev.net/msx-13/msx-13--name-hash-size-url.txt`
- GitHub 开源测试语料补充：
  - `dotnet/Open-XML-SDK`
  - `LibreOffice/core`
  - `plutext/docx4j`
  - `apache/tika`
  - `closedxml/closedxml`

下载脚本和日志：

- 下载脚本：`tools/download_web_samples.py`
- GOVDOCS1 / MSX-13 下载日志：`testdata/web-samples/download-log.csv`
- GitHub 补充下载日志：`testdata/web-samples/github-download-log.csv`
- MSX-13 原始 URL 清单：`testdata/web-samples/sources/msx-13--name-hash-size-url.txt`

## 3. 测试方法

新增批量兼容性检查工具：

- `cmd/compatcheck/main.go`

核心判定：

- 对每个样本调用 `officeread.Extract`。
- 程序无 panic 且 `Extract` 返回 nil error，记为通过。
- `Extract` 返回 error，记为失败。
- 提取文本和图片数量仅作为观测指标；空文本输出不单独判为失败，因为部分样本可能本身为空文档、图片型文档或结构特殊文档。

执行命令：

```powershell
New-Item -ItemType Directory -Force -Path .gocache,testdata\web-samples\reports | Out-Null
$env:GOCACHE=(Resolve-Path .gocache).Path
go run -buildvcs=false ./cmd/compatcheck `
  -limit 1000 `
  -json testdata\web-samples\reports\compat-report.json `
  -csv testdata\web-samples\reports\compat-report.csv `
  testdata\web-samples\samples
```

原始结果：

- JSON：`testdata/web-samples/reports/compat-report.json`
- CSV：`testdata/web-samples/reports/compat-report.csv`

## 4. 测试结果

| 格式 | 测试数 | 通过 | 错误 | Panic | 空输出 | 文本量 | 图片数 | 总耗时 | 最慢单文件 |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| `.doc` | 1000 | 1000 | 0 | 0 | 2 | 43.68 MB | 3409 | 781.8 s | 10.1 s |
| `.docx` | 1000 | 1000 | 0 | 0 | 46 | 16.83 MB | 1607 | 273.7 s | 18.6 s |
| `.ppt` | 1000 | 1000 | 0 | 0 | 0 | 8.87 MB | 10560 | 251.7 s | 2.4 s |
| `.pptx` | 1000 | 999 | 1 | 0 | 0 | 10.91 MB | 24016 | 566.7 s | 9.6 s |
| `.xls` | 1000 | 1000 | 0 | 0 | 0 | 56.77 MB | 127 | 1050.9 s | 167.4 s |
| `.xlsx` | 1000 | 1000 | 0 | 0 | 0 | 325.58 MB | 350 | 2105.7 s | 1012.6 s |

汇总：

- 总测试文件：6000
- 通过：5999
- 失败：1
- Panic：0
- 通过率：99.9833%

## 5. 失败样本

| 文件 | 格式 | 错误 | 耗时 |
|---|---|---|---:|
| `testdata/web-samples/samples/pptx/00024860.pptx` | `.pptx` | `zip: checksum error` | 789 ms |

该失败来自 ZIP 校验错误，表现为 `Extract` 返回 error；没有发生 panic。

## 6. 性能观察

最慢样本集中在电子表格，特别是大文本量或结构复杂的 `.xlsx` / `.xls`：

| 文件 | 格式 | 耗时 | 提取文本量 |
|---|---|---:|---:|
| `testRecordSizeExceeded.xlsx` | `.xlsx` | 1012.6 s | 181333430 bytes |
| `008055.xls` | `.xls` | 167.4 s | 17684747 bytes |
| `00012389.xlsx` | `.xlsx` | 70.6 s | 11076610 bytes |
| `016161.xls` | `.xls` | 66.7 s | 8179182 bytes |
| `00014878.xlsx` | `.xlsx` | 60.0 s | 9758002 bytes |

结论：兼容性稳定性很好，但大表格解析存在明显长尾耗时，后续可针对 `.xlsx` / `.xls` 的超大单元格、共享字符串和行列遍历路径做性能优化或增加超时保护。

## 7. 回归验证

兼容性工具构建通过：

```powershell
go build -buildvcs=false ./cmd/compatcheck
```

全量 Go 测试通过：

```powershell
$env:GOCACHE=(Resolve-Path .gocache).Path
go test -buildvcs=false ./...
```

结果：

```text
ok      officeread                    401.930s
?       officeread/cmd/compatcheck    [no test files]
?       officeread/cmd/extracttest    [no test files]
?       officeread/cmd/officeread     [no test files]
```

## 8. 局限性

- 公开网络语料可能包含重复文件、损坏文件或非典型测试夹具；本轮保留这类样本，用于观察真实输入容错能力。
- 测试判定以“无 panic 且提取过程返回成功”为主，没有人工校验每个文件的语义提取完整性。
- `.pptx` 目录下载到 1008 个样本，本轮按统一规则测试前 1000 个；多出的 8 个保留在样本目录中。
- MSX-13 中部分历史 URL 已失效，因此 `.docx` / `.xlsx` 等格式使用 GitHub 开源项目测试文件补足到 1000 个。

## 9. 结论

`officeread` 在 6 类 Office 格式、共 6000 个公开网络样本上的兼容性结果为：

- 稳定性：0 panic。
- 成功率：5999 / 6000，约 99.9833%。
- 唯一失败：1 个 `.pptx` 文件因 ZIP checksum 错误返回 error。
- 主要改进方向：大 `.xlsx` / `.xls` 的极端耗时样本需要性能治理或可配置超时。
