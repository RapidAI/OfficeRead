# officeread follow-up report (2026-07-05)

## Scope

This follow-up records four pieces of evidence gathered against the current worktree on 2026-07-05:

- current sample-pool coverage for supported online document formats
- rerun status of the historical only failing compatibility sample
- current focused performance baseline rerun for the retained hotspot workloads
- retained performance optimizations validated by same-day hotspot and broad reruns

## Sample coverage

The online sample pool currently satisfies the original target of 1000 samples for each supported
format:

| Format | Sample count |
|---|---:|
| `.doc` | 1000 |
| `.docx` | 1000 |
| `.ppt` | 1000 |
| `.pptx` | 1008 |
| `.xls` | 1000 |
| `.xlsx` | 1000 |

Source evidence:

- `tools/download_web_samples.py`
- `testdata/web-samples/samples/*`

## Compatibility follow-up

### Historical failing sample rerun

The historical compatibility report had exactly one known failure:

- `testdata/web-samples/samples/pptx/00024860.pptx`
- previous error: `zip: checksum error`

Root-cause inspection on 2026-07-05 showed that the package itself was readable and the corruption
was isolated to one media member:

- corrupt entry: `ppt/media/image45.png`
- failure type: bad CRC on the image payload

Code change:

- OOXML image extraction now skips unreadable/corrupt media members instead of failing the whole
  document extraction path.

Targeted rerun command:

```powershell
& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck `
  -json testdata\web-samples\reports\compat-rerun-20260705-failing-pptx-after-image-skip.json `
  -csv testdata\web-samples\reports\compat-rerun-20260705-failing-pptx-after-image-skip.csv `
  testdata\web-samples\samples\pptx\00024860.pptx
```

Observed result after the fix:

- `00024860.pptx`: `ok=true`, `textBytes=2864`, `images=47`, `error=""`, `panic=""`, `282 ms`

Regression coverage:

```powershell
& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run "TestCorruptOOXMLImageCRCDoesNotFailExtraction|TestInvalidOOXMLImageIsDropped|TestMalformedImageAltXMLDoesNotFailExtraction|TestExtractedSampleImagesAreValid|TestWrittenSampleImagesAreValid" -count=1 .
```

Result:

```text
ok  	officeread	58.883s
```

### Fresh full compatibility sweep

A fresh per-format compatibility rerun was completed on 2026-07-05 against the current worktree.

Output files:

- `testdata/web-samples/reports/compat-doc-20260705.json`
- `testdata/web-samples/reports/compat-docx-20260705.json`
- `testdata/web-samples/reports/compat-ppt-20260705.json`
- `testdata/web-samples/reports/compat-pptx-20260705.json`
- `testdata/web-samples/reports/compat-xls-20260705.json`
- `testdata/web-samples/reports/compat-xlsx-20260705.json`

Observed result:

| Format | Total | OK | Errors | Panics | Empty | `>10s` | Max ms | Avg ms |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| `.doc` | 1000 | 1000 | 0 | 0 | 2 | 7 | 105010 | 646.68 |
| `.docx` | 1000 | 1000 | 0 | 0 | 46 | 0 | 6175 | 84.45 |
| `.ppt` | 1000 | 1000 | 0 | 0 | 0 | 0 | 1897 | 69.12 |
| `.pptx` | 1000 | 1000 | 0 | 0 | 0 | 1 | 14544 | 313.85 |
| `.xls` | 1000 | 1000 | 0 | 0 | 0 | 1 | 12670 | 171.25 |
| `.xlsx` | 1000 | 1000 | 0 | 0 | 0 | 0 | 5701 | 88.44 |

Aggregate:

- full sweep pass rate: `6000 / 6000`
- errors: `0`
- panics: `0`

Slowest `.doc` files before the retained UTF-16 optimization:

| File | Millis | Text bytes |
|---|---:|---:|
| `001538-2.doc` | 105010 | 430692 |
| `001538.doc` | 104216 | 430692 |
| `005313.doc` | 42069 | 291520 |
| `003336.doc` | 40224 | 268880 |
| `003336-2.doc` | 31774 | 268880 |

## Current focused performance rerun

The retained performance baseline was remeasured on 2026-07-05 using the standard focused hotspot
sets.

### 6-file `.xls` hotspot batch

Command:

```powershell
& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck `
  -json testdata\web-samples\reports\perf-rerun-20260705-current-xls6.json `
  -csv testdata\web-samples\reports\perf-rerun-20260705-current-xls6.csv `
  testdata\web-samples\samples\xls\008055.xls `
  testdata\web-samples\samples\xls\013623.xls `
  testdata\web-samples\samples\xls\016161.xls `
  testdata\web-samples\samples\xls\018548.xls `
  testdata\web-samples\samples\xls\019088.xls `
  testdata\web-samples\samples\xls\019089.xls
```

Observed total:

- current rerun total: `16514 ms`
- retained baseline reference: `15947 ms`

### 30-file mixed keyset

Command output files:

- `testdata/web-samples/reports/perf-rerun-20260705-current-keyset.json`
- `testdata/web-samples/reports/perf-rerun-20260705-current-keyset.csv`

Observed totals:

- current rerun total: `54578 ms`
- current rerun `.xls` subset: `20668 ms`
- retained baseline reference total: `48484 ms`
- retained baseline reference `.xls` subset: `18047 ms`

Output stability check against the retained baseline:

- keyset `textBytes` comparison: `NO_DIFF`

Slowest files in the current rerun:

| File | Format | Millis | Text bytes |
|---|---|---:|---:|
| `008055.xls` | `.xls` | 6117 | 17684747 |
| `testRecordSizeExceeded.xlsx` | `.xlsx` | 4111 | 181333430 |
| `00003763.docx` | `.docx` | 4008 | 452979 |
| `016161.xls` | `.xls` | 3289 | 8179182 |
| `00012389.xlsx` | `.xlsx` | 2890 | 11148655 |

Interpretation:

- this rerun reaffirms that the dominant remaining cost is still concentrated in large spreadsheet
  markdown backfill paths rather than OOXML image extraction
- current hotspot direction remains `missingMarkdownText -> markdownBackfillExactLinesCovered ->
  markdownBackfillExactSet -> markdownVisibleTableCells`

## Retained `.doc` hotspot optimization

The fresh full compatibility sweep exposed a new dominant legacy `.doc` long-tail in the UTF-16
fallback scan path. CPU profiling on `testdata/web-samples/samples/doc/001538-2.doc` showed that
almost all time was being spent inside:

- `hasUTF16Evidence(...)`
- `utf16Strings(...)`

Root cause:

- `utf16Strings(...)` repeatedly rescanned the same growing raw byte slice to recompute UTF-16
  evidence, which turned large fallback ranges into an avoidable quadratic-style hotspot.

Retained change:

- `utf16Strings(...)` now keeps incremental zero-byte and printable-ASCII counters and uses them for
  the UTF-16 evidence check instead of rescanning `raw` on each decision point.

Targeted validation:

- single-file profile before:
  - `001538-2.doc`: `96464 ms`
- single-file profile after:
  - `001538-2.doc`: `692 ms`
- hotspot-set rerun after rebuilding `compatcheck.exe`:
  - `001538-2.doc`: `716 ms`
  - `001538.doc`: `713 ms`
  - `003336-2.doc`: `416 ms`
  - `003336.doc`: `422 ms`
- output stability:
  - `001538*.doc` text bytes unchanged: `430692`
  - `003336*.doc` text bytes unchanged: `268880`

Optimization-after full `.doc` rerun:

- output files:
  - `testdata/web-samples/reports/compat-doc-20260705-after-utf16-evidence-cache.json`
  - `testdata/web-samples/reports/compat-doc-20260705-after-utf16-evidence-cache.csv`
- aggregate result:
  - before: `total=1000`, `ok=1000`, `errors=0`, `panics=0`, `empty=2`, `over10Sec=7`,
    `maxMillis=105010`, `millis=646683`, `avgMillis=646.68`
  - after: `total=1000`, `ok=1000`, `errors=0`, `panics=0`, `empty=2`, `over10Sec=0`,
    `maxMillis=7186`, `millis=220879`, `avgMillis=220.88`
  - total time reduction: `425804 ms`
  - relative speedup: `65.84%`
- slow-file comparison:
  - `001538-2.doc`: `105010 ms -> 740 ms`
  - `001538.doc`: `104216 ms -> 652 ms`
  - `003336-2.doc`: `31774 ms -> 374 ms`
  - `003336.doc`: `40224 ms -> 378 ms`
  - `005313.doc`: `42069 ms -> 634 ms`

Regression coverage:

```text
ok  	officeread	223.658s
```

## 2026-07-06 xlsx follow-up

This round continued from the retained 2026-07-06 `simpleInlineWorksheetCandidate(...)` tag-scan
optimization and focused on the remaining `.xlsx` worksheet markdown-preparation hotspot around
`appendWorksheetText(...)`.

### Retained baseline before this round

The already-retained `simpleInlineWorksheetCandidate(...)` tag-oriented scan remains the current
baseline for this branch.

Evidence files:

- `testdata/web-samples/reports/perf-exp-ai-assistant-tagscan-candidate-xlsx-pair.json`
- `testdata/web-samples/reports/perf-exp-ai-assistant-tagscan-candidate-keyset.json`
- `testdata/web-samples/reports/compat-xlsx-20260706-after-tagscan-candidate.json`

Observed retained baseline improvements:

- focused `.xlsx` pair total: `7017 ms -> 5217 ms`
- `00012389.xlsx`: `3223 ms -> 2244 ms`
- `testRecordSizeExceeded.xlsx`: `3794 ms -> 2973 ms`
- broader mixed keyset total: `52414 ms -> 41602 ms`
- broader mixed keyset `.xlsx` subset: `26286 ms -> 22324 ms`
- output parity against the previous retained baseline: `NO_DIFF`
- `.xlsx` full compatibility after the tag-scan retain: `1000 / 1000`, `errors=0`, `panics=0`

### Newly retained worksheet markdown row-slice optimization

Root cause:

- the remaining `00012389.xlsx` cost was still concentrated in worksheet markdown row collection
  and row finalization work inside `appendWorksheetText(...)`
- the old path stored markdown row values in `map[int]string`, even though worksheet rows are
  collected in ascending column order and are later compacted into a contiguous row slice

Retained change:

- switch worksheet markdown row collection from `map[int]string` to `[]string`
- compact trailing empty cells with `compactPreparedWorksheetMarkdownRow([]string)`
- preserve existing output behavior and markdown row limits

Current code locations:

- `extract.go`
- `appendWorksheetText(...)`
- `compactPreparedWorksheetMarkdownRow(...)`

Focused benchmark and pair evidence:

- isolated `.xlsx` pair evidence:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-rowslice-xlsx-pair-isolated.json`
  - total: `4933 ms`
  - `00012389.xlsx`: `2037 ms`
  - `testRecordSizeExceeded.xlsx`: `2896 ms`
- relative to the retained tag-scan baseline pair:
  - pair total: `5217 ms -> 4933 ms`
  - `00012389.xlsx`: `2244 ms -> 2037 ms`
  - `testRecordSizeExceeded.xlsx`: `2973 ms -> 2896 ms`

Broader isolated keyset evidence:

- evidence file:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-rowslice-keyset-isolated-clean.json`
- totals:
  - current total: `39627 ms`
  - current `.xlsx` subset: `20801 ms`
- relative to the retained tag-scan baseline:
  - total: `41602 ms -> 39627 ms`
  - `.xlsx` subset: `22324 ms -> 20801 ms`

Output stability:

- keyset parity check against the retained tag-scan baseline: `NO_DIFF`
- compared fields: `textBytes`, `images`

Compatibility follow-up:

- output files:
  - `testdata/web-samples/reports/compat-xlsx-20260706-after-worksheet-rowslice.json`
  - `testdata/web-samples/reports/compat-xlsx-20260706-after-worksheet-rowslice.csv`
- aggregate result:
  - `.xlsx`: `total=1000`, `ok=1000`, `errors=0`, `panics=0`, `empty=0`, `over10Sec=0`,
    `maxMillis=3158`, `millis=55540`

Regression coverage:

```text
ok  	officeread	(cached)
?   	officeread/cmd/compatcheck	[no test files]
?   	officeread/cmd/extracttest	[no test files]
?   	officeread/cmd/officeread	[no test files]
?   	officeread/cmd/profileextract	[no test files]
```

Decision:

- retain the worksheet markdown row-slice optimization
- the evidence is consistent across focused reruns, isolated mixed-keyset reruns, output parity,
  and full `.xlsx` compatibility
- earlier concurrent row-slice runs remain excluded from decision-making because they were polluted
  by overlapping heavy perf activity

## 2026-07-06 rejected: markdown table-cell RTF prefix fold helper

- experiment:
  - in `prepareMarkdownTableCellValue(...)`, replace the per-cell
    `strings.HasPrefix(strings.ToLower(trimmed), ...)` checks for `{\rtf` / `\rtf` with the
    existing allocation-free `hasPrefixFold(...)` helper
- rationale:
  - `00012389.xlsx` still shows a large markdown-preparation delta versus the no-markdown worksheet
    benchmark
  - the simple single-line cell fast path runs on a large number of markdown cells, so removing two
    `strings.ToLower(...)` allocations per cell looked promising

Validation:

- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$' -benchmem -count=3 ./`
- isolated pair rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-markdown-rtf-prefix-xlsx-pair-isolated.json`
- isolated mixed keyset rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-markdown-rtf-prefix-keyset-isolated.json`
- parity comparison:
  - compared against `perf-exp-ai-assistant-rowslice-keyset-isolated-clean.json`

Observed results:

- focused `00012389.xlsx` worksheet benchmark improved:
  - before:
    - `1195405100 / 1156936700 / 1181891500 ns/op`
    - `574233000 / 573144264 / 573430416 B/op`
    - `9823036 / 9822747 / 9822846 allocs/op`
  - experiment:
    - `1113008800 / 1120007800 / 1122634100 ns/op`
    - `560224592 / 559984024 / 560425224 B/op`
    - `9172733 / 9172644 / 9172728 allocs/op`
- no-markdown benchmark remained in the same general range and confirmed the markdown side was the
  changed path
- isolated `.xlsx` pair improved slightly:
  - baseline pair: `4933 ms`
  - experiment pair: `4891 ms`
  - `00012389.xlsx`: `2037 ms -> 2152 ms` on this rerun shape, but pair median still landed
    slightly lower because `testRecordSizeExceeded.xlsx` improved to `2739 ms`
- parity result:
  - `NO_DIFF`

Decisive integrated result:

- isolated mixed keyset total:
  - baseline total: `39627 ms`
  - experiment total: `39569 ms`
- but the relevant `.xlsx` subset regressed:
  - baseline `.xlsx`: `20801 ms`
  - experiment `.xlsx`: `20969 ms`

Interpretation:

- the helper swap clearly reduced allocations and improved the focused worksheet benchmark
- however, once rerun in the retained isolated mixed keyset, the `.xlsx` subset moved in the wrong
  direction
- because the experiment only touches the `.xlsx` markdown cell path, the mixed total improvement
  is not strong enough to offset the `.xlsx` subset regression

Decision:

- Reverted.

## 2026-07-06 rejected: shared-string-only RTF prefix fold helper

- experiment:
  - keep the retained generic `prepareMarkdownTableCellValue(...)` unchanged
  - add the allocation-free `{\rtf` / `\rtf` prefix check only on the shared-string worksheet
    markdown path in `appendSharedStringWorksheetTextPrepared(...)`
  - leave the simple-inline worksheet markdown path on the retained generic helper so
    `testRecordSizeExceeded.xlsx` would not inherit the broad helper change
- rationale:
  - the earlier global helper swap proved that the `strings.ToLower(...)` allocations inside
    markdown preparation were real work
  - but the pair regression suggested the win was being paid for on the simple-inline side, so the
    next step was to narrow the helper change to the shared-string path only

Validation:

- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused serial rerun on the retained tree before the experiment:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXSharedStringMarkdownPrepare00012389$|BenchmarkXLSXSharedStringPrepared00012389$|BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$' -benchmem -benchtime=1x ./`
- targeted tests:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run 'TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue|TestExtractXLSX|TestExtractWorksheet|Test.*Text|Test.*Visible' ./`
- focused serial rerun after narrowing the helper:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXSharedStringPrepared00012389$|BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$' -benchmem -benchtime=1x ./`
- decisive pair rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-shared-rtf-fold-pair-serial.json`

Observed results:

- retained-tree serial baseline before the narrowing:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `957514600 ns/op`, `149250416 B/op`,
    `1502156 allocs/op`
  - `BenchmarkXLSXSharedStringPrepared00012389`: `656889700 ns/op`, `149165656 B/op`,
    `1502124 allocs/op`
  - `BenchmarkXLSXSharedStringMarkdownPrepare00012389`: `106169600 ns/op`, `13221176 B/op`,
    `650347 allocs/op`
- narrowed shared-string helper improved the focused shared-string side:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `858380300 ns/op`, `135347632 B/op`,
    `851933 allocs/op`
  - `BenchmarkXLSXSharedStringPrepared00012389`: `608464500 ns/op`, `135128744 B/op`,
    `851881 allocs/op`
- targeted tests remained green:
  - `ok   officeread`
- decisive pair still regressed against the retained baseline:
  - retained baseline pair total: `4574 ms`
  - experiment pair total: `5127 ms`
  - `00012389.xlsx`: `1758 ms -> 1650 ms`
  - `testRecordSizeExceeded.xlsx`: `2816 ms -> 3477 ms`
  - pair output parity remained clean:
    - `00012389.xlsx text=11148655 images=0`
    - `testRecordSizeExceeded.xlsx text=181333430 images=0`

Interpretation:

- the narrowed helper confirmed that the shared-string markdown preparation path itself can benefit
  materially from removing the per-cell lowercase allocations
- however, even after avoiding the simple-inline helper rewrite, the integrated pair still moved in
  the wrong direction because `testRecordSizeExceeded.xlsx` slowed down more than `00012389.xlsx`
  sped up
- that makes this another case where the focused shared-string win does not translate into the
  retained pair gate

Decision:

- Reverted.

## 2026-07-06 rejected: shared-string-only single-pass markdown escape helper

- experiment:
  - keep `cleanMarkdownTableCellValue(...)` and the generic `prepareMarkdownTableCellValue(...)`
    unchanged
  - only on the shared-string worksheet markdown path in
    `appendSharedStringWorksheetTextPrepared(...)`, replace the final two
    `strings.ReplaceAll(..., "\\", "\\\\")` / `strings.ReplaceAll(..., "|", "\\|")`
    passes with one shared helper that:
    - returns the original string immediately when there is no `\` or `|`
    - otherwise scans once and emits escaped output for both marker families
- rationale:
  - fresh `BenchmarkXLSXSharedStringMarkdownPrepare00012389` profiling still showed the prepare
    stage as a meaningful slice of the remaining shared-string markdown delta
  - unlike the earlier rejected “plain-text direct return” and prefix-fold swaps, this attempt
    keeps the exact same prepare semantics and only narrows the final escaping work

Validation:

- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- targeted tests:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run 'TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue|TestExtractXLSX|TestExtractWorksheet|Test.*Text|Test.*Visible' ./`
- focused serial benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXSharedStringMarkdownPrepare00012389$|BenchmarkXLSXSharedStringPrepared00012389$|BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$' -benchmem -benchtime=1x ./`
  - repeated once more for the shared-string prepared / worksheet pair in the same session
- decisive pair rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-shared-escape-pair-serial.json`

Observed results:

- targeted tests stayed green:
  - `ok   officeread`
- focused shared-string benchmarks improved in the expected direction:
  - `BenchmarkXLSXSharedStringMarkdownPrepare00012389`:
    - retained baseline earlier in the same session: `106169600 ns/op`
    - experiment: `70045700 ns/op`
  - `BenchmarkXLSXSharedStringPrepared00012389`:
    - retained baseline earlier in the same session: `656889700 ns/op`
    - experiment reruns: `652670100 ns/op`, then `634358700 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    - retained baseline earlier in the same session: `957514600 ns/op`
    - experiment reruns: `916600200 ns/op`, then `916320700 ns/op`
- decisive pair still regressed versus the retained gate:
  - retained baseline pair total: `4574 ms`
  - experiment pair total: `5400 ms`
  - `00012389.xlsx`: `1758 ms -> 1736 ms`
  - `testRecordSizeExceeded.xlsx`: `2816 ms -> 3664 ms`
  - pair output parity remained clean:
    - `00012389.xlsx text=11148655 images=0`
    - `testRecordSizeExceeded.xlsx text=181333430 images=0`

Interpretation:

- this confirms that the final markdown escaping work in the shared-string path is real and can be
  reduced in focused benches without changing output
- but, just like the earlier shared-string-local prepare experiments, the integrated pair moved in
  the wrong direction because `testRecordSizeExceeded.xlsx` slowed down more than the shared-string
  hotspot sped up
- that makes this another false friend: the local prepare win is not sufficient to beat the
  retained pair baseline

Decision:

- Reverted.

## 2026-07-06 rejected: reuse visible xlsx comment items across text + markdown phases

- experiment:
  - during one `.xlsx` extraction, parse visible comment items once and reuse them between:
    - text phase comment collection
    - structured markdown comment section generation
- rationale:
  - full `Extract(00012389.xlsx)` profiling showed a non-trivial tail in:
    - `xlsxVisibleCommentsText`
    - `xlsxVisibleCommentPartSourcesUncached`
    - `visibleXlsxCommentItems(...)`
  - source inspection confirmed that the text path and markdown path each re-open and re-parse the
    same visible comment XML parts

Validation:

- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$' -benchmem -count=3 ./`
- serial focused pair rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-comment-reuse-xlsx-pair-isolated-serial.json`
- serial `.xlsx` keyset rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-comment-reuse-xlsx-keyset-isolated-serial.json`
- parity comparison:
  - compared against `perf-exp-ai-assistant-rowslice-keyset-isolated-clean.json`
- note:
  - earlier concurrent reruns named `perf-exp-ai-assistant-comment-reuse-*.json` and
    `compat-xlsx-20260706-after-comment-reuse.json` were treated as invalid exploratory runs and
    were not used for the decision

Observed results:

- focused extract benchmark was noisy and not decisive:
  - experiment:
    - `2068472600 / 2237816900 / 2915861700 ns/op`
    - around `13.49M allocs/op`
  - the worksheet-only benchmark was expectedly unchanged in semantics and also noisy
- output parity stayed clean:
  - `NO_DIFF`

Decisive serial integrated result:

- focused `.xlsx` pair regressed:
  - baseline pair: `4933 ms`
  - experiment pair: `5117 ms`
  - `00012389.xlsx`: `2037 ms -> 2216 ms`
  - `testRecordSizeExceeded.xlsx`: `2896 ms -> 2901 ms`
- `.xlsx` keyset regressed heavily:
  - baseline `.xlsx`: `20801 ms`
  - experiment `.xlsx`: `27577 ms`

Interpretation:

- the duplicate comment parsing looked real in code, but removing it did not help the retained
  integrated `.xlsx` workload
- the cost of carrying and reusing the comment-item structures was worse than the repeated parsing
  it was meant to replace, at least on the retained serial workload

Decision:

- Reverted.

## 2026-07-06 rejected: worksheet markdown prepared-cell cache

- experiment:
  - add a worksheet-local cache for markdown table-cell preparation:
    - key: raw markdown cell text collected in `appendWorksheetText(...)` /
      `appendSimpleInlineWorksheetTextPrepared(...)`
    - value: final prepared markdown table-cell text after
      `cleanMarkdownTableCellValue(...)` and `prepareMarkdownTableCellValue(...)`
- rationale:
  - repeated profiling kept the dominant remaining `00012389.xlsx` gap in the markdown cell cleanup
    and preparation chain
  - a targeted temporary analysis on `00012389.xlsx` showed very high repetition in collected
    markdown cells:
    - rows: `31340`
    - total collected cells: `1416419`
    - non-empty cells: `567787`
    - unique non-empty cells: `131754`
    - repeated unique cells: `26721`
    - repeated uses across those repeated cells: `462754`
  - top repeated values included:
    - `Not OA` (`29244`)
    - `Journal` (`29111`)
    - `Active` (`19818`)
    - `Social Sciences` (`12296`)
    - `Health Sciences` (`11903`)

Validation:

- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$' -benchmem -count=3 ./`
- isolated pair reruns:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-mdcell-cache-xlsx-pair-isolated.json`
  - `testdata/web-samples/reports/perf-exp-ai-assistant-mdcell-cache-xlsx-pair-isolated-rerun.json`
- isolated mixed keyset rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-mdcell-cache-keyset-isolated.json`
- isolated `.xlsx`-only keyset rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-mdcell-cache-xlsx-keyset-isolated.json`
- full `.xlsx` compatibility rerun:
  - `testdata/web-samples/reports/compat-xlsx-20260706-after-mdcell-cache.json`
- parity checks:
  - compared against `perf-exp-ai-assistant-rowslice-keyset-isolated-clean.json`

Observed results:

- focused benchmark improved materially:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`
    - `1053138300 / 1017769600 / 1307922700 ns/op`
    - `581352288 / 580601536 / 580302264 B/op`
    - `9202727 / 9202571 / 9202457 allocs/op`
  - compared with the retained row-slice baseline:
    - baseline allocs were around `9.82M/op`
    - experiment allocs dropped to about `9.20M/op`
- second isolated pair rerun improved:
  - baseline pair: `4933 ms`
  - rerun pair: `4604 ms`
  - `00012389.xlsx`: `2037 ms -> 1854 ms`
  - `testRecordSizeExceeded.xlsx`: `2896 ms -> 2750 ms`
- isolated `.xlsx`-only keyset improved:
  - baseline `.xlsx`: `20801 ms`
  - experiment `.xlsx`: `19000 ms`
- parity stayed clean:
  - `NO_DIFF`

Conflicting integrated evidence:

- the first isolated pair rerun regressed:
  - `4933 ms -> 5968 ms`
- mixed keyset total regressed even though the `.xlsx` subset improved:
  - baseline total: `39627 ms`
  - experiment total: `41001 ms`
  - baseline `.xlsx`: `20801 ms`
  - experiment `.xlsx`: `18997 ms`
- most importantly, the full `.xlsx` corpus rerun regressed on wall time:
  - baseline full `.xlsx` rerun:
    - `testdata/web-samples/reports/compat-xlsx-20260706-after-worksheet-rowslice.json`
    - `millis=55540`, `maxMillis=3158`
  - experiment full `.xlsx` rerun:
    - `testdata/web-samples/reports/compat-xlsx-20260706-after-mdcell-cache.json`
    - `millis=62156`, `maxMillis=3765`
  - compatibility itself remained green:
    - `.xlsx`: `1000 / 1000`, `errors=0`, `panics=0`

Interpretation:

- the duplicate-cell evidence was real, and the cache did improve focused and selected `.xlsx`
  reruns
- however, the broader end-to-end full `.xlsx` sweep moved in the wrong direction by a meaningful
  margin
- under the retained decision standard, the full integrated `.xlsx` regression outweighs the local
  and subset wins

Decision:

- Reverted.

## 2026-07-06 rejected: remove redundant markdown single-line control check

- experiment:
  - in `cleanMarkdownTableCellValue(...)`, remove the trailing
    `!maybeControlFragmentText(value)` guard from the single-line fast path and rely on
    `!maybeDiscardableHiddenOfficeText(value)` alone
- rationale:
  - profile evidence on `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx` showed the single-line
    branch spending notable time in:
    - `cleanMarkdownTableCellValue(...)`
    - `maybeDiscardableHiddenOfficeText(...)`
    - `maybeControlFragmentText(...)`
  - source inspection showed `maybeDiscardableHiddenOfficeText(...)` already calls
    `maybeHiddenOrControlText(...)`, which in turn falls through to `maybeControlFragmentText(...)`,
    making the second explicit check look redundant

Validation:

- profiling:
  - `tmp-shape/bench-00012389-xlsx-text.cpu`
  - `tmp-shape/bench-00012389-xlsx-nomd.cpu`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$' -benchmem -count=3 ./`
- isolated pair rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-markdown-singlecheck-xlsx-pair-isolated.json`
- isolated mixed keyset rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-markdown-singlecheck-keyset-isolated.json`
- parity comparison:
  - compared against `perf-exp-ai-assistant-rowslice-keyset-isolated-clean.json`

Observed results:

- focused worksheet benchmark improved:
  - baseline:
    - `1195405100 / 1156936700 / 1181891500 ns/op`
    - `574233000 / 573144264 / 573430416 B/op`
    - `9823036 / 9822747 / 9822846 allocs/op`
  - experiment:
    - `1141803900 / 1134629600 / 1166671000 ns/op`
    - `573675872 / 573071944 / 573422792 B/op`
    - `9821461 / 9821316 / 9821388 allocs/op`
- output parity:
  - `NO_DIFF`

Decisive integrated result:

- isolated `.xlsx` pair regressed:
  - baseline pair: `4933 ms`
  - experiment pair: `5082 ms`
- isolated mixed keyset regressed:
  - baseline total: `39627 ms`
  - experiment total: `40396 ms`
  - baseline `.xlsx`: `20801 ms`
  - experiment `.xlsx`: `22247 ms`

Interpretation:

- despite the local benchmark improvement, the integrated `.xlsx` workload moved sharply in the
  wrong direction
- the source-level redundancy is real, but the explicit second check still appears to help the
  retained end-to-end balance, likely through branch shape or earlier short-circuit behavior on the
  real worksheet mix

Decision:

- Reverted.

## Retained `.pptx` OOXML lookup optimization

The next dominant OOXML hotspot after the `.doc` fix was concentrated in `.pptx` relationship and
part lookup churn.

CPU profiling on `testdata/web-samples/samples/pptx/00022381.pptx` before the retained change
showed the main cost clustered in:

- `ooxmlPartKeyCandidates(...)`
- `ooxmlPartKeyMatches(...)`
- `ooxmlFile(...)`
- `ooxmlCleanPartName(...)`
- `ooxmlEscapePartName(...)`

Root cause:

- `pptxVisibleRelatedTextParts(...)` and its recursive helpers repeatedly recomputed normalized part
  names and fallback key candidates for the same OOXML package members while walking visible slide
  relationships.

Retained change:

- add a per-extraction `ooxmlLookup` index built from the package file map
- thread that lookup through the hot `.pptx` visible-related-text traversal so repeated part-name
  normalization and fallback candidate matching can be reused instead of rebuilt on every step

## 2026-07-06 rerun: current retained baseline remains noisy but stable enough

Fresh repeat-aware reruns were collected on 2026-07-06 against the current retained worktree
before trying any new code change.

Output files:

- `testdata/web-samples/reports/perf-rerun-20260706-current-xlsx-pair.json`
- `testdata/web-samples/reports/perf-rerun-20260706-current-xlsx-pair.csv`
- `testdata/web-samples/reports/perf-rerun-20260706-current-keyset-repeat3.json`
- `testdata/web-samples/reports/perf-rerun-20260706-current-keyset-repeat3.csv`

Observed results:

- focused `.xlsx` pair rerun:
  - total `.xlsx`: `7017 ms`
  - `00012389.xlsx`: `3223 ms`, `min=2447`, `max=3335`, `textBytes=11148655`, `images=0`
  - `testRecordSizeExceeded.xlsx`: `3794 ms`, `min=3684`, `max=4319`, `textBytes=181333430`,
    `images=0`
- repeat-aware 30-file mixed keyset rerun:
  - total: `52414 ms`
  - `.docx` subset: `8896 ms`
  - `.xls` subset: `17232 ms`
  - `.xlsx` subset: `26286 ms`
  - all `textBytes` / `images` remained aligned with the retained baseline outputs

Interpretation:

- `00012389.xlsx` and `testRecordSizeExceeded.xlsx` remain the main `.xlsx` timing anchors
- run-to-run noise is still meaningful on this machine, so new candidates still need both focused
  pair evidence and broader keyset confirmation before retention

## 2026-07-06 rejected: simple-inline `.xlsx` cleanText reuse across text and markdown paths

Experiment:

- in `appendSimpleInlineWorksheetTextPrepared(...)`, compute the worksheet text-cleaning result
  once and reuse it for the markdown cell-cleaning path instead of calling `cleanText(...)` twice
  for the same inlineStr cell

Why:

- the current `testRecordSizeExceeded.xlsx` hotspot split still shows both
  `appendWorksheetValue(...)` and `cleanMarkdownTableCellValue(...)` as first-order cost centers
- the idea was to remove one full `cleanText(...)` pass per simple-inline cell without changing the
  later markdown filtering steps

Validation:

- focused hotspot benchmarks:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx -count=3`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx -count=3`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx -count=3`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx -count=3`
- focused repeat-aware pair rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-cleantext-reuse-xlsx-pair.json`
  - `testdata/web-samples/reports/perf-exp-ai-assistant-cleantext-reuse-xlsx-pair.csv`
- broader repeat-aware mixed keyset rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-cleantext-reuse-keyset-rerun.json`
  - `testdata/web-samples/reports/perf-exp-ai-assistant-cleantext-reuse-keyset-rerun.csv`
- targeted regression:
  - `go test -run 'Test.*XLSX|Test.*Xlsx|Test.*Excel|Test.*Worksheet|Test.*Markdown|Test.*Visible|Test.*Text' -count=1 ./`
- repository regression after revert:
  - `go test ./...`

Observed results:

- focused pair rerun looked mildly positive and preserved output parity:
  - baseline pair: `7017 ms`
  - experiment pair: `6820 ms`
  - `00012389.xlsx`: `3223 ms -> 3161 ms`
  - `testRecordSizeExceeded.xlsx`: `3794 ms -> 3659 ms`
- but the benchmark signal was mixed rather than decisively better:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`:
    `2779553500 / 3148206400 / 3067274200 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    `2981400500 / 3033095600 / 2758831700 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`:
    `2433843400 / 2358677000 / 2553582300 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    `1815033000 / 1910429400 / 1720473000 ns/op`
- and the broader repeat-aware mixed keyset regressed despite identical outputs:
  - baseline mixed keyset total: `52414 ms`
  - experiment mixed keyset total: `55400 ms`
  - baseline `.xlsx` subset: `26286 ms`
  - experiment `.xlsx` subset: `29901 ms`
  - `00012389.xlsx` inside the broader keyset worsened sharply:
    `2976 ms -> 4373 ms`
  - `testRecordSizeExceeded.xlsx` improved there:
    `3794 ms -> 2992 ms`
  - keyset `textBytes` / `images`: `NO_DIFF`

Interpretation:

- this is another case where a targeted `.xlsx` pair improvement did not survive the broader mixed
  retained workload
- the gain was too concentrated in `testRecordSizeExceeded.xlsx`, while `00012389.xlsx` became
  materially worse under the wider rerun

Decision:

- Reverted.

## 2026-07-06 current baseline re-profile and benchmark-noise cleanup

Fresh focused `.xlsx` profiles were captured again on 2026-07-06 against the retained baseline:

- `tmp-shape/xlsx-00012389-current-20260706.prof`
- `tmp-shape/xlsx-testrecord-simpleinline-current-20260706.prof`
- `tmp-shape/xlsx-testrecord-textonly-current-20260706.prof`

Observed current benchmark baselines:

- `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1484923733 ns/op`
- `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`: `2068850333 ns/op`
- `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx`: `1302769767 ns/op`

Current profile split:

- `00012389.xlsx` remains primarily a general worksheet-reader / markdown-preparation workload:
  - `appendWorksheetText(...)`: `81.14%` cumulative
  - `encoding/xml.(*Decoder).RawToken(...)`: `39.86%` cumulative
  - `cleanText(...)`: `11.71%` cumulative
  - `cleanMarkdownTableCellValue(...)`: `15.43%` cumulative
  - `prepareMarkdownTableCellValue(...)`: `5.57%` cumulative
- `testRecordSizeExceeded.xlsx` simple-inline markdown path still concentrates inside one pipeline:
  - `appendSimpleInlineWorksheetTextPrepared(...)`: `70.19%` cumulative
  - `appendWorksheetValue(...)`: `27.10%` cumulative
  - `cleanText(...)`: `27.88%` cumulative
  - `cleanMarkdownTableCellValue(...)`: `14.95%` cumulative
  - `prepareMarkdownTableCellValue(...)`: `6.56%` cumulative
- `testRecordSizeExceeded.xlsx` text-only path confirms that removing markdown does materially help,
  but the remaining text path is still dominated by:
  - `appendSimpleInlineWorksheetTextPrepared(...)`: `61.01%` cumulative
  - `appendWorksheetValue(...)`: `33.09%` cumulative
  - `cleanText(...)`: `28.78%` cumulative

Interpretation:

- the two main `.xlsx` hotspots still have different shapes:
  - `00012389.xlsx`: XML decoder plus worksheet markdown preparation
  - `testRecordSizeExceeded.xlsx`: simple-inline cell pipeline, especially repeated text cleaning
- this refresh strengthens the earlier conclusion that future retained work should target a broader
  structural reduction rather than another helper-level micro-opt

### Benchmark harness cleanup

While reviewing the new profiles, the simple-inline benchmark helpers were found to perform an
extra pre-timer `simpleInlineWorksheetCandidate(...)` verification pass after already selecting a
positive worksheet. That extra call did not affect `ns/op`, but it did pollute cpuprofiles.

Code change:

- remove the redundant verification call in:
  - `benchmarkXLSXSimpleInlineSampleFile(...)`
  - `benchmarkXLSXSimpleInlineTextOnlySampleFile(...)`

Validation:

- rerun the focused simple-inline benchmarks with cpuprofile:
  - `tmp-shape/xlsx-bench-pollution-check-20260706.prof`
- repository regression:
  - `go test ./...`

Observed result:

- the benchmark body numbers stayed in the same rough range, while the remaining
  `simpleInlineWorksheetCandidate(...)` time in cpuprofile is now attributable to the one-time
  worksheet-selection setup rather than a second redundant rescan

Decision:

- Retained as benchmark-fidelity cleanup only; it does not change production extraction behavior.

## 2026-07-06 retained: tag-oriented `simpleInlineWorksheetCandidate(...)` for `.xlsx`

Experiment:

- replace the current byte-by-byte `simpleInlineWorksheetCandidate(...)` scan with a tag-oriented
  validator:
  - jump between XML tags with `IndexByte('<')` / `IndexByte('>')`
  - reject the same broad classes of disqualifying worksheet features while inspecting actual tags
  - keep the existing fast-path extraction body unchanged

Why:

- a fresh full-extract profile on
  `testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
  showed that this precheck is still a real production cost, not just benchmark setup noise:
  - `simpleInlineWorksheetCandidate(...)`: `19.78%` cumulative in full extract
- the hotspot shape suggested that keeping the same fast path but making the admission scan cheaper
  was a realistic structural win

Focused evidence before applying:

- full extract:
  - `BenchmarkExtractXLSXHotspots/...testRecordSizeExceeded.xlsx`: `3405481067 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `3052220600 ns/op`
- simple-inline focused:
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`: `2068850333 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx`: `1302769767 ns/op`
- repeat-aware current baseline reruns:
  - focused `.xlsx` pair:
    - `testdata/web-samples/reports/perf-rerun-20260706-current-xlsx-pair.json`
    - total `.xlsx`: `7017 ms`
  - mixed 30-file keyset:
    - `testdata/web-samples/reports/perf-rerun-20260706-current-keyset-repeat3.json`
    - total: `52414 ms`
    - `.xlsx` subset: `26286 ms`

Validation after applying:

- focused benchmarks:
  - `BenchmarkExtractXLSXHotspots/...testRecordSizeExceeded.xlsx -count=3`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx -count=3`
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx -count=3`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx -count=3`
- repeat-aware focused pair:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-tagscan-candidate-xlsx-pair.json`
  - `testdata/web-samples/reports/perf-exp-ai-assistant-tagscan-candidate-xlsx-pair.csv`
- repeat-aware mixed keyset:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-tagscan-candidate-keyset.json`
  - `testdata/web-samples/reports/perf-exp-ai-assistant-tagscan-candidate-keyset.csv`
- output parity against the current retained keyset rerun:
  - `NO_DIFF` for `textBytes/images`
- targeted regression:
  - `go test -run 'Test.*XLSX|Test.*Xlsx|Test.*Excel|Test.*Worksheet|Test.*Markdown|Test.*Visible|Test.*Text' -count=1 ./`
- repository regression:
  - `go test ./...`
- full-format compatibility recheck for the affected format:
  - `testdata/web-samples/reports/compat-xlsx-20260706-after-tagscan-candidate.json`
  - `testdata/web-samples/reports/compat-xlsx-20260706-after-tagscan-candidate.csv`

Observed results:

- focused full-extract benchmarks improved materially:
  - `testRecordSizeExceeded.xlsx`: `3405481067 ns/op -> 3016184500 / 2739670400 / 2766510000`
  - `00012389.xlsx`: `3052220600 ns/op -> 2275989100 / 2299209300 / 2208585600`
- focused simple-inline benchmarks remained positive:
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`:
    `1942252800 / 2095717100 / 2064004200 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx`:
    about `1.24-1.27 s/op`
- focused repeat-aware pair improved strongly with matching outputs:
  - baseline pair total: `7017 ms`
  - experiment pair total: `5217 ms`
  - `00012389.xlsx`: `3223 ms -> 2244 ms`
  - `testRecordSizeExceeded.xlsx`: `3794 ms -> 2973 ms`
- broader repeat-aware mixed keyset also improved strongly:
  - baseline total: `52414 ms`
  - experiment total: `41602 ms`
  - baseline `.xlsx` subset: `26286 ms`
  - experiment `.xlsx` subset: `22324 ms`
  - baseline `.docx` subset: `8896 ms`
  - experiment `.docx` subset: `6310 ms`
  - baseline `.xls` subset: `17232 ms`
  - experiment `.xls` subset: `12968 ms`
- full `.xlsx` compatibility remained clean:
  - `total=1000`, `ok=1000`, `errors=0`, `panics=0`, `empty=0`
  - `millis=71205`, `maxMillis=4055`, `over10Sec=0`

Interpretation:

- unlike several earlier `.xlsx` micro-opts, this change survived all three bars:
  - focused hotspot benchmarks
  - repeat-aware focused pair rerun
  - broader mixed keyset rerun
- the strong `NO_DIFF` output check plus the 1000-sample `.xlsx` compatibility sweep makes this
  robust enough to keep

Decision:

- Retained.

## 2026-07-06 rejected: single-segment markdown cell shortcut in `appendWorksheetText(...)`

Experiment:

- for worksheet markdown collection in `appendWorksheetText(...)`, store the first markdown cell
  fragment directly in a temporary string and only fall back to `strings.Builder` when a cell
  receives multiple text segments

Why:

- `00012389.xlsx` is now mainly a worksheet markdown-preparation problem
- the idea was to reduce `markdownCellText.WriteString(...)` and `markdownCellText.String()` churn on
  the common case where a worksheet cell only contributes one text fragment

Validation:

- focused benchmarks only, before spending pair / keyset runs:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx -count=3`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx -count=3`
  - guard sample:
    - `BenchmarkExtractXLSXHotspots/...testRecordSizeExceeded.xlsx -count=3`
- targeted regression:
  - `go test -run 'Test.*XLSX|Test.*Xlsx|Test.*Excel|Test.*Worksheet|Test.*Markdown|Test.*Visible|Test.*Text' -count=1 ./`
- repository regression after revert:
  - `go test ./...`

Observed results:

- `00012389.xlsx` regressed immediately at the focused benchmark layer:
  - full extract:
    - retained current: about `2208585600..2299209300 ns/op`
    - candidate: `2623091300 / 2810685000 / 2904131400 ns/op`
  - worksheet hotspot:
    - retained current: `1484923733 ns/op`
    - candidate: `1793445800 / 1834013100 / 1847650900 ns/op`
- the guard sample was also not clearly better:
  - `testRecordSizeExceeded.xlsx` full extract:
    `2914910000 / 3467534200 / 3759054200 ns/op`

Interpretation:

- this shortcut removed some local builder work, but it did not survive the real worksheet hotspot
  and likely traded one allocation shape for another in the common path

Decision:

- Reverted before broader validation.

Validation:

- repository formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused single-file profiles:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-00022381-pptx-after-ooxml-lookup.pprof testdata\web-samples\samples\pptx\00022381.pptx`
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-rerun.pprof testdata\web-samples\samples\xls\006087.xls`
- focused `.pptx` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\compat-pptx-hotset-after-ooxml-lookup.json -csv testdata\web-samples\reports\compat-pptx-hotset-after-ooxml-lookup.csv testdata\web-samples\samples\pptx\00022381.pptx testdata\web-samples\samples\pptx\00026428.pptx testdata\web-samples\samples\pptx\00020725.pptx testdata\web-samples\samples\pptx\406730.pptx`
- full 1000-file `.pptx` rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata\web-samples\reports\compat-pptx-20260705-after-ooxml-lookup.json -csv testdata\web-samples\reports\compat-pptx-20260705-after-ooxml-lookup.csv -limit 1000 testdata\web-samples\samples\pptx`

Observed results:

- single-file hotspot:
  - `00022381.pptx`: `10738 ms -> 606 ms`
- focused 4-file hotspot rerun:
  - `00022381.pptx`: `14544 ms -> 545 ms`
  - `00026428.pptx`: `6759 ms -> 823 ms`
  - `00020725.pptx`: `6505 ms -> 957 ms`
  - `406730.pptx`: `6320 ms -> 1505 ms`
- output stability on the hotspot rerun:
  - `textBytes`: unchanged for all 4 files
  - `images`: unchanged for all 4 files
- full 1000-file `.pptx` rerun:
  - before: `millis=313850`, `avgMillis=313.85`, `over10Sec=1`, `maxMillis=14544`
  - after: `millis=139507`, `avgMillis=139.51`, `over10Sec=0`, `maxMillis=1865`
  - total time reduction: `174343 ms`
  - relative speedup: `55.55%`

Current slowest `.pptx` files after the retained change:

| File | Millis | Text bytes | Images |
|---|---:|---:|---:|
| `00022153.pptx` | 1865 | 27737 | 38 |
| `216691.pptx` | 1858 | 7769 | 9 |
| `00025988.pptx` | 1740 | 5652 | 30 |
| `00022144.pptx` | 1645 | 12907 | 11 |
| `00026993.pptx` | 1599 | 55253 | 50 |

Interpretation:

- the previous `.pptx` long tail was largely OOXML part-lookup overhead rather than XML decoding
  alone
- with that cost reduced, the remaining `.pptx` baseline is now short enough that `.xls`
  markdown-backfill work again stands out as the clearest next optimization target

## Retained `.xls` markdown-backfill query-cache optimization

After the `.pptx` rerun, the clearest remaining spreadsheet hotspot was still the repeated
containment work inside `missingMarkdownText(...)`.

Retained change:

- cache repeated coverage / containment queries inside one `missingMarkdownText(...)` call
- deduplicate candidate variants before running the expensive visible-line containment checks

Validation evidence:

- full regression:
  - `ok officeread 149.920s`
- focused hotspot profile:
  - `006087.xls`: `12279 ms -> 5326 ms`
- repeated 6-file `.xls` hotspot rerun:
  - before: `15566 ms`
  - after: `15242 ms`
  - relative speedup: `2.08%`
- mixed 30-file keyset rerun:
  - before: `50489 ms`
  - after: `48656 ms`
  - relative speedup: `3.63%`
  - `.xls` subset: `18325 ms -> 17763 ms`
- output stability:
  - keyset retained-vs-experiment file outputs: `NO_DIFF`

Additional rejected follow-up on the same hotspot:

- tried shrinking the BIFF markdown-backfill source to cell text only, excluding sheet names,
  header/footer text, and comment text already represented elsewhere in markdown
- result: clean output parity, but clear performance regression
  - retained 6-file `.xls` hotspot batch: `15242 ms -> 20070 ms`
  - mixed 30-file keyset rerun: `48656 ms -> 54907 ms`
  - mixed keyset `.xls` subset: `17763 ms -> 21835 ms`
- decision: reverted; the query-cache optimization remains the retained `.xls` baseline

- tried a cheaper visible-line prefilter: a global byte-presence mask for joined visible markdown
  lines, used to reject impossible `visibleLineContainsLine(...)` probes before `strings.Contains`
- result: improved the hottest single file, but still regressed the retained `.xls` workloads
  - `006087.xls`: `5326 ms -> 5159 ms`
  - retained 6-file `.xls` hotspot batch: `15242 ms -> 17346 ms`
  - mixed 30-file keyset rerun: `48656 ms -> 49392 ms`
  - mixed keyset `.xls` subset: `17763 ms -> 20311 ms`
- decision: reverted; the retained `.xls` query-cache baseline still wins on the target workload

- tried deduplicating the containment haystacks (`tableRaw`, `tableVisible`, `tableComparable`,
  `visibleLines`) before joining them for substring checks
- the hottest sample really did have a lot of repetition
  - `006087.xls` visible lines: `54601 total`, `48582 unique`, `6019 duplicates`
  - `006087.xls` table lines: `49999 total`, `43980 unique`, `6019 duplicates`
- result: strong single-file win, but broad workload loss again
  - `006087.xls`: `5326 ms -> 4778 ms`
  - retained 6-file `.xls` hotspot batch: `15242 ms -> 17657 ms`
  - mixed 30-file keyset rerun: `48656 ms -> 53685 ms`
  - mixed keyset `.xls` subset: `17763 ms -> 20827 ms`
- decision: reverted; duplicate haystack removal overfit the single hotspot and lost on the retained
  benchmark mix

Measurement tooling update:

- `cmd/compatcheck` now supports `-repeat N`
  - when `N > 1`, each file is extracted `N` times
  - `millis` in JSON/CSV becomes the per-file median
  - JSON/CSV also record `minMillis`, `maxMillis`, and raw `runs`
  - repeat runs also guard output stability by checking `textBytes/images` across runs
- follow-up fix:
  - odd repeat counts still use the middle element
  - even repeat counts now use the standard median (integer average of the two middle runs), rather
    than the earlier upper-middle shortcut
- this was added because recent `.xls` hotspot reruns showed too much single-run wall-time variance
  to trust one-shot numbers for close decisions
- first repeat-aware baseline check on the current retained `.xls` state:
  - command: `compatcheck -repeat 3 ...xls-hotspot-6-files...`
  - 6-file `.xls` hotspot median total: `19093 ms`
  - per-file medians:
    - `008055.xls`: `6847 ms` (`5847..6873`)
    - `013623.xls`: `2759 ms` (`2464..2836`)
    - `016161.xls`: `3024 ms` (`2947..3104`)
    - `018548.xls`: `2133 ms` (`2018..2194`)
    - `019088.xls`: `2161 ms` (`2087..2215`)
    - `019089.xls`: `2169 ms` (`2066..2275`)
- decision: keep using the retained query-cache parser baseline, but evaluate future close-call
  `.xls` experiments with repeat-aware median runs instead of one-shot totals

Hotspot input-shape note:

- inspected the raw text inputs feeding `.xls` markdown backfill on two hotspot samples by exporting
  plain-text extraction and counting candidate-line shape
- result:
  - `006087.xls`
    - non-empty text lines: `41347`
    - unique trimmed lines: `41346`
    - duplicate trimmed lines: `1`
    - plain-line ratio (rough `markdownPlainVisibleLine` approximation): `0.73%`
  - `008055.xls`
    - non-empty text lines: `1181198`
    - unique trimmed lines: `1181198`
    - duplicate trimmed lines: `0`
    - plain-line ratio: `0.01%`
- implication:
  - backfill input lines are effectively unique on the retained hotspot files
  - future `.xls` work should focus on reducing the cost of containment checks themselves or their
    invocation pattern, not on candidate-line deduplication or a plain-line-only fast path

- tried removing small-slice allocation in per-line variant assembly inside `missingMarkdownText`
  by using a fixed `[4]string` array plus an explicit count
- result:
  - narrow repeat-aware `.xls6` hotspot batch improved:
    - baseline median total: `19093 ms`
    - fixed-array variant assembly: `18374 ms`
  - but the broader repeat-aware retained keyset regressed:
    - baseline total: `51517 ms`
    - fixed-array variant assembly: `54008 ms`
  - baseline `.xls` subset: `17264 ms`
  - fixed-array variant assembly `.xls` subset: `20639 ms`
- decision: reverted; another case where the `.xls6` hotspot set alone would have been misleading

- tried progressive variant checking in `missingMarkdownText`
  - check the base line first
  - only compute/check the next variant if the earlier one was not already covered
  - stop generating later variants as soon as a variant covers the line
- result: catastrophic regression on the actual `.xls` hotspot set
  - `006087.xls`: `5326 ms -> 18652 ms`
  - repeat-aware `.xls6` hotspot median total: `19093 ms -> 37584 ms`
  - `008055.xls`: `6847 ms -> 13924 ms`
  - `013623.xls`: `2759 ms -> 8583 ms`
- decision: reverted immediately; this is not a promising direction

Benchmarking update:

- added committed hotspot benchmarks in [extract_bench_test.go](/D:/workprj/officeread/extract_bench_test.go)
  so future `.xls` optimization work can use `go test -bench` directly, not only compatcheck JSON
  reruns
- current benchmark set covers:
  - `006087.xls`
  - `008055.xls`
  - `016161.xls`
- benchmark smoke baseline (`-benchtime=1x`):
  - `006087.xls`: `5033971000 ns/op`, `714237984 B/op`, `6098297 allocs/op`
  - `008055.xls`: `7935391100 ns/op`, `1977264776 B/op`, `6044348 allocs/op`
  - `016161.xls`: `3613330100 ns/op`, `890202280 B/op`, `3389250 allocs/op`
- implication:
  - future parser changes should be screened both against repeat-aware compatcheck medians and
    against benchmark allocation deltas on these retained hotspot files

- tried replacing the per-raw-line `candidateLineCache` in `missingMarkdownText` with an in-place
  normalized `lines` slice reuse
- result was mixed in a way that does not justify keeping it:
  - good:
    - `repeat=3` `.xls6` hotspot median total: `19093 ms -> 15556 ms`
    - repeat-aware keyset total: `51517 ms -> 49659 ms`
  - bad:
    - `repeat=2` `.xls6` hotspot median total: `15954 ms -> 20178 ms`
    - keyset `.xls` subset: `17264 ms -> 17642 ms`
  - `006087.xls` benchmark: `5033971000 ns/op -> 5218890900 ns/op`
- decision: reverted; too inconsistent to count as a real retained `.xls` win

- tried preallocating `biffTextParts` output storage from BIFF payload size to reduce `appendPart`
  growth churn
- result again looked excellent on the narrow hotspot measurements:
  - `006087.xls` memory: `714237984 B/op -> 672588824 B/op`
  - `008055.xls` memory: `1977264776 B/op -> 1540425272 B/op`
  - repeat-aware `.xls6` hotspot median total: `19093 ms -> 16237 ms`
- but the broader repeat-aware keyset was decisively worse:
  - total: `51517 ms -> 60898 ms`
  - `.xls` subset: `17264 ms -> 25346 ms`
  - `.docx` subset: `6423 ms -> 8492 ms`
- decision: reverted; another optimization that overfit the hotspot slice and failed the retained
  wider workload

- tried eliding `missingMarkdownText` variant work by checking the base/visible/markdown forms
  progressively and only building the escaped table-cell variant for lines containing `|`
- result: not stable enough to keep
  - progressive variant branch:
    - repeat-aware `.xls6` hotspot median total: `23111 ms -> 22637 ms`
    - broader repeat-aware keyset total: `51517 ms -> 55521 ms`
    - keyset `.docx` subset: `6423 ms -> 7533 ms`
  - escaped-table-only guard:
    - repeat-aware `.xls6` hotspot median total: `23111 ms -> 20124 ms`
    - broader repeat-aware keyset total: `51517 ms -> 58143 ms`
    - keyset `.xlsx` subset: `27830 ms -> 31201 ms`
- decision: both experiments were reverted; they preserved `textBytes` / `images`, but neither
  survived the retained 30-file repeat-aware keyset

- profiled the current retained `006087.xls` baseline again and confirmed the shape of the hotspot:
  - source text lines: `41347` non-empty, average length `6.02`, only `5` single-character lines
  - generated markdown lines: `54601` non-empty, `49999` contain `|`
  - CPU profile still concentrated in
    `missingMarkdownText -> markdownBackfillContainment.visibleLineContainsLine -> strings.Contains`
- tried splitting table-cell short exact matches away from `visibleLines`
  - first branch:
    - stop adding table rows to `visibleLines`
    - add all visible table cells to a new exact-match set used before the `< 12 runes` table guard
  - second branch:
    - same change, but only cache short table cells (`< 12` runes)
- result: both branches overfit the `.xls` hotspot batch
  - all-cells branch:
    - `repeat=3` `.xls6` hotspot median total: `23111 ms -> 19458 ms`
    - broader repeat-aware keyset total: `51517 ms -> 58602 ms`
    - keyset `.xls` subset: `17264 ms -> 21689 ms`
  - short-cells-only branch:
    - `repeat=3` `.xls6` hotspot median total: `23111 ms -> 16658 ms`
    - broader repeat-aware keyset total: `51517 ms -> 59085 ms`
    - keyset `.xls` subset: `17264 ms -> 21772 ms`
- decision: both branches were reverted; removing table rows from `visibleLines` appears too costly on
  the retained wider workload even when hotspot `.xls` files get dramatically faster

- tried a narrower short-table-cell exact-hit branch without changing `visibleLines`
  - keep the retained containment structure intact
  - only add a short visible table-cell exact set
  - probe that exact set before the existing `< 12 runes` early return in `tableTextContainsLine`
- result: negative immediately on the hotspot batch, so the experiment was stopped before the wider
  keyset rerun
  - focused regression suite: `ok officeread 13.733s`
  - hotspot benchmarks:
    - `006087.xls`: `5573310400 ns/op -> 6289948800 ns/op`
    - `008055.xls`: `7905695900 ns/op -> 7716413100 ns/op`
    - `016161.xls`: `3466805000 ns/op -> 3567597900 ns/op`
  - `repeat=3` `.xls6` hotspot median total: `23111 ms -> 24728 ms`
- decision: reverted immediately; adding short table-cell exact hits without removing the existing
  visible-line path only adds build/probe overhead on the retained `.xls` hotspot workload

- tried raising the `visibleLineContainsLine` minimum query length from 4 runes to 6 runes
  - this was the cheapest possible query-side filter: no new index structures, no markdown rebuild,
    just a narrower gate before the `visibleJoined` substring search
- result: immediately negative on the hotspot batch
  - focused regression suite: `ok officeread 13.364s`
  - hotspot benchmarks:
    - `006087.xls`: `5573310400 ns/op -> 7056552700 ns/op`
    - `008055.xls`: `7905695900 ns/op -> 9562072200 ns/op`
    - `016161.xls`: `3466805000 ns/op -> 9245684500 ns/op`
  - `repeat=3` `.xls6` hotspot median total: `23111 ms -> 26823 ms`
  - hotspot `textBytes/images`: `NO_DIFF`
- decision: reverted immediately; this short-query gate made the retained `.xls` hotspot path much
  worse, so it was not worth a broader keyset rerun

- tried lowering the `tableTextContainsLine` short-query threshold from 12 runes to 8 runes
  - this lets more short cell candidates go through the existing table containment checks before they
    fall through to `visibleLineContainsLine`
  - importantly, it did not add any new containment structures or caches
- result: hotspot measurements looked good, but the wider retained keyset still regressed
  - focused regression suite: `ok officeread 11.786s`
  - hotspot benchmarks:
    - `006087.xls`: `5573310400 ns/op -> 5069138700 ns/op`
    - `008055.xls`: `7905695900 ns/op -> 5947403400 ns/op`
    - `016161.xls`: `3466805000 ns/op -> 2719632600 ns/op`
  - `repeat=3` `.xls6` hotspot median total: `23111 ms -> 21219 ms`
  - broader repeat-aware keyset total: `51517 ms -> 59328 ms`
  - keyset `.xls` subset: `17264 ms -> 24039 ms`
  - keyset `.xlsx` subset: `27830 ms -> 27026 ms`
  - keyset `.docx` subset: `6423 ms -> 8263 ms`
  - output stability: `NO_DIFF`
- decision: reverted; letting more short queries enter the table containment path helps the narrow
  `.xls` hotspot slice, but it still overfits badly against the retained wider workload

- retained: `.xls`-only escaped-table fallback guard inside `biffMarkdown(...)`
  - changed the legacy Excel markdown backfill path to skip
    `markdownBackfillVisibleText(escapeMarkdownTableCell(line))` unless the BIFF source line
    actually contains `|`
  - scoped only to `missingMarkdownTextXLS(...)`, so `.docx/.pptx/.xlsx` continue using the generic
    markdown backfill behavior unchanged
- validation:
  - focused regression suite: `ok officeread 12.475s`
  - full suite: `ok officeread 164.626s`
  - hotspot benchmarks:
    - `006087.xls`: `5573310400 ns/op -> 4984286500 ns/op`
    - `008055.xls`: `7905695900 ns/op -> 5697742600 ns/op`
    - `016161.xls`: `3466805000 ns/op -> 2939523800 ns/op`
  - `repeat=3` `.xls6` hotspot median total: `23111 ms -> 19143 ms`
  - broader repeat-aware keyset total: `51517 ms -> 47982 ms`
  - keyset `.xls` subset: `17264 ms -> 17885 ms`
  - keyset `.xlsx` subset: `27830 ms -> 23596 ms`
  - keyset `.docx` subset: `6423 ms -> 6501 ms`
  - output stability: `NO_DIFF`
- note:
  - this does not improve every `.xls` sub-metric in isolation, but it is the first recent
    `.xls`-motivated backfill change that wins on both the repeat-aware wider keyset and the hotspot
    batch without changing output

- tried a dedicated `.xls` no-image-alt backfill path on top of the retained guard
  - split `missingMarkdownTextXLS(...)` onto a separate implementation that hardcodes the
    “no images / no image alts” case
  - removed `markdownImageAltSet(nil)` construction and the per-line image-alt exclusion checks only
    for legacy Excel workbook backfill
- result: negative on the hotspot batch, so it was reverted without spending a wider keyset run
  - focused regression suite: `ok officeread 12.800s`
  - hotspot benchmarks:
    - `006087.xls`: `4984286500 ns/op -> 5168345700 ns/op`
    - `008055.xls`: `5697742600 ns/op -> 7785369900 ns/op`
    - `016161.xls`: `2939523800 ns/op -> 3448675000 ns/op`
  - `repeat=3` `.xls6` hotspot median total: `19143 ms -> 24676 ms`
  - hotspot output stability: `NO_DIFF`
- decision: reverted immediately; even this very local “no image alts” specialization added enough
  duplicated control flow to outweigh the tiny saved work

## Next work

- A retained optimization was added after the first follow-up rerun:
  `splitMarkdownTableRow(...)` now uses a byte-scan splitter instead of the earlier rune/builder
  loop. Validation evidence:
  - focused regression suite: `ok officeread 10.501s`
  - full suite: `ok officeread 219.548s`
  - same-day keyset rerun: `54578 ms -> 50489 ms`
  - same-day keyset `.xls` subset: `20668 ms -> 18325 ms`
- The detailed experiment record for that retained optimization lives in
  `testdata/web-samples/reports/performance-optimization-report.md`.
- The detailed experiment record for the retained legacy `.doc` UTF-16 hotspot optimization should be
  kept alongside the same report set.
- The detailed experiment record for the retained `.pptx` OOXML lookup optimization should also be
  kept alongside the same report set.
- The detailed experiment record for the retained `.xls` markdown-backfill query-cache optimization
  should also be kept alongside the same report set.

- continue performance work on spreadsheet markdown backfill / exact-set construction, using the
  current rerun as the comparison baseline for the next retained optimization attempt
- the next focused work item should stay on `.xls`, where `006087.xls` still reruns at `5326 ms`
  and CPU time remains concentrated under
  `missingMarkdownText -> markdownBackfillContainment.visibleLineContainsLine`

## Additional experiment update

- tried a `visibleJoined` bigram prefilter in `visibleLineContainsLine(...)`
  - built a 2-byte adjacency mask for joined visible markdown text
  - used it only as a necessary-condition precheck before the existing
    `strings.Contains(c.visibleJoined, line)` call
- result: mixed hotspot bench signal, but clearly negative on repeat-aware reruns, so reverted
  - full suite: `ok officeread 150.392s`
  - hotspot benchmarks:
    - `006087.xls`: `4984286500 ns/op -> 5092775400 ns/op`
    - `008055.xls`: `5697742600 ns/op -> 5703741100 ns/op`
    - `016161.xls`: `2939523800 ns/op -> 2551241800 ns/op`
  - `repeat=3` `.xls6` hotspot median total: `19143 ms -> 21792 ms`
  - broader repeat-aware keyset total: `47982 ms -> 57913 ms`
  - keyset `.docx` subset: `6501 ms -> 8096 ms`
  - keyset `.xls` subset: `17885 ms -> 22350 ms`
  - keyset `.xlsx` subset: `23596 ms -> 27467 ms`
  - output stability: `NO_DIFF`
- decision: reverted; the extra precheck cost more than it filtered on the retained wider workload

- tried disabling visible-line substring containment entirely for `.xls` backfill
  - kept coverage set, table containment, exact visible hits, and the retained `.xls` pipe guard
  - only removed `visibleLineContains(variant)` from the `.xls` backfill decision path
  - local shape analysis on the current retained hotspot outputs was striking:
    - `006087.xls`: `visibleSubstringCovered=0`, `dependsVisibleSubstring=0`
    - `008055.xls`: `visibleSubstringCovered=0`, `dependsVisibleSubstring=0`
- result: excellent hotspot win, but still negative on the retained wider keyset, so reverted
  - focused regression suite: `ok officeread 54.625s`
  - full suite: `ok officeread 200.171s`
  - hotspot benchmarks:
    - `006087.xls`: `4984286500 ns/op -> 1888721600 ns/op`
    - `008055.xls`: `5697742600 ns/op -> 5977595100 ns/op`
    - `016161.xls`: `2939523800 ns/op -> 2482314400 ns/op`
  - `repeat=3` `.xls6` hotspot median total: `19143 ms -> 17221 ms`
  - broader repeat-aware keyset total: `47982 ms -> 53164 ms`
  - keyset `.docx` subset: `6501 ms -> 6664 ms`
  - keyset `.xls` subset: `17885 ms -> 20558 ms`
  - keyset `.xlsx` subset: `23596 ms -> 25942 ms`
  - output stability: `NO_DIFF`
- decision: reverted; the hotspot evidence is real, but the wider retained workload still needs
  some version of that visible substring coverage path

- tried a narrower `.xls` short-digit visible-substring guard after the broader disablement branch
  - only skipped `visibleLineContains(variant)` for normalized `.xls` lines that were all ASCII
    digits and 4-7 bytes long
  - this was based on the refined keyset analysis:
    - `008055.xls`, `013623.xls`, `016161.xls`, `018548.xls`, `019088.xls`, `019089.xls`, and
      `002505.xls` showed no observed dependency on `visibleLineContainsLine(...)`
    - the remaining `006087.xls` dependent examples were short numeric substrings like `61047`,
      `30357`, `21595`, `123311`, which matched longer visible numeric cells such as `610479`,
      `303573`, `215950`, `1233116`
- result: still bad in repeat-aware reruns, so reverted
  - focused regression suite: `ok officeread 46.377s`
  - full suite: `ok officeread 198.504s`
  - hotspot benchmarks:
    - `006087.xls`: `4984286500 ns/op -> 2067601000 ns/op`
    - `008055.xls`: `5697742600 ns/op -> 8132269000 ns/op`
    - `016161.xls`: `2939523800 ns/op -> 3584242600 ns/op`
  - same-day rerun baseline `.xls6` median total: `19100 ms`
  - experiment `.xls6` median total: `22356 ms`
  - same-day rerun baseline keyset total: `51865 ms`
  - experiment keyset total: `66333 ms`
  - output stability: `NO_DIFF`
- decision: reverted; even this very targeted short-digit rule overfit the `006087.xls` shape and
  fell apart on the retained repeat-aware workload

- tried a short-digit exact-set plus visible-substring guard for `.xls`
  - added an exact set only for visible table cells that were all ASCII digits and 4-7 bytes long
  - used that exact set before the `< 12 runes` early return in `tableTextContainsLine(...)`
  - only skipped `.xls` `visibleLineContains(...)` after a short numeric line failed that exact-hit
    check
- result: correctness looked nicer, but performance still regressed badly, so reverted
  - focused regression suite: `ok officeread 13.687s`
  - full suite: `ok officeread 172.405s`
  - hotspot benchmarks:
    - `006087.xls`: `4984286500 ns/op -> 2190423400 ns/op`
    - `008055.xls`: `5697742600 ns/op -> 8050013700 ns/op`
    - `016161.xls`: `2939523800 ns/op -> 3451519400 ns/op`
  - same-day rerun baseline `.xls6` median total: `19100 ms`
  - experiment `.xls6` median total: `22821 ms`
  - same-day rerun baseline keyset total: `51865 ms`
  - experiment keyset total: `65921 ms`
  - output stability: `NO_DIFF`
- decision: reverted; rescuing exact short numeric cells was not enough to make the short-digit
  guard economically viable

- tried feeding the `.xls` markdown-backfill path with direct `biffText(data)` lines
  - removed the `strings.Join(biffText(data), "\n") -> missingMarkdownTextXLS(...) ->
    markdownBackfillRawLines(...)` round-trip
  - passed BIFF text straight into a dedicated line-based helper so `.xls` could skip the
    join-and-split conversion before backfill matching
- result: output stayed stable, but performance regressed clearly enough to reject the idea
  - full suite: `ok officeread 139.279s`
  - hotspot benchmarks:
    - `006087.xls`: `4984286500 ns/op -> 5405430900 ns/op`
    - `008055.xls`: `5697742600 ns/op -> 8471192500 ns/op`
    - `016161.xls`: `2939523800 ns/op -> 3449793300 ns/op`
  - same-day rerun baseline `.xls6` median total: `19100 ms`
  - experiment `.xls6` median total: `22473 ms`
  - same-day rerun baseline keyset `.xls` subset: `17801 ms`
  - experiment keyset `.xls` subset: `20832 ms`
  - output stability: `NO_DIFF`
- decision: reverted; saving the join/split work was not enough to offset the extra dedicated
  `.xls` control flow and the retained baseline remains faster

## Additional retained optimization update

- retained a shared-workbook-text optimization for normal `.xls` extraction
  - normal legacy Excel extraction was rescanning the same workbook bytes for:
    - `Text` generation via `extractLegacyTextWithMetadata(...)`
    - `StructuredMarkdown` generation via `biffMarkdown(...)`
    - and then another `biffText(...)` call inside markdown backfill
  - the new path computes workbook BIFF text once in `extractLegacyWithDepth(...)`, reuses it to
    assemble `.xls` text output, and passes the same text slice into `biffMarkdownWithText(...)`
    so markdown backfill can reuse it without another workbook scan
- validation:
  - focused regression suite: `ok officeread 12.319s`
  - hotspot benchmarks:
    - `006087.xls`: `4984286500 ns/op -> 5062447200 ns/op`
    - `008055.xls`: `5697742600 ns/op -> 6485531800 ns/op`
    - `016161.xls`: `2939523800 ns/op -> 2645504200 ns/op`
  - repeat-aware `.xls6` hotspot batch:
    - same-day rerun baseline total: `19100 ms`
    - experiment total: `16327 ms`
  - broader repeat-aware keyset:
    - same-day rerun baseline total: `51865 ms`
    - experiment total: `49269 ms`
    - same-day rerun baseline `.xls` subset: `17801 ms`
    - experiment `.xls` subset: `15657 ms`
    - same-day rerun baseline `.docx` subset: `6407 ms`
    - experiment `.docx` subset: `6360 ms`
    - same-day rerun baseline `.xlsx` subset: `27657 ms`
    - experiment `.xlsx` subset: `27252 ms`
  - output stability: `NO_DIFF`
  - full suite: `ok officeread 160.342s`
- interpretation:
  - this is the first recent `.xls` optimization attempt after the backfill-query-cache work that
    wins on both the repeat-aware hotspot batch and the broader mixed keyset without changing
    output
  - unlike the rejected query-side branches, it removes repeated workbook scanning upstream of the
    markdown-backfill decision path, which matches the current evidence better

- retained a two-stage markdown-backfill precompute optimization
  - `missingMarkdownTextWithOptions(...)` used to build `exact`, `coverage`, and `containment`
    structures through separate markdown scans
  - the earlier shared-precompute variant collapsed everything into one shared build, but the final
    retained version performed better by splitting the work into two stages:
    - build `exact` first for the existing early-return check
    - only if that check fails, build the shared `coverage + containment` structures
- validation:
  - focused regression suite: `ok officeread 10.896s`
  - hotspot benchmarks:
    - `006087.xls`: `5062447200 ns/op -> 4817586900 ns/op`
    - `008055.xls`: `6485531800 ns/op -> 4644516300 ns/op`
    - `016161.xls`: `2645504200 ns/op -> 1938585400 ns/op`
  - repeat-aware `.xls6` hotspot batch:
    - same-day rerun baseline total: `19100 ms`
    - experiment total: `15740 ms`
  - broader repeat-aware keyset:
    - same-day rerun baseline total: `51865 ms`
    - experiment total: `46830 ms`
    - same-day rerun baseline `.xls` subset: `17801 ms`
    - experiment `.xls` subset: `14754 ms`
    - same-day rerun baseline `.docx` subset: `6407 ms`
    - experiment `.docx` subset: `6608 ms`
    - same-day rerun baseline `.xlsx` subset: `27657 ms`
    - experiment `.xlsx` subset: `25468 ms`
  - output stability: `NO_DIFF`
  - full suite: `ok officeread 137.914s`
- interpretation:
  - the final two-stage version wins more clearly on `.xls` itself than the earlier all-at-once
    shared-build variant
  - the current `006087.xls` profile still spends most time in
    `markdownBackfillContainment.visibleLineContainsLine(...)`, but the separate
    `exact` and `coverage/containment` preprocessing now avoid the old three-pass setup while still
    preserving the existing exact-check early exit

- tried disabling the `.xls` backfill query-result caches
  - local replay analysis on `006087.xls`, `008055.xls`, and `016161.xls` showed 0% hit rate for
    `coverageCache`, `tableContainsCache`, and `visibleContainsCache`
  - based on that, disabled those three caches only on the `.xls` path and kept the generic path
    unchanged
- result: looked superficially promising on whole-keyset wall time, but the affected `.xls` subset
  still regressed against the same-day baseline, so reverted
  - full suite: `ok officeread 191.013s`
  - hotspot benchmarks:
    - `006087.xls`: `4984286500 ns/op -> 5172676100 ns/op`
    - `008055.xls`: `5697742600 ns/op -> 8398555800 ns/op`
    - `016161.xls`: `2939523800 ns/op -> 2790799400 ns/op`
  - same-day rerun baseline `.xls6` total: `19100 ms`
  - experiment `.xls6` total: `19576 ms`
  - same-day rerun baseline keyset `.xls` subset: `17801 ms`
  - experiment keyset `.xls` subset: `18153 ms`
  - output stability: `NO_DIFF`
- decision: reverted; even a 0% measured hit rate was not enough to justify removing those caches
  on the retained `.xls` workloads

- tried lazy containment construction for `.xls` backfill
  - local replay showed `008055.xls`, `016161.xls`, and `002505.xls` were fully satisfied by
    `coverageContains(...)`, while only `006087.xls` actually touched `visibleLineContains(...)`
  - based on that, delayed `markdownBackfillContainmentSet(markdown)` on the `.xls` path until the
    first real `table/visible contains` query
- result: clearly slower, so reverted
  - full suite: `ok officeread 183.801s`
  - hotspot benchmarks:
    - `006087.xls`: `4984286500 ns/op -> 5184848800 ns/op`
    - `008055.xls`: `5697742600 ns/op -> 8462841300 ns/op`
    - `016161.xls`: `2939523800 ns/op -> 3564291300 ns/op`
  - same-day rerun baseline `.xls6` total: `19100 ms`
  - experiment `.xls6` total: `26977 ms`
  - same-day rerun baseline keyset `.xls` subset: `17801 ms`
  - experiment keyset `.xls` subset: `29299 ms`
  - output stability: `NO_DIFF`
- decision: reverted; avoiding the upfront containment build cost looked elegant, but in practice it
  made the retained `.xls` workload much worse

- tried an `.xls` exact-coverage fast path before variant assembly
  - local replay showed the retained `.xls` hotspot samples were overwhelmingly solved by direct
    `coverage[line]` exact hits, with no observed `visible/comparable/alternate` coverage hits in
    `006087.xls`, `008055.xls`, `016161.xls`, or `002505.xls`
  - used that to add a very narrow `.xls`-only early return before building the variant list
- result: still clearly slower, so reverted
  - full suite: `ok officeread 175.404s`
  - hotspot benchmarks:
    - `006087.xls`: `4984286500 ns/op -> 5161550900 ns/op`
    - `008055.xls`: `5697742600 ns/op -> 8229030100 ns/op`
    - `016161.xls`: `2939523800 ns/op -> 3513155100 ns/op`
  - same-day rerun baseline `.xls6` total: `19100 ms`
  - experiment `.xls6` total: `24716 ms`
  - same-day rerun baseline keyset `.xls` subset: `17801 ms`
  - experiment keyset `.xls` subset: `24307 ms`
  - output stability: `NO_DIFF`
- decision: reverted; even the dominant exact-hit shape was not enough to make this extra early
  branch pay for itself

- tried an even narrower `.xls` direct `coverage[line]` exact-hit fast path
  - fresh replay showed the retained `.xls` hotspots were almost entirely resolved by exact
    `coverage[line]` hits, with no observed `visible/comparable/alternate` coverage hits on
    `006087.xls`, `008055.xls`, `016161.xls`, or `002505.xls`
  - based on that, added a tiny `.xls`-only early return before variant assembly when
    `coverage[line]` already matched exactly
- result: still clearly slower, so reverted
  - full suite: `ok officeread 175.404s`
  - hotspot benchmarks:
    - `006087.xls`: `4984286500 ns/op -> 5161550900 ns/op`
    - `008055.xls`: `5697742600 ns/op -> 8229030100 ns/op`
    - `016161.xls`: `2939523800 ns/op -> 3513155100 ns/op`
  - same-day rerun baseline `.xls6` total: `19100 ms`
  - experiment `.xls6` total: `24716 ms`
  - same-day rerun baseline keyset `.xls` subset: `17801 ms`
  - experiment keyset `.xls` subset: `24307 ms`
  - output stability: `NO_DIFF`
- decision: reverted; even the “obvious” direct exact-hit branch still lost on repeat-aware data

## 2026-07-05 continued: isolate `.xls` markdown-backfill cost before the next runtime experiment

- no new runtime optimization was retained in this pass
- instead, added a narrower benchmark to separate the `.xls` markdown-backfill hotspot from the
  rest of `Extract(...)`:
  - code change: [extract_bench_test.go](D:/workprj/officeread/extract_bench_test.go)
  - new benchmark: `BenchmarkXLSMarkdownBackfillHotspots`
- rationale:
  - fresh `006087.xls` profile still shows the dominant cost in
    `markdownBackfillContainment.visibleLineContainsLine(...)`
  - whole-file `BenchmarkExtractXLSHotspots` is still useful, but it mixes BIFF parsing, markdown
    rendering, and backfill work together
  - isolating `missingMarkdownTextXLS(res.StructuredMarkdown, res.Text)` gives a faster and more
    targeted loop for the next containment experiments

Fresh evidence gathered in this pass:
- focused single-file profile refresh:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/profileextract -out testdata\web-samples\reports\profile-006087-xls-current-turn.pprof testdata\web-samples\samples\xls\006087.xls`
  - observed:
    - `006087.xls`: `5085 ms`
    - `pprof -top` still dominated by `strings.Contains(...)` under
      `markdownBackfillContainment.visibleLineContainsLine(...)`
    - current focused hotspot shape:
      - `markdownBackfillContainment.visibleLineContainsLine(...)`: `3.17s`
      - `markdownBackfillBuildCoverageContainment(...)`: `0.63s`
      - `markdownBackfillExactSet(...)`: `0.25s`
- current whole-file hotspot benchmark refresh:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSHotspots|BenchmarkXLSMarkdownBackfillHotspots' -benchmem -benchtime=1x ./`
  - observed:
    - whole extract:
      - `006087.xls`: `4963697200 ns/op`, `617333064 B/op`, `5130493 allocs/op`
      - `008055.xls`: `6463587000 ns/op`, `1587257536 B/op`, `4587163 allocs/op`
      - `016161.xls`: `2898870500 ns/op`, `722625640 B/op`, `2623389 allocs/op`
    - isolated `.xls` markdown backfill:
      - `006087.xls`: `1508848200 ns/op`, `331088128 B/op`, `3453696 allocs/op`
      - `008055.xls`: `3269654600 ns/op`, `573343336 B/op`, `1532727 allocs/op`
      - `016161.xls`: `1398286500 ns/op`, `262641464 B/op`, `968709 allocs/op`

Additional shape analysis from `006087.xls`:
- source text size:
  - `41347` non-empty lines
  - average normalized line length about `6.02`
  - most common line length: `6` characters (`31301` lines)
- current markdown visible-line approximation:
  - `54601` visible lines
  - average visible-line length about `83.02`
  - `49169` visible lines are still at least `64` bytes long
- takeaway:
  - the remaining hotspot is still “many very short queries against very large joined visible-line
    haystacks”, not “a few very long queries”
  - that makes broad runtime changes risky: earlier short-query / bucket experiments improved
    `006087.xls` alone but regressed the wider retained `.xls` workload

Decision:
- retain the new isolated benchmark
- keep the current exact-first / shared-rest runtime baseline unchanged for now
- use the narrower benchmark plus the fresh profile shape to evaluate the next containment idea
  before touching `extract.go` again

## 2026-07-05 retained: consult table exact hits before the short-line substring cutoff

- runtime change in [extract.go](D:/workprj/officeread/extract.go):
  - in `markdownBackfillContainment.tableTextContainsLine(...)`, move the existing
    `tableRawExact` / `tableVisibleExact` checks before the `< 12 runes` early return
- motivation:
  - fresh local diagnosis on `006087.xls` showed that the expensive
    `visibleLineContainsLine(...)` path was only being exercised by short `4-6` byte queries
  - every observed visible-line hit in that slice was digit-only
  - the existing containment builder already had exact table-line maps for raw and visible table
    text, but the old short-line cutoff prevented those exact hits from being used for short cell
    values
- scope:
  - no matching rule was removed
  - this only changes the order of existing exact checks versus the old short-line early return
  - substring matching for longer lines and comparable-table text stays unchanged

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- isolated `.xls` markdown-backfill benchmark:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkXLSMarkdownBackfillHotspots -benchmem -benchtime=1x ./`
- whole extract hotspot benchmark:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench BenchmarkExtractXLSHotspots -benchmem -benchtime=1x ./`
- repeat-aware `.xls6` rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-short-exact-before-runecheck-xls6.json -csv testdata\web-samples\reports\perf-exp-short-exact-before-runecheck-xls6.csv ...`
- repeat-aware mixed keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-short-exact-before-runecheck-keyset.json -csv testdata\web-samples\reports\perf-exp-short-exact-before-runecheck-keyset.csv ...`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused regression suite:
  - `ok officeread 16.610s`
- isolated `.xls` markdown-backfill benchmark:
  - `006087.xls`: `1508848200 ns/op -> 1085084500 ns/op`
  - `008055.xls`: `3269654600 ns/op -> 2093789500 ns/op`
  - `016161.xls`: `1398286500 ns/op -> 1105105100 ns/op`
- whole extract hotspot benchmark:
  - previous retained baseline:
    - `006087.xls`: `4963697200 ns/op`
    - `008055.xls`: `6463587000 ns/op`
    - `016161.xls`: `2898870500 ns/op`
  - after this change:
    - `006087.xls`: `4695615800 ns/op`
    - `008055.xls`: `4719727300 ns/op`
    - `016161.xls`: `1879915300 ns/op`
- repeat-aware `.xls6` hotspot batch:
  - same-day rerun baseline total: `19100 ms`
  - experiment total: `16482 ms`
  - per-file medians:
    - `002505.xls`: `2268 ms -> 1805 ms`
    - `006087.xls`: `4946 ms -> 4948 ms`
    - `008055.xls`: `5734 ms -> 4722 ms`
    - `016161.xls`: `2451 ms -> 2046 ms`
    - `019088.xls`: `1737 ms -> 1464 ms`
    - `019089.xls`: `1964 ms -> 1497 ms`
- repeat-aware mixed keyset:
  - same-day rerun baseline total: `51865 ms`
  - experiment total: `45205 ms`
  - same-day rerun baseline `.xls` subset: `17801 ms`
  - experiment `.xls` subset: `14646 ms`
  - same-day rerun baseline `.docx` subset: `6407 ms`
  - experiment `.docx` subset: `6829 ms`
  - same-day rerun baseline `.xlsx` subset: `27657 ms`
  - experiment `.xlsx` subset: `23730 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: `NO_DIFF` on both `.xls6` and mixed keyset
- full repository regression:
  - `ok officeread 140.389s`

Decision:
- Retained. This is a small ordering change, but it materially improves the repeat-aware `.xls`
  workload and also helps the broader mixed keyset overall, while preserving output and keeping the
  full repository regression clean.

## 2026-07-05 retained refinement: narrow the short table exact fast path to `.xls` only

- follow-up change in [extract.go](D:/workprj/officeread/extract.go):
  - keep the short table exact-before-cutoff behavior only on the `.xls` path
  - restore the previous generic behavior for `missingMarkdownText(...)` used by `.docx` / `.xlsx`
- rationale:
  - the prior all-format version improved totals, but it also nudged the mixed-keyset `.docx`
    subset from `6407 ms` to `6829 ms`
  - the original evidence for the optimization came from `.xls`-specific short numeric table-cell
    hits, so the broader scope was unnecessary

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- repeat-aware `.xls6` rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-short-exact-xls-only-xls6.json -csv testdata\web-samples\reports\perf-exp-short-exact-xls-only-xls6.csv ...`
- repeat-aware mixed keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json -csv testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.csv ...`
- focused regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownBackfillTreats.*|TestMarkdownTableCell.*|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- repeat-aware `.xls6` hotspot batch:
  - same-day rerun baseline total: `19100 ms`
  - prior all-format version: `16482 ms`
  - `.xls`-only refinement: `16118 ms`
- repeat-aware mixed keyset:
  - same-day rerun baseline total: `51865 ms`
  - prior all-format version: `45205 ms`
  - `.xls`-only refinement: `44994 ms`
  - by subset:
    - `.docx`: `6407 ms -> 6829 ms -> 6433 ms`
    - `.xls`: `17801 ms -> 14646 ms -> 14471 ms`
    - `.xlsx`: `27657 ms -> 23730 ms -> 24090 ms`
  - per-file `.docx` medians:
    - `00003763.docx`: `3855 ms -> 3944 ms`
    - `223624.docx`: `2552 ms -> 2489 ms`
- output stability:
  - retained-vs-baseline `textBytes` / `images`: `NO_DIFF` on both `.xls6` and mixed keyset
- focused regression suite:
  - `ok officeread 11.220s`
- full repository regression:
  - `ok officeread 152.104s`

Decision:
- Retained, and this supersedes the broader all-format version from the previous step.
- The `.xls`-only scope keeps essentially all of the intended `.xls` gain, improves the mixed-keyset
  total a bit further, and removes nearly all of the `.docx` regression introduced by the broader
  variant.

## 2026-07-05 rejected: one-pass tag scan for `simpleInlineWorksheetCandidate(...)`

- explored the current `.xlsx` hotspot after the `.xls` baseline had stabilized
- fresh current profiles showed:
  - `testRecordSizeExceeded.xlsx`: `4358 ms`, dominated by
    `appendSimpleInlineWorksheetText(...)`
  - within that path, `simpleInlineWorksheetCandidate(...)` was still spending visible time in many
    repeated `bytes.Contains(...)` / `bytes.Index(...)` scans over large worksheet XML payloads
- experiment:
  - rewrote `simpleInlineWorksheetCandidate(...)` from repeated whole-buffer marker scans into a
    single pass over XML start tags
  - kept the same broad disqualifier intent: formulas, `<v>`, shared/bool/str cells, hidden
    markers, hyperlinks, data validations, header/footer, and phonetic markers still forced the
    fast path off

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- focused `.xlsx` pair replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-simple-inline-candidate-xlsx-pair.json -csv testdata\web-samples\reports\perf-exp-simple-inline-candidate-xlsx-pair.csv testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx`
- post-revert regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused `.xlsx` regression slice:
  - `ok officeread 8.482s`
- focused `.xlsx` pair replay against the current retained keyset baseline:
  - `00012389.xlsx`: `2460 ms -> 2735 ms`
  - `testRecordSizeExceeded.xlsx`: `3967 ms -> 4040 ms`
- output stability:
  - sample outputs remained functionally stable, but the runtime direction was clearly negative
- post-revert regression:
  - `ok officeread 5.336s`
  - `ok officeread (cached)` on full suite

Decision:
- Reverted. Replacing many coarse `bytes.Contains(...)` passes with one tag scan looked attractive
  in profile shape, but on the retained `.xlsx` hotspots it still ran slower overall.

## 2026-07-05 rejected: single-`<t>` fast path in `simpleInlineCellText(...)`

- retained bench infrastructure improvement:
  - added `.xlsx` hotspot benchmarks in [extract_bench_test.go](D:/workprj/officeread/extract_bench_test.go)
  - new benchmarks:
    - `BenchmarkExtractXLSXHotspots`
    - `BenchmarkXLSXSimpleInlineHotspots`
    - `BenchmarkXLSXWorksheetTextHotspots`
- fresh baseline from those new benches showed:
  - `testRecordSizeExceeded.xlsx`
    - whole extract: `4640854300 ns/op`
    - simple inline fast path: `4190777000 ns/op`
    - worksheet text path: `4650526500 ns/op`
  - `00012389.xlsx`
    - whole extract: `2965216100 ns/op`
    - worksheet text path: `2066764500 ns/op`
- experiment:
  - added a narrow fast path in `simpleInlineCellText(...)`:
    - if a cell had only one `<t>...</t>` segment, return it directly
    - only fall back to the existing builder loop for multi-segment rich text
  - motivation:
    - on the fresh `.xlsx` profile, `simpleInlineCellText(...)` still showed up as a per-cell
      cost inside `appendSimpleInlineWorksheetText(...)`

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go extract_bench_test.go`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- `.xlsx` hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- focused `.xlsx` pair replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-simple-inline-celltext-xlsx-pair.json -csv testdata\web-samples\reports\perf-exp-simple-inline-celltext-xlsx-pair.csv testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx`
- mixed keyset replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-simple-inline-celltext-keyset.json -csv testdata\web-samples\reports\perf-exp-simple-inline-celltext-keyset.csv ...`
- post-revert regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused `.xlsx` regression slice:
  - `ok officeread 8.494s`
- hotspot benchmarks looked promising:
  - `testRecordSizeExceeded.xlsx`
    - whole extract: `4640854300 ns/op -> 4570901400 ns/op`
    - simple inline: `4190777000 ns/op -> 3958273900 ns/op`
    - worksheet text: `4650526500 ns/op -> 3879766600 ns/op`
  - `00012389.xlsx`
    - worksheet text: `2066764500 ns/op -> 1470050600 ns/op`
- output stability:
  - mixed keyset `textBytes` / `images`: `NO_DIFF`
- but repeat-aware replays clearly rejected it:
  - focused `.xlsx` pair against the current retained keyset baseline:
    - `00012389.xlsx`: `2460 ms -> 2763 ms`
    - `testRecordSizeExceeded.xlsx`: `3967 ms -> 4761 ms`
  - mixed keyset total:
    - current retained baseline: `44994 ms`
    - experiment: `51552 ms`
  - subset totals:
    - `.docx`: `6433 ms -> 7819 ms`
    - `.xls`: `14471 ms -> 16127 ms`
    - `.xlsx`: `24090 ms -> 27606 ms`
- post-revert regression:
  - focused `.xlsx` slice: `ok officeread 7.784s`
  - full suite: `ok officeread 134.471s`

Decision:
- Reverted. This one is a good reminder that a tight benchmark win can still lose badly on the
  retained repeat-aware workload once GC and broader runtime interaction enter the picture.

## 2026-07-05 rejected: byte-based cell reference parsing in simple inline `.xlsx` path

- after adding the new `.xlsx` hotspot benches, tried another low-risk-looking allocation
  experiment in `appendSimpleInlineWorksheetText(...)`
- experiment:
  - avoid `string(cellRef)` allocation for the `r=` attribute on each inlineStr cell
  - add a byte-oriented `cellRefIndexesBytes(...)`
  - use it only inside the simple inline `.xlsx` path
- motivation:
  - this looked like a tidy way to shave per-cell allocation without changing text cleaning or
    worksheet semantics

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- `.xlsx` hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- focused `.xlsx` pair replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-cellref-bytes-xlsx-pair.json -csv testdata\web-samples\reports\perf-exp-cellref-bytes-xlsx-pair.csv testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx`
- mixed keyset replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-cellref-bytes-keyset.json -csv testdata\web-samples\reports\perf-exp-cellref-bytes-keyset.csv ...`
- post-revert regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused `.xlsx` regression slice:
  - `ok officeread 7.869s`
- narrow signals were mixed:
  - `.xlsx` pair:
    - `00012389.xlsx`: `2460 ms -> 2436 ms`
    - `testRecordSizeExceeded.xlsx`: `3967 ms -> 4015 ms`
  - hotspot benchmarks:
    - `testRecordSizeExceeded.xlsx` whole extract: `4640854300 ns/op -> 4537746500 ns/op`
    - `00012389.xlsx` whole extract: `2965216100 ns/op -> 3293568900 ns/op`
- output stability:
  - mixed keyset `textBytes` / `images`: `NO_DIFF`
- wider validation clearly rejected it:
  - mixed keyset total:
    - current retained baseline: `44994 ms`
    - experiment: `56450 ms`
  - subset totals:
    - `.docx`: `6433 ms -> 8028 ms`
    - `.xls`: `14471 ms -> 17589 ms`
    - `.xlsx`: `24090 ms -> 30833 ms`
- post-revert regression:
  - focused `.xlsx` slice: `ok officeread 7.728s`
  - full suite: `ok officeread (cached)`

Decision:
- Reverted. Another case where a seemingly careful micro-allocation optimization did not survive
  repeat-aware mixed-workload validation.

## 2026-07-05 rejected: stop simple-inline row/col tracking after Markdown row cap

- used the new `.xlsx` hotspot benches to split the simple-inline path one layer deeper:
  - `BenchmarkXLSXSimpleInlineHotspots`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots`
  - `BenchmarkXLSXWorksheetTextHotspots`
- one concrete new measurement from that split:
  - on `testRecordSizeExceeded.xlsx`
    - simple-inline with Markdown collection: `3764269900 ns/op`
    - simple-inline text-only (`md=nil`): `3526142900 ns/op`
  - so the Markdown side work in this path is real, but smaller than the full text extraction cost
- experiment:
  - after `md.rows` reaches the existing `maxMarkdownTableRows` cap, stop parsing `r=` cell refs
    and stop row/column bookkeeping for the remainder of the worksheet
  - text output would continue exactly as before; only the already-capped Markdown collector would
    do less tail work

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- `.xlsx` hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots|BenchmarkXLSXWorksheetTextHotspots|BenchmarkExtractXLSXHotspots' -benchmem -benchtime=1x ./`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- focused `.xlsx` pair replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-inline-markdown-stop-after-cap-xlsx-pair.json -csv testdata\web-samples\reports\perf-exp-inline-markdown-stop-after-cap-xlsx-pair.csv testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx`
- post-revert regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`

Observed results:
- benchmark view was noisy and not convincing:
  - `testRecordSizeExceeded.xlsx`
    - whole extract: `4640854300 ns/op -> 4576078000 ns/op`
    - simple inline: `3764269900 ns/op -> 4184935800 ns/op`
    - text-only simple inline: `3526142900 ns/op -> 3166360900 ns/op`
  - `00012389.xlsx`
    - worksheet text: `1915889800 ns/op -> 1594668300 ns/op`
- focused `.xlsx` regression slice:
  - `ok officeread 10.395s`
- focused `.xlsx` pair replay against the current retained baseline:
  - `00012389.xlsx`: `2460 ms -> 2574 ms`
  - `testRecordSizeExceeded.xlsx`: `3967 ms -> 4157 ms`
- post-revert regression:
  - `ok officeread 5.513s`

Decision:
- Reverted. The idea was semantically conservative, but the real `.xlsx` pair still slowed down, so
  it does not earn a place in the retained baseline.

## 2026-07-05 retained benchmark refinement: isolate prepared simple-inline `.xlsx` cost

- benchmark-only code change:
  - in [extract.go](D:/workprj/officeread/extract.go), split the body of
    `appendSimpleInlineWorksheetText(...)` into a helper
    `appendSimpleInlineWorksheetTextPrepared(...)`
  - in [extract_bench_test.go](D:/workprj/officeread/extract_bench_test.go), move
    `simpleInlineWorksheetCandidate(...)` validation outside the timed loop and benchmark the prepared
    helper directly
- rationale:
  - the earlier simple-inline `.xlsx` microbench was still charging the candidate gate cost to the
    extraction loop, which distorted hotspot attribution

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go extract_bench_test.go`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- corrected `.xlsx` hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- text-only simple-inline profile:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineTextOnlyHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchtime=1x -cpuprofile tmp-shape\xlsx-simpleinline-textonly-prepared.prof ./`
  - `& 'C:\Program Files\Go\bin\go.exe' tool pprof -top .\officeread.test.exe tmp-shape\xlsx-simpleinline-textonly-prepared.prof`

Observed results:
- focused `.xlsx` regression slice:
  - `ok officeread 7.866s`
- corrected benchmark view:
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2890754000 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `4761816100 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `2126471600 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1922540000 ns/op`
- corrected profile takeaway:
  - the prepared simple-inline text-only profile no longer supports optimizing
    `simpleInlineWorksheetCandidate(...)` first
  - the remaining timed cost is concentrated in per-cell text extraction / cleaning work such as
    `appendWorksheetValue(...)`, `cleanText(...)`, and `simpleInlineCellText(...)`

Decision:
- Retained. This does not change runtime behavior, but it gives a more trustworthy `.xlsx`
  measurement harness for future optimization turns.

## 2026-07-05 rejected: skip the second trim inside `appendWorksheetValue(...)`

- experiment:
  - after `cleanText(...)`, write worksheet values with `appendTrimmedTextBlock(...)` directly
    instead of routing through `appendCleanedTextBlock(...)`
- rationale:
  - `cleanText(...)` already returns a trimmed string, so this looked like a small redundant-work
    removal on the heavy `.xlsx` / worksheet path

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- repeat-aware `.xlsx` pair replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-ai-assistant-worksheetvalue-xlsx-pair.json -csv testdata\web-samples\reports\perf-exp-ai-assistant-worksheetvalue-xlsx-pair.csv testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx`
- repeat-aware mixed keyset replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-ai-assistant-worksheetvalue-keyset.json -csv testdata\web-samples\reports\perf-exp-ai-assistant-worksheetvalue-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`

Observed results:
- focused `.xlsx` regression slice:
  - `ok officeread 8.280s`
- hotspot benchmarks improved on the `.xlsx` samples:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `4827813500 ns/op -> 4779573600 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `3350374500 ns/op -> 2911278000 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2890754000 ns/op -> 2640816000 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `4761816100 ns/op -> 4543836300 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `2126471600 ns/op -> 2017237700 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1922540000 ns/op -> 1721084500 ns/op`
- repeat-aware `.xlsx` pair:
  - `00012389.xlsx`: `2460 ms -> 2535 ms`
  - `testRecordSizeExceeded.xlsx`: `3967 ms -> 5051 ms`
- repeat-aware mixed keyset:
  - retained baseline total: `44994 ms`
  - experiment total: `45549 ms`
  - by subset:
    - `.docx`: `6433 ms -> 7994 ms`
    - `.xls`: `14471 ms -> 14020 ms`
    - `.xlsx`: `24090 ms -> 23535 ms`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: `NO_DIFF`

Decision:
- Reverted. The change helped `.xls` and `.xlsx`, but it regressed `.docx` enough to lose on the
  retained mixed workload overall.

## 2026-07-05 rejected: gate expensive control-fragment helpers behind narrower shape checks

- experiment:
  - in `maybeControlFragmentText(...)`, try the large first-character literal switch first
  - only call helper recognizers such as `isLegacyObjectReference(...)`,
    `looksLikeOLEIdentifierFragment(...)`, `looksLikeOLEWrapperStreamName(...)`, and
    `looksLikeOOXMLMarkupNameFragment(...)` when the string shape looked more plausible
- rationale:
  - the corrected `.xlsx` profile still showed visible cost under
    `cleanTextFastPath -> maybeControlFragmentText`, so a conservative gating pass looked like a good
    candidate

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- repeat-aware `.xlsx` pair replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-control-fragment-gating-xlsx-pair.json -csv testdata\web-samples\reports\perf-exp-control-fragment-gating-xlsx-pair.csv testdata\web-samples\samples\xlsx\apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata\web-samples\samples\xlsx\00012389.xlsx`
- repeat-aware mixed keyset replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-control-fragment-gating-keyset.json -csv testdata\web-samples\reports\perf-exp-control-fragment-gating-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`
- post-revert regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused `.xlsx` regression slice:
  - `ok officeread 7.680s`
- hotspot benchmarks looked excellent:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `4827813500 ns/op -> 3867231400 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `3350374500 ns/op -> 2212112400 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2890754000 ns/op -> 1832484100 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `4761816100 ns/op -> 3450764400 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `2126471600 ns/op -> 1289478700 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1922540000 ns/op -> 1080451600 ns/op`
- repeat-aware `.xlsx` pair also improved:
  - `00012389.xlsx`: `2460 ms -> 2256 ms`
  - `testRecordSizeExceeded.xlsx`: `3967 ms -> 3539 ms`
- repeat-aware mixed keyset improved overall:
  - retained baseline total: `44994 ms`
  - experiment total: `43569 ms`
  - by subset:
    - `.docx`: `6433 ms -> 6362 ms`
    - `.xls`: `14471 ms -> 13009 ms`
    - `.xlsx`: `24090 ms -> 24198 ms`
- but output stability failed:
  - `00010400.xlsx` changed from `textBytes=2778014` to `2778070`

Decision:
- Reverted. Even with very strong performance numbers, this experiment cannot be retained because it
  changed extraction output on a retained keyset sample.

## 2026-07-05 rejected: keep OOXML markup filtering early, but gate the other control-fragment helpers

- follow-up experiment:
  - keep `looksLikeOOXMLMarkupNameFragment(...)` at the top of `maybeControlFragmentText(...)`
    so strings like `relationships` remain filtered
  - still move the other helper recognizers behind narrower shape checks:
    `isLegacyObjectReference(...)`, `looksLikeOLEIdentifierFragment(...)`, and
    `looksLikeOLEWrapperStreamName(...)`
- rationale:
  - the previous broad gating experiment proved the direction had real speed upside
  - the output drift was specifically caused by delayed OOXML markup-name filtering, so this version
    tried to preserve that safety while keeping most of the gain

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- repeat-aware mixed keyset rerun #1:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-control-fragment-gating-safe-keyset.json -csv testdata\web-samples\reports\perf-exp-control-fragment-gating-safe-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`
- repeat-aware mixed keyset rerun #2:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-control-fragment-gating-safe-keyset-rerun2.json -csv testdata\web-samples\reports\perf-exp-control-fragment-gating-safe-keyset-rerun2.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`
- repeat-aware `.docx` pair replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata\web-samples\reports\perf-exp-control-fragment-gating-safe-docx-pair.json -csv testdata\web-samples\reports\perf-exp-control-fragment-gating-safe-docx-pair.csv testdata\web-samples\samples\docx\00003763.docx testdata\web-samples\samples\docx\223624.docx`
- hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- output stability:
  - both mixed-keyset reruns preserved `textBytes` / `images`: `NO_DIFF`
- first repeat-aware mixed keyset rerun looked barely positive:
  - retained baseline total: `44994 ms`
  - experiment total: `44816 ms`
  - by subset:
    - `.docx`: `6433 ms -> 8006 ms`
    - `.xls`: `14471 ms -> 13474 ms`
    - `.xlsx`: `24090 ms -> 23336 ms`
- second repeat-aware mixed keyset rerun went the other way:
  - retained baseline total: `44994 ms`
  - experiment total: `46322 ms`
  - by subset:
    - `.docx`: `6433 ms -> 8657 ms`
    - `.xls`: `14471 ms -> 13395 ms`
    - `.xlsx`: `24090 ms -> 24270 ms`
- repeat-aware `.docx` pair:
  - `00003763.docx`: `3890 ms` median (`3885..4098`)
  - `223624.docx`: `2535 ms` median (`2529..2552`)
- hotspot benchmarks still looked attractive on `.xlsx`:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `4084598800 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2307492100 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1209657100 ns/op`
- full repository regression:
  - `ok officeread 131.632s`

Decision:
- Reverted. This narrower version fixed the output drift, but the repeat-aware mixed workload still
  did not show a clear, stable win across reruns, especially on the `.docx` subset.

## 2026-07-05 rejected: micro-optimize OLE identifier helper internals

- experiment:
  - keep the outer control-fragment decision flow unchanged
  - only optimize internals inside the OLE identifier helper chain:
    - `oleIdentifierAssignmentValue(...)`
    - `looksLikeGUIDString(...)`
  - replace some generic lowercase / rune-based checks with ASCII-oriented logic
- rationale:
  - the corrected `.xlsx` profile still showed measurable cost in
    `looksLikeOLEIdentifierFragment(...)` and `oleIdentifierAssignmentValue(...)`
  - this looked like a low-risk place to try a pure implementation-level optimization

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused `.xlsx` regression slice:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*XLSX|Test.*Worksheet|Test.*OOXML|Test.*Markdown|Test.*Visible' -count=1 .`
- `.xlsx` hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots|BenchmarkXLSXWorksheetTextHotspots' -benchmem -benchtime=1x ./`
- repeat-aware mixed keyset replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-ole-identifier-helper-keyset.json -csv testdata\web-samples\reports\perf-exp-ole-identifier-helper-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`

Observed results:
- focused `.xlsx` regression slice:
  - `ok officeread 8.073s`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: `NO_DIFF`
- but both narrow and broad performance were clearly worse:
  - hotspot benchmarks:
    - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `4827813500 ns/op -> 4694417700 ns/op`
    - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `3350374500 ns/op -> 3521552600 ns/op`
    - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2890754000 ns/op -> 3117330200 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `4761816100 ns/op -> 4882147800 ns/op`
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1922540000 ns/op -> 1964246300 ns/op`
  - repeat-aware mixed keyset:
    - retained baseline total: `44994 ms`
    - experiment total: `51716 ms`
    - by subset:
      - `.docx`: `6433 ms -> 8812 ms`
      - `.xls`: `14471 ms -> 18204 ms`
      - `.xlsx`: `24090 ms -> 24700 ms`

Decision:
- Reverted. This implementation-level helper tweak preserved output but made the retained workload
  substantially slower.

## 2026-07-05 rejected: ASCII fast path for short-query rune cutoffs in markdown backfill

- experiment:
  - leave all containment and matching rules unchanged
  - only replace a few `utf8.RuneCountInString(...) < N` gates with an equivalent helper:
    - if `len(s) < N`, return true immediately
    - if all bytes are ASCII and `len(s) >= N`, skip rune counting
    - only fall back to `utf8.RuneCountInString(...)` for non-ASCII strings
- touched sites:
  - `markdownBackfillContainment.tableTextContainsLine(...)`
  - `markdownBackfillContainment.visibleLineContainsLine(...)`
  - `markdownBackfillComparableContains(...)`
- rationale:
  - the retained `.xls` hotspot still spends most time in containment checks, and these short-query
    rune cutoffs sit directly on that path

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- `.xls` backfill hotspot benchmark:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots' -benchmem -benchtime=1x ./`
- repeat-aware mixed keyset replay:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-rune-count-below-keyset.json -csv testdata\web-samples\reports\perf-exp-rune-count-below-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`

Observed results:
- focused markdown / legacy regression suite:
  - `ok officeread 12.281s`
- output stability:
  - retained-vs-experiment `textBytes` / `images`: `NO_DIFF`
- hotspot benchmark regressed across all three retained `.xls` samples:
  - `006087.xls`: `964043400 ns/op -> 1410474400 ns/op`
  - `008055.xls`: `2236016400 ns/op -> 2497370400 ns/op`
  - `016161.xls`: `889076200 ns/op -> 1056510600 ns/op`
- repeat-aware mixed keyset also regressed clearly:
  - retained baseline total: `44994 ms`
  - experiment total: `50170 ms`
  - by subset:
    - `.docx`: `6433 ms -> 7886 ms`
    - `.xls`: `14471 ms -> 15407 ms`
    - `.xlsx`: `24090 ms -> 26877 ms`

Decision:
- Reverted. Even this purely equivalent ASCII fast path made the retained workload slower.

## 2026-07-05 benchmark refinement: split `.xls` markdown-backfill build costs

- measurement-only change:
  - extend [extract_bench_test.go](/D:/workprj/officeread/extract_bench_test.go) with:
    - `BenchmarkXLSMarkdownExactSetHotspots`
    - `BenchmarkXLSMarkdownCoverageContainmentHotspots`
- rationale:
  - the retained `.xls` hotspot picture had diverged:
    - `006087.xls` profile was dominated by visible-line containment checks
    - `008055.xls` profile still showed heavy exact-set and map-churn cost
  - the previous benchmark set only measured the whole `missingMarkdownTextXLS(...)` path, which was
    too coarse to separate those shapes

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract_bench_test.go`
- benchmark run:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots|BenchmarkXLSMarkdownExactSetHotspots|BenchmarkXLSMarkdownCoverageContainmentHotspots' -benchmem -benchtime=1x ./`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`

Observed results on the retained baseline:
- focused markdown / legacy regression:
  - `ok officeread 11.046s`
- whole backfill:
  - `006087.xls`: `1339661800 ns/op`, `331063944 B/op`, `3453712 allocs/op`
  - `008055.xls`: `2654517100 ns/op`, `573348400 B/op`, `1532733 allocs/op`
  - `016161.xls`: `1128520700 ns/op`, `261876776 B/op`, `968662 allocs/op`
- exact-set build only:
  - `006087.xls`: `247864400 ns/op`, `49135424 B/op`, `590853 allocs/op`
  - `008055.xls`: `1124659900 ns/op`, `237955208 B/op`, `1515380 allocs/op`
  - `016161.xls`: `539833200 ns/op`, `129712200 B/op`, `962010 allocs/op`
- coverage+containment build only:
  - `006087.xls`: `569707600 ns/op`, `260579304 B/op`, `2692516 allocs/op`
  - `008055.xls`: `1620559700 ns/op`, `780180304 B/op`, `2437498 allocs/op`
  - `016161.xls`: `997024500 ns/op`, `485550680 B/op`, `3025037 allocs/op`

Interpretation:
- `006087.xls` remains more containment-heavy than exact-set-heavy
- `008055.xls` remains expensive on both axes, with especially large allocation volume in
  coverage+containment construction
- future `.xls` work should keep treating `006087` and `008055` as different hotspot shapes rather
  than assuming one local win transfers to the other

## 2026-07-05 rejected: build `.xls` exact set and coverage/containment in one combined markdown scan

- experiment:
  - keep the matching rules unchanged
  - change `missingMarkdownTextWithOptions(...)` so it builds:
    - exact-set coverage
    - coverage map
    - containment tables
    in one shared markdown pass instead of two separate passes
- rationale:
  - after the new split benchmarks, avoiding a second full markdown scan looked like a natural
    structural optimization

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots|BenchmarkXLSMarkdownExactSetHotspots|BenchmarkXLSMarkdownCoverageContainmentHotspots' -benchmem -benchtime=1x ./`

Observed results:
- focused markdown / legacy regression:
  - `ok officeread 12.509s`
- whole backfill benchmark:
  - `006087.xls`: `1120972800 ns/op -> 842505700 ns/op`
  - `008055.xls`: `2163713000 ns/op -> 4012425300 ns/op`
  - `016161.xls`: `1017316800 ns/op -> 2067587200 ns/op`
- split benchmarks did not justify retaining it:
  - exact-set-only and coverage-only costs remained in the same general range
  - the regression showed up in the integrated backfill path, consistent with worse peak heap / GC
    behavior from keeping both large structures live together

Decision:
- Reverted. The single-pass idea helped `006087.xls`, but it badly regressed `008055.xls` and
  `016161.xls`, so it does not beat the retained baseline.

## 2026-07-05 rejected: pre-size `.xls` markdown-backfill exact / coverage maps from markdown line count

- experiment:
  - keep all matching logic unchanged
  - pre-split markdown once in `markdownBackfillExactSet(...)` and
    `markdownBackfillBuildCoverageContainment(...)`
  - initialize the main `exact` and `coverage` maps with `len(lines)` capacity
- rationale:
  - the fresh split benchmarks showed large map churn on `008055.xls`, so a low-risk capacity hint
    looked worth checking

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots|BenchmarkXLSMarkdownExactSetHotspots|BenchmarkXLSMarkdownCoverageContainmentHotspots' -benchmem -benchtime=1x ./`

Observed results:
- focused markdown / legacy regression:
  - `ok officeread 12.305s`
- whole backfill benchmark:
  - `006087.xls`: `1120972800 ns/op -> 1342938200 ns/op`
  - `008055.xls`: `2163713000 ns/op -> 2535357200 ns/op`
  - `016161.xls`: `1017316800 ns/op -> 924406600 ns/op`
- split benchmarks were mixed, not broadly positive:
  - exact-set-only:
    - `006087.xls`: `213330600 ns/op -> 298933200 ns/op`
    - `008055.xls`: `1033797000 ns/op -> 1005764400 ns/op`
    - `016161.xls`: `540791500 ns/op -> 587664700 ns/op`
  - coverage+containment-only:
    - `006087.xls`: `802837100 ns/op -> 772214300 ns/op`
    - `008055.xls`: `1947232300 ns/op -> 2193059900 ns/op`
    - `016161.xls`: `1124992300 ns/op -> 1273685400 ns/op`

Decision:
- Reverted. A simple capacity hint was not enough to produce a stable broad win, and it regressed
  two of the three retained `.xls` hotspot samples in the integrated path.

## 2026-07-05 retained: reuse precomputed visible table text when building `.xls` comparable containment

- change:
  - add `markdownBackfillComparableFromVisibleText(...)`
  - when building table containment in:
    - `markdownBackfillBuildCoverageContainment(...)`
    - `markdownBackfillContainmentSet(...)`
  - compute `visible := markdownBackfillVisibleText(...)` once per table row and derive
    `comparable` from that visible text instead of recomputing `markdownBackfillVisibleText(...)`
    inside `markdownBackfillComparableText(...)`
- rationale:
  - the new split `.xls` benchmarks showed that containment construction still carried very large
    allocation volume
  - for table rows, the old path was visibly recomputing the same normalized visible string just to
    reach the comparable form

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots|BenchmarkXLSMarkdownExactSetHotspots|BenchmarkXLSMarkdownCoverageContainmentHotspots' -benchmem -benchtime=1x ./`
- repeat-aware retained mixed keyset:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-visible-comparable-reuse-keyset.json -csv testdata\web-samples\reports\perf-exp-visible-comparable-reuse-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused markdown / legacy regression:
  - `ok officeread 11.648s`
- full repository regression:
  - `ok officeread 129.909s`
- hotspot benchmarks:
  - whole backfill:
    - `006087.xls`: `1339661800 ns/op -> 1204912700 ns/op`
    - `008055.xls`: `2654517100 ns/op -> 1943883800 ns/op`
    - `016161.xls`: `1128520700 ns/op -> 800417600 ns/op`
  - exact-set-only:
    - `006087.xls`: `247864400 ns/op -> 265334900 ns/op`
    - `008055.xls`: `1124659900 ns/op -> 895644500 ns/op`
    - `016161.xls`: `539833200 ns/op -> 395615800 ns/op`
  - coverage+containment-only:
    - `006087.xls`: `569707600 ns/op -> 509293800 ns/op`
    - `008055.xls`: `1620559700 ns/op -> 1538942900 ns/op`
    - `016161.xls`: `997024500 ns/op -> 900134500 ns/op`
- repeat-aware retained mixed keyset:
  - retained baseline total: `44994 ms`
  - experiment total: `41943 ms`
  - by subset:
    - `.docx`: `6433 ms -> 6107 ms`
    - `.xls`: `14471 ms -> 13012 ms`
    - `.xlsx`: `24090 ms -> 22824 ms`
- output stability against the retained current-keyset report:
  - `textBytes` / `images`: `NO_DIFF`

Decision:
- Retained. This is the first recent `.xls` markdown-backfill change in this area that improves all
  three retained hotspot samples, improves the repeat-aware mixed keyset, preserves output, and
  passes the full repository regression suite.

## 2026-07-05 retained: guard HTML and `<br>` table-cell expansion work inside `.xls` markdown backfill

- change:
  - narrow `markdownVisibleTableCells(...)` so it:
    - only calls `markdownVisibleHTMLText(...)` when the normalized cell still contains `<`
    - only splits the cell by `<br>` when the normalized cell actually contains `<br>`
- rationale:
  - fresh post-profile work still showed `markdownVisibleTableCells(...)` as a meaningful cost under
    `.xls` exact-set and coverage construction, especially on the large table-heavy samples
  - most cells do not contain HTML tags or `<br>` markers, so unconditional work there was paying
    for many empty cases

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots|BenchmarkXLSMarkdownExactSetHotspots|BenchmarkXLSMarkdownCoverageContainmentHotspots' -benchmem -benchtime=1x ./`
- repeat-aware retained mixed keyset:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata\web-samples\reports\perf-exp-tablecell-htmlbr-guards-keyset.json -csv testdata\web-samples\reports\perf-exp-tablecell-htmlbr-guards-keyset.csv @((Get-Content -Raw testdata\web-samples\reports\perf-exp-short-exact-xls-only-keyset.json | ConvertFrom-Json).inputs)`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused markdown / legacy regression:
  - `ok officeread 11.506s`
- full repository regression:
  - `ok officeread 128.440s`
- hotspot benchmarks against the previous retained visible/comparable-reuse baseline:
  - whole backfill:
    - `006087.xls`: `1204912700 ns/op -> 1048374300 ns/op`
    - `008055.xls`: `1943883800 ns/op -> 1964941500 ns/op`
    - `016161.xls`: `800417600 ns/op -> 671926200 ns/op`
  - exact-set-only:
    - `006087.xls`: `265334900 ns/op -> 241945500 ns/op`
    - `008055.xls`: `895644500 ns/op -> 918970900 ns/op`
    - `016161.xls`: `395615800 ns/op -> 412875600 ns/op`
  - coverage+containment-only:
    - `006087.xls`: `509293800 ns/op -> 483859100 ns/op`
    - `008055.xls`: `1538942900 ns/op -> 1343779200 ns/op`
    - `016161.xls`: `900134500 ns/op -> 784124000 ns/op`
- allocation deltas were especially large on the retained `.xls` hotspots:
  - whole backfill allocs/op:
    - `006087.xls`: `3003778 -> 2532354`
    - `008055.xls`: `1532722 -> 215688`
    - `016161.xls`: `968729 -> 413958`
- repeat-aware retained mixed keyset:
  - previous retained baseline total: `41943 ms`
  - experiment total: `40807 ms`
  - by subset:
    - `.docx`: `6107 ms -> 5929 ms`
    - `.xls`: `13012 ms -> 12026 ms`
    - `.xlsx`: `22824 ms -> 22852 ms`
- output stability against the retained current-keyset report:
  - `textBytes` / `images`: `NO_DIFF`

Decision:
- Retained. This change keeps output stable, improves the repeat-aware mixed workload again, passes
  the full suite, and further reduces `.xls` hotspot allocation pressure.

## 2026-07-05 rejected: replace table-row split in `markdownLikelyTableRow(...)` with early unescaped-pipe counting

- experiment:
  - keep table-row detection semantics aligned with the old rule
  - but replace `len(splitMarkdownTableRow(trimmed)) >= 3` with a lighter byte scan that returns as
    soon as it sees two unescaped `|` separators
- rationale:
  - fresh `006087.xls` profiling still showed string-search and row-shape work around markdown table
    handling, so avoiding a full split in the yes/no predicate looked promising

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots|BenchmarkXLSMarkdownExactSetHotspots|BenchmarkXLSMarkdownCoverageContainmentHotspots' -benchmem -benchtime=1x ./`

Observed results against the retained HTML/`<br>`-guard baseline:
- focused markdown / legacy regression:
  - `ok officeread 11.280s`
- whole backfill:
  - `006087.xls`: `1048374300 ns/op -> 1130252700 ns/op`
  - `008055.xls`: `1964941500 ns/op -> 2102875200 ns/op`
  - `016161.xls`: `671926200 ns/op -> 774968100 ns/op`
- split benchmarks were mixed:
  - exact-set-only improved:
    - `006087.xls`: `241945500 ns/op -> 209047300 ns/op`
    - `008055.xls`: `918970900 ns/op -> 635728700 ns/op`
    - `016161.xls`: `412875600 ns/op -> 409518400 ns/op`
  - but coverage+containment regressed:
    - `006087.xls`: `483859100 ns/op -> 531688700 ns/op`
    - `008055.xls`: `1343779200 ns/op -> 1478287900 ns/op`
    - `016161.xls`: `784124000 ns/op -> 1019494900 ns/op`

Decision:
- Reverted. The lighter predicate helped exact-set work, but it clearly lost on the integrated
  retained `.xls` path.

## 2026-07-05 rejected: skip segment-level HTML text cleanup when a `<br>` segment contains no `<`

- experiment:
  - keep the retained outer `markdownVisibleTableCells(...)` guards
  - within the `<br>` segment loop, only call `markdownVisibleHTMLText(...)` when the segment itself
    still contains `<`
- rationale:
  - after the retained outer guards, this looked like one more low-risk place to avoid empty-case
    string work

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots|BenchmarkXLSMarkdownExactSetHotspots|BenchmarkXLSMarkdownCoverageContainmentHotspots' -benchmem -benchtime=1x ./`

Observed results against the retained HTML/`<br>`-guard baseline:
- focused markdown / legacy regression:
  - `ok officeread 14.047s`
- whole backfill:
  - `006087.xls`: `1048374300 ns/op -> 1248900300 ns/op`
  - `008055.xls`: `1964941500 ns/op -> 2019862500 ns/op`
  - `016161.xls`: `671926200 ns/op -> 1003614300 ns/op`
- split benchmarks again looked locally attractive but lost on the integrated path:
  - exact-set-only:
    - `006087.xls`: `241945500 ns/op -> 188114700 ns/op`
    - `008055.xls`: `918970900 ns/op -> 691349500 ns/op`
    - `016161.xls`: `412875600 ns/op -> 286006000 ns/op`
  - coverage+containment:
    - `006087.xls`: `483859100 ns/op -> 697108800 ns/op`
    - `008055.xls`: `1343779200 ns/op -> 1367041500 ns/op`
    - `016161.xls`: `784124000 ns/op -> 989985600 ns/op`

Decision:
- Reverted. Another reminder that shaving local helper work inside table-cell extraction does not
  automatically beat the current retained integrated `.xls` workload.

## 2026-07-05 rejected: skip visible-line substring containment for non-exact queries containing `|`

- experiment:
  - in `markdownBackfillContainment.visibleLineContainsLine(...)`
  - keep exact `visibleExact` hits
  - but skip the fallback `strings.Contains(c.visibleJoined, line)` when the query still contains `|`
- rationale:
  - the fresh `006087.xls` profile remained dominated by visible-line substring containment
  - table-like candidates containing `|` already flow through `tableTextContainsLine(...)`, so this
    looked like a plausible way to trim redundant visible-line probes

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots|BenchmarkXLSMarkdownExactSetHotspots|BenchmarkXLSMarkdownCoverageContainmentHotspots' -benchmem -benchtime=1x ./`

Observed results against the retained HTML/`<br>`-guard baseline:
- focused markdown / legacy regression:
  - `ok officeread 14.278s`
- whole backfill:
  - `006087.xls`: `1048374300 ns/op -> 1107237700 ns/op`
  - `008055.xls`: `1964941500 ns/op -> 2163977600 ns/op`
  - `016161.xls`: `671926200 ns/op -> 920007200 ns/op`
- exact-set-only remained positive:
  - `006087.xls`: `241945500 ns/op -> 187595700 ns/op`
  - `008055.xls`: `918970900 ns/op -> 820619500 ns/op`
  - `016161.xls`: `412875600 ns/op -> 348670700 ns/op`
- but coverage+containment clearly regressed:
  - `006087.xls`: `483859100 ns/op -> 794799600 ns/op`
  - `008055.xls`: `1343779200 ns/op -> 1914329200 ns/op`
  - `016161.xls`: `784124000 ns/op -> 1116100000 ns/op`

Decision:
- Reverted. Skipping visible-line substring checks for `|`-bearing queries looked plausible, but it
  lost badly on the retained integrated `.xls` workload.

## 2026-07-05 benchmark refinement: replay real `.xls` containment queries separately

- measurement-only change:
  - extend [extract_bench_test.go](/D:/workprj/officeread/extract_bench_test.go) with:
    - `BenchmarkXLSVisibleContainmentReplayHotspots`
    - `BenchmarkXLSTableContainmentReplayHotspots`
  - add a helper that reconstructs the actual candidate/variant query set used by
    `missingMarkdownTextWithOptions(...)` and replays those queries directly against prepared
    containment state
- rationale:
  - recent profiles kept saying `006087.xls` was dominated by visible-line containment, but the
    existing benchmarks still only showed whole-path timings
  - to choose the next structural optimization well, we needed to know whether the cost was really
    in `visibleLineContainsLine(...)`, in `tableTextContainsLine(...)`, or in the way queries were
    shaped

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract_bench_test.go`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- focused replay benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls$' -benchmem -benchtime=1x ./`
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSTableContainmentReplayHotspots/006087.xls$' -benchmem -benchtime=1x ./`

Observed results:
- focused markdown / legacy regression:
  - `ok officeread 18.636s`
- replay benchmarks on `006087.xls`:
  - `BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls`: `11825385100 ns/op`, `0 B/op`, `0 allocs/op`
  - `BenchmarkXLSTableContainmentReplayHotspots/006087.xls`: `6055400 ns/op`, `144 B/op`, `7 allocs/op`

Interpretation:
- the dominant remaining `006087.xls` cost is decisively in visible-line containment replay, not in
  table containment replay
- the zero-allocation shape of the visible replay benchmark confirms that the next meaningful `.xls`
  optimization needs to reduce pure substring-search work, not merely allocation churn

## 2026-07-05 rejected: large-visible-haystack suffix index for visible-line containment

- experiment:
  - add an optional suffix-array index for `markdownBackfillContainment.visibleJoined`
  - build it only when the joined visible-line haystack exceeds a size threshold
  - route `visibleLineContainsLine(...)` through the suffix index when present
- rationale:
  - the new replay benchmarks showed that `006087.xls` spends almost all remaining time in repeated
    visible-line substring probes against one large haystack, which is exactly the kind of workload
    a suffix index should help

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- focused replay / hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls$|BenchmarkXLSVisibleContainmentReplayHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`

Observed results:
- focused markdown / legacy regression:
  - `ok officeread 12.091s`
- replay benchmarks:
  - `006087.xls` visible replay:
    - baseline: `11825385100 ns/op`, `0 B/op`, `0 allocs/op`
    - suffix-index experiment: `26957200 ns/op`, `293280 B/op`, `36660 allocs/op`
  - `008055.xls` visible replay under the suffix-index experiment:
    - `2314441100 ns/op`, `9441472 B/op`, `1180182 allocs/op`
- integrated hotspot benchmarks:
  - `006087.xls`: `1108063500 ns/op -> 865251400 ns/op`
  - `008055.xls`: `1738540000 ns/op -> 2506948700 ns/op`
  - `016161.xls`: `781564800 ns/op -> 949258000 ns/op`
  - coverage+containment build also regressed badly on `008055.xls`:
    - `1343779200 ns/op -> 3665734000 ns/op`

Interpretation:
- the suffix index absolutely helps the `006087.xls` replay shape
- but the build cost and per-query lookup overhead do not survive the broader retained `.xls`
  workload, especially on `008055.xls`

Decision:
- Reverted. This was the first experiment to clearly beat the `006087.xls` replay shape, but it
  still lost on the integrated retained hotspot mix, so it cannot enter the baseline yet.

## 2026-07-05 rejected: 4-byte visible substring Bloom-style prefilter

- experiment:
  - add a light-weight 4-byte substring bitset for large `visibleJoined` haystacks
  - before `strings.Contains(c.visibleJoined, line)`, reject queries whose first 4-byte window is
    definitely absent from the haystack
- rationale:
  - after the suffix-array experiment, the goal was to keep targeting the right `006087.xls`
    replay bottleneck, but with a much cheaper build/query structure

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- focused replay / hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls$|BenchmarkXLSVisibleContainmentReplayHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`

Observed results:
- focused markdown / legacy regression:
  - `ok officeread 33.154s`
- replay benchmarks:
  - `006087.xls` visible replay regressed:
    - baseline: `10330479400 ns/op`, `0 B/op`, `0 allocs/op`
    - quad-filter experiment: `15930762700 ns/op`, `0 B/op`, `0 allocs/op`
  - `008055.xls` visible replay also remained expensive:
    - `2314441100 ns/op`, `9441472 B/op`, `1180182 allocs/op`
- integrated hotspot benchmarks regressed across the board:
  - `006087.xls`: `1108063500 ns/op -> 2576643600 ns/op`
  - `008055.xls`: `1738540000 ns/op -> 4603367800 ns/op`
  - `016161.xls`: `781564800 ns/op -> 1917064200 ns/op`

Interpretation:
- this prefilter was too lossy to reduce the real replay workload
- the extra build/check work simply stacked on top of the existing `strings.Contains(...)` cost

Decision:
- Reverted. This was cheaper than the suffix-array experiment, but it still failed to improve the
  retained replay and hotspot shapes.

## 2026-07-05 rejected: exact 4-byte window set targeted only at the `006087`-style visible replay shape

- experiment:
  - add a more selective visible-line prefilter than the earlier Bloom-style bitset:
    - build an exact set of 4-byte windows from `visibleJoined`
    - only enable it for haystacks shaped like `006087.xls`:
      - moderate joined size
      - many visible lines
      - relatively small max visible-line length
  - on lookup, reject only when the query's first 4-byte window is absent from that exact set
- rationale:
  - the earlier fuzzy 4-byte filter had too many false positives
  - the stats gathered this turn showed a strong shape split:
    - `006087.xls`: `visibleJoined=4587667`, `visibleLines=54601`, `queries=41339`, `visibleMax=271`
    - `008055.xls`: `visibleJoined=22697129`, `visibleLines=16917`, `queries=1181171`, `visibleMax=11343`
  - so a targeted exact-window set looked like a plausible way to help `006087` without touching the
    very different `008055` shape

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- focused replay / hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls$|BenchmarkXLSVisibleContainmentReplayHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`

Observed results:
- focused markdown / legacy regression:
  - `ok officeread 11.226s`
- integrated hotspot benchmarks versus the retained baseline:
  - `006087.xls`: `1174221600 ns/op -> 1152815600 ns/op`
  - `008055.xls`: `2713071900 ns/op -> 1609151700 ns/op`
  - `016161.xls`: `808511400 ns/op -> 739579100 ns/op`
  - containment-build view:
    - `006087.xls`: `532688500 ns/op`
    - `008055.xls`: `1648300000 ns/op`
- but the key target did not move enough:
  - `006087.xls` visible replay remained effectively unchanged:
    - baseline: `10330479400 ns/op`
    - targeted exact-window experiment: `10305839900 ns/op`

Interpretation:
- the exact first-window presence test is precise, but still not discriminative enough to prune the
  real `006087.xls` query stream
- the remaining visible replay cost is coming from queries whose leading 4-byte window is common,
  so the expensive full substring scan still runs almost every time

Decision:
- Reverted. This was much more targeted than the earlier fuzzy 4-byte filter, but it still did not
  materially reduce the actual `006087.xls` visible replay bottleneck.

## 2026-07-05 rejected: lazy visible-line prefix buckets keyed by the first 8 bytes of the query

- experiment:
  - for `006087`-like visible containment shapes, enable lazy per-prefix buckets:
    - when `visibleLineContainsLine(...)` sees a query of length at least 8
    - use the query's first 8 bytes as a key
    - lazily build and cache a joined subset of `visibleLines` that contain that prefix
    - run the final `strings.Contains(...)` against that bucket instead of the whole `visibleJoined`
- rationale:
  - fresh shape stats showed that the replay workload had far fewer unique first-8-byte prefixes
    than total queries:
    - `006087.xls`: `queries=41339`, `unique8=329`, `max8=8`
  - this made lazy prefix bucketing look like a plausible way to cut the effective haystack size
    without paying an up-front full-text index cost

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 .`
- focused replay / hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls$|BenchmarkXLSVisibleContainmentReplayHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`

Observed results:
- focused markdown / legacy regression:
  - `ok officeread 10.850s`
- integrated hotspot view looked encouraging:
  - `006087.xls`: `1083877100 ns/op -> 985549600 ns/op`
  - `008055.xls`: `2151532100 ns/op -> 1705911700 ns/op`
  - `016161.xls`: `637898400 ns/op -> 711650300 ns/op`
- containment-build view also looked acceptable:
  - `006087.xls`: `477150800 ns/op`
  - `008055.xls`: `1667158500 ns/op`
- but the key target did not materially improve:
  - `006087.xls` visible replay:
    - baseline: `10855155300 ns/op`
    - prefix-bucket experiment: `10938179300 ns/op`
    - allocs/op also appeared: `39703000 B/op`, `935 allocs/op`

Interpretation:
- shrinking the candidate haystack by prefix was not enough to reduce the real replay workload
- the expensive queries still hit large buckets, so the final substring scans remained dominant

Decision:
- Reverted. This was the most promising “cheap structural” experiment so far, but it still failed on
  the actual `006087.xls` visible replay target.
## 2026-07-05 rejected: narrower shape-gated suffix index for `006087`-like visible replay

- experiment:
  - revisit the earlier suffix-array direction, but only enable it for a narrower visible-line
    shape:
    - `len(visibleJoined)` between `1 MiB` and `8 MiB`
    - `len(visibleLines) >= 30000`
    - `visibleMaxLen <= 512`
  - this intentionally targets the measured `006087.xls` shape while excluding larger
    `008055.xls`-style haystacks
- rationale:
  - the broad suffix-array version had already proven that indexed replay can crush the pure
    substring-search cost on `006087.xls`
  - the next question was whether tighter gating could keep that win without dragging the broader
    retained hotspot set

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused sample regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestExtractSamples/(testdata/web-samples/xls/006087.xls|testdata/web-samples/xls/008055.xls|testdata/web-samples/xls/016161.xls)$' ./`
- focused replay / hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls$|BenchmarkXLSVisibleContainmentReplayHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`

Observed results:
- focused sample regression stayed green:
  - `ok officeread 0.707s`
- the targeted replay shape improved dramatically again:
  - `006087.xls` visible replay:
    - baseline: `10855155300 ns/op`, `0 B/op`, `0 allocs/op`
    - narrower shape-gated suffix experiment: `30303700 ns/op`, `296976 B/op`, `36721 allocs/op`
- but the retained integrated hotspot mix still did not hold up consistently:
  - `006087.xls`: `1083877100 ns/op -> 822505700 ns/op`
  - `008055.xls`: `2151532100 ns/op -> 2110890500 ns/op`
  - `016161.xls`: `637898400 ns/op -> 705303900 ns/op`
  - containment-build view:
    - `006087.xls`: `767560000 ns/op`
    - `008055.xls`: `1993346500 ns/op`

Interpretation:
- the narrower gate successfully isolates the right replay shape, so the core algorithmic signal is
  real
- but even with that tighter gate, suffix-index construction and lookup costs still do not produce a
  stable enough integrated win across the retained `.xls` hotspot mix

Decision:
- Reverted. This is now the second suffix-index variant to beat the pure `006087.xls` replay target
  decisively while still failing the broader retained workload.

## 2026-07-05 baseline rerun after reverting the narrower suffix gate

- validation:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls$|BenchmarkXLSTableContainmentReplayHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata/web-samples/reports/perf-rerun-20260705-current-keyset-rerun2.json -csv testdata/web-samples/reports/perf-rerun-20260705-current-keyset-rerun2.csv <keyset>`
- benchmark baseline after revert:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `1110686900 ns/op`, `302044800 B/op`, `2532336 allocs/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `2338682200 ns/op`, `552087192 B/op`, `215683 allocs/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `877876400 ns/op`, `253360952 B/op`, `413942 allocs/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `679919100 ns/op`, `235380488 B/op`, `2006870 allocs/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1557958400 ns/op`, `701203200 B/op`, `917397 allocs/op`
  - `BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls`: `14059815600 ns/op`, `0 B/op`, `0 allocs/op`
  - `BenchmarkXLSTableContainmentReplayHotspots/006087.xls`: `4019000 ns/op`, `144 B/op`, `7 allocs/op`
- repeat-aware keyset rerun:
  - output parity held for all 30 inputs:
    - `.docx`: `2/2 ok`, `1213267` text bytes, `43` images, total `6931 ms`
    - `.xls`: `7/7 ok`, `32445758` text bytes, `0` images, total `17830 ms`
    - `.xlsx`: `21/21 ok`, `279385415` text bytes, `1` image, total `23756 ms`
- note:
  - these rerun totals are better than the earlier reference file
    `perf-rerun-20260705-current-keyset.json`, but the code is the same retained baseline after
    revert, so this should be treated as normal run-to-run variance rather than evidence of a new
    retained optimization

## 2026-07-05 retained: shape-gated short-digit visible substring set for `.xls` replay hits

- experiment:
  - keep the existing visible substring semantics, but special-case the exact workload shape that
    remained after the suffix-array branches were reverted
  - when `visibleJoined` is between `1 MiB` and `8 MiB`, `visibleLines >= 30000`, and
    `visibleMaxLen <= 512`, precompute all ASCII-digit substrings of widths `4..7` from contiguous
    visible numeric runs
  - for `visibleLineContainsLine(...)`, if the query is all ASCII digits and `4..7` bytes long,
    answer from that precomputed set instead of scanning the full `visibleJoined` haystack
- rationale:
  - fresh instrumentation on the integrated path showed that `006087.xls` was no longer making
    tens of thousands of expensive visible replay calls in practice
  - it was making only `404` visible substring queries, all hits, and the leading examples were
    short numeric fragments such as `61047`, `30357`, `21595`, `123311`, and `5200`
  - those values are covered by substring semantics inside longer visible numeric runs, so an exact
    short-digit substring set can preserve current behavior while removing the repeated full-haystack
    scan cost

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused markdown / legacy regression suite:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^Test.*Markdown|Test.*XLS|Test.*Biff|Test.*BIFF|Test.*Visible|Test.*Legacy|TestMarkdownBackfill|TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns|TestDOCXEscapedTableTextDoesNotNeedAdditionalText|TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable|TestUnicodeBinaryNoiseFilterKeepsRealCJKText|TestLegacySamplesDoNotEmitUnicodeBinaryNoise|TestBIFFHeaderFooterControlCodesAreNotVisibleText|TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput|TestLegacyXLSDropsEmbeddedPDFText|TestBIFFCellRecordInMarkdownBounds' -count=1 ./`
- full repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`
- focused sample regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestExtractSamples/(testdata/web-samples/xls/006087.xls|testdata/web-samples/xls/008055.xls|testdata/web-samples/xls/016161.xls)$' ./`
- focused replay / hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls$|BenchmarkXLSTableContainmentReplayHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
- repeat-aware broader keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 2 -json testdata/web-samples/reports/perf-exp-xls-short-digit-substrings-keyset.json -csv testdata/web-samples/reports/perf-exp-xls-short-digit-substrings-keyset.csv <keyset>`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xls-short-digit-substrings-xls6.json -csv testdata/web-samples/reports/perf-exp-xls-short-digit-substrings-xls6.csv <xls6>`

Observed results:
- focused regression suite:
  - `ok officeread 11.537s`
- full repository regression:
  - `ok officeread 149.727s`
- focused sample regression:
  - `ok officeread 2.822s`
- focused replay / hotspot benchmarks versus the reverted baseline:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `1110686900 ns/op -> 760822200 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `2338682200 ns/op -> 2104182200 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `877876400 ns/op -> 814307500 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `679919100 ns/op -> 495409100 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1557958400 ns/op -> 1489791600 ns/op`
  - `BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls`: `14059815600 ns/op -> 22748600 ns/op`
  - `BenchmarkXLSTableContainmentReplayHotspots/006087.xls`: `4019000 ns/op -> 4088400 ns/op`
- repeat-aware broader keyset rerun:
  - output parity held for all 30 inputs
  - `.xls`: `17830 ms -> 15160 ms`
  - `.docx`: `6931 ms -> 7447 ms`
  - `.xlsx`: `23756 ms -> 27947 ms`
  - because the code path is `.xls`-only, the non-`.xls` movement should be read as runtime noise,
    not as a causal regression from this change
- repeat-aware `.xls6` hotspot rerun:
  - output parity held for all 6 inputs
  - total: `23111 ms -> 10461 ms`
  - per-file:
    - `006087.xls`: `5905 ms -> 1351 ms`
    - `008055.xls`: `7629 ms -> 3922 ms`
    - `016161.xls`: `2786 ms -> 1665 ms`
    - `002505.xls`: `3096 ms -> 1387 ms`
    - `019088.xls`: `1871 ms -> 1077 ms`
    - `019089.xls`: `1824 ms -> 1059 ms`

Decision:
- Retained. Unlike the earlier short-digit guard branches, this preserves the existing visible
  substring semantics and converts the measured short-digit `.xls` hit shape into a cheap exact-set
  query. The hotspot and repeat-aware `.xls` reruns both improved materially with stable output.

## 2026-07-06 rejected: skip exact precheck only for the `006087`-like short-text `.xls` shape

- experiment:
  - add a very narrow `.xls`-only escape hatch before `markdownBackfillExactSet(...)`
  - when:
    - `escapedTableOnlyWhenPipe == true`
    - `shortTableExactBeforeMinLen == true`
    - `len(lines) >= 30000`
    - `len(text) <= 512 KiB`
  - skip the exact precheck entirely and go straight to the retained
    `coverage + containment + missing-line` path
- rationale:
  - fresh exact-precheck instrumentation showed a split outcome across the retained hotspot set:
    - `006087.xls`: `36339` exact hits before the first miss, and that miss happened at the very
      end on short numeric fragment `61047`
    - `008055.xls`: exact precheck succeeded completely with `1181169` exact hits and no miss
    - `016161.xls`: exact precheck succeeded completely with `461603` exact hits and no miss
  - that made `006087.xls` look like a plausible narrow candidate for skipping the expensive exact
    phase, while keeping the retained early-return behavior for the other heavy `.xls` cases

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused replay / hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$|BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls$' -benchmem -benchtime=1x ./`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xls-skip-exact-006087shape-xls6.json -csv testdata/web-samples/reports/perf-exp-xls-skip-exact-006087shape-xls6.csv <xls6>`

Observed results:
- focused hotspot view was mixed:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `760822200 ns/op -> 577143000 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `2104182200 ns/op -> 2109205900 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `814307500 ns/op -> 886576100 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `495409100 ns/op -> 682640500 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1489791600 ns/op -> 1550357600 ns/op`
  - `BenchmarkXLSVisibleContainmentReplayHotspots/006087.xls`: `22748600 ns/op -> 26933400 ns/op`
- repeat-aware `.xls6` hotspot rerun regressed against the retained short-digit-substring baseline:
  - retained baseline total: `10461 ms`
  - experiment total: `11758 ms`
  - per-file:
    - `006087.xls`: `1351 ms -> 1652 ms`
    - `008055.xls`: `3922 ms -> 4826 ms`
    - `016161.xls`: `1665 ms -> 1654 ms`
    - `002505.xls`: `1387 ms -> 1470 ms`
    - `019088.xls`: `1077 ms -> 1063 ms`
    - `019089.xls`: `1059 ms -> 1093 ms`

Interpretation:
- the exact-precheck skip helped the isolated `006087.xls` backfill benchmark, but it removed too
  much useful structure from the integrated path
- once measured on the retained `.xls6` batch, the branch clearly lost to the already-retained
  short-digit visible substring optimization

Decision:
- Reverted. The evidence says the exact precheck is still worth keeping even when `006087.xls`
  misses only at the tail.

## 2026-07-06 rejected: `.xls` exact precheck with direct-exact first and full helper on miss

- experiment:
  - keep the retained generic path unchanged for non-`.xls` callers
  - only for the existing `.xls` exact precheck, split the logic into:
    - first try a direct `exact[line]` lookup
    - only if that misses, fall back to full `markdownBackfillExactSetContainsLine(...)`
- rationale:
  - fresh hit-shape instrumentation showed that the retained `.xls` hotspots were extremely skewed:
    - `006087.xls`: `direct=40935`, `visible=0`, `markdown=0`, `escaped=0`, `misses=404`
    - `008055.xls`: `direct=1181169`, `visible=0`, `markdown=0`, `escaped=0`, `misses=0`
    - `016161.xls`: `direct=461603`, `visible=0`, `markdown=0`, `escaped=0`, `misses=0`
  - that made it look plausible that the `.xls` exact precheck could avoid calling the more general
    helper almost all the time without changing any matching semantics

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused replay / hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xls-direct-exact-first-xls6.json -csv testdata/web-samples/reports/perf-exp-xls-direct-exact-first-xls6.csv <xls6>`

Observed results:
- focused hotspot view regressed on the main retained heavy samples:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `700638000 ns/op -> 757829500 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1589953500 ns/op -> 2353962100 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `658728600 ns/op -> 865504500 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: about `803362600 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: about `1555288300 ns/op`
- repeat-aware `.xls6` hotspot rerun also lost against the retained short-digit-substring baseline:
  - retained baseline total: `10461 ms`
  - experiment total: `12051 ms`
  - per-file:
    - `002505.xls`: `1387 ms -> 1436 ms`
    - `006087.xls`: `1351 ms -> 1692 ms`
    - `008055.xls`: `3922 ms -> 4933 ms`
    - `016161.xls`: `1665 ms -> 1771 ms`
    - `019088.xls`: `1077 ms -> 1062 ms`
    - `019089.xls`: `1059 ms -> 1157 ms`

Interpretation:
- even though the hit distribution looked perfectly aligned with a direct-first strategy, the extra
  front-door branch plus the duplicated direct lookup still made the retained `.xls` path slower
- the existing generic helper remains the better integrated choice for the current workload mix

Decision:
- Reverted. This is another case where a compelling local distribution story did not survive the
  repeat-aware retained benchmarks.

## 2026-07-06 rejected: prune table-row raw/visible entries from `.xls` exact-set construction

- experiment:
  - keep the generic exact-set builder unchanged for non-`.xls` paths
  - for the `.xls` exact precheck only, build a narrower exact set:
    - on markdown lines recognized as table rows, do not add the whole trimmed row string
    - also do not add the row-level `markdownVisibleLineText(...)` result
    - still add every extracted visible table cell exactly as before
- rationale:
  - fresh origin instrumentation showed that the retained heavy `.xls` exact hits were dominated by
    cell-level matches, not row-level matches:
    - `008055.xls`: `rawHits=0`, `visibleHits=3`, `cellHits=1181166`, `miss=0`
    - `016161.xls`: `rawHits=0`, `visibleHits=6`, `cellHits=461597`, `miss=0`
    - `006087.xls`: `rawHits=4598`, `visibleHits=1`, `cellHits=36336`, `miss=404`
  - a temporary replay with the pruned exact set preserved the same hit/miss totals on those
    samples while shrinking exact-set size:
    - `008055.xls`: `1305746 -> 1288836`
    - `016161.xls`: `566236 -> 520619`
    - `006087.xls`: `86539 -> 42557`

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused replay / hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/006087.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
- repeat-aware `.xls6` hotspot rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xls-prune-table-row-exact-xls6.json -csv testdata/web-samples/reports/perf-exp-xls-prune-table-row-exact-xls6.csv <xls6>`

Observed results:
- focused hotspot view regressed across the retained `.xls` samples:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 845788200 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 2210761600 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 870965300 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `173512500 ns/op -> 235957700 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 853177900 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `278034100 ns/op -> 384341000 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: about `635961400 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: about `1723214500 ns/op`
- repeat-aware `.xls6` hotspot rerun also lost against the retained baseline:
  - retained baseline total: `10461 ms`
  - experiment total: `11677 ms`
  - per-file:
    - `002505.xls`: `1387 ms -> 1650 ms`
    - `006087.xls`: `1351 ms -> 1758 ms`
    - `008055.xls`: `3922 ms -> 5085 ms`
    - `016161.xls`: `1665 ms -> 2069 ms`
    - `019088.xls`: `1077 ms -> 1115 ms`
    - `019089.xls`: `1059 ms -> 1100 ms`

Interpretation:
- reducing exact-set size was not enough; the removed row-level entries were apparently still
  helping the integrated `.xls` backfill flow in ways the isolated hit-origin counts did not
  capture
- another reminder that exact-set cardinality alone is not the right optimization target here

Decision:
- Reverted.

## 2026-07-06 rejected: shared-string prepared plain-cell markdown fast path

- experiment:
  - leave `cleanMarkdownTableCellValue(...)` and the simple-inline pipeline untouched
  - only inside `appendSharedStringWorksheetTextPrepared(...)`, add a narrow fast path for
    markdown-collected shared-string values:
    - reject anything containing newline or obvious markdown / hidden-reference markers such as
      `\\`, `|`, `/`, `%`, `=`, `:`, brackets, braces, `<`, `>`, `#`, `&`, or `rid`
    - require `cleanTextFastPath(...)` to succeed
    - require the cleaned value to pass the existing discard / control checks
    - if all of the above hold, return `normalizeMarkdownTextLine(cleaned)` directly and skip the
      generic `cleanMarkdownTableCellValue(...) + prepareMarkdownTableCellValue(...)` path
- rationale:
  - the newly retained shared-string stage splits showed:
    - `Prepared + Markdown`: about `676 ms`
    - `Prepared NoMarkdown`: about `281 ms`
    - `MarkdownClean`: about `519 ms`
    - `MarkdownPrepare`: about `145 ms`
  - that made a shared-string-only markdown common path look like the best way to chase the proven
    `cleanMarkdownTableCellValue(...)` cost without perturbing the simple-inline helper stack

Validation:
- targeted regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run 'TestPreparePlainSharedStringMarkdownCell|TestCleanMarkdownTableCellValue|TestPrepareMarkdownTableCellValueSingleLineFastPath|TestExtractXLSX|TestExtractWorksheet|Test.*Text|Test.*Visible' ./`
- shared-string stage benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXSharedStringPrepared00012389$|BenchmarkXLSXSharedStringPreparedNoMarkdown00012389$|BenchmarkXLSXSharedStringMarkdownClean00012389$|BenchmarkXLSXSharedStringMarkdownPrepare00012389$' -benchmem -benchtime=1x ./`
- integrated shared-string hotspot checks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$' -benchmem -benchtime=1x ./`
- repeat-aware serial pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-shared-plaincell-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-shared-plaincell-pair-serial.csv testdata/web-samples/samples/xlsx/00012389.xlsx testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- decisive 21-file `.xlsx` keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata/web-samples/reports/perf-exp-ai-assistant-shared-plaincell-xlsx-keyset-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-shared-plaincell-xlsx-keyset-serial.csv <same-21-xlsx-inputs>`

Observed results:
- targeted regression stayed green before revert:
  - `ok officeread 32.272s`
- shared-string stage benchmarks improved strongly:
  - `BenchmarkXLSXSharedStringPrepared00012389`: `675977600 ns/op -> 560218800 ns/op`
  - `BenchmarkXLSXSharedStringPreparedNoMarkdown00012389`: `281461300 ns/op -> 278195600 ns/op`
  - `BenchmarkXLSXSharedStringMarkdownClean00012389`: `519055600 ns/op -> 303134700 ns/op`
  - `BenchmarkXLSXSharedStringMarkdownPrepare00012389`: `144604600 ns/op -> 66192000 ns/op`
- integrated shared-string hotspot checks also improved materially:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `1643944300 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `841605300 ns/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`: `540814900 ns/op`
  - allocations dropped sharply on the shared-string worksheet path:
    - extract path: `618491424 B/op, 3266791 allocs/op -> 601107792 B/op, 2545723 allocs/op`
    - worksheet path: `149352024 B/op, 1502160 allocs/op -> 136761256 B/op, 873681 allocs/op`
- repeat-aware serial pair rerun was slightly positive:
  - experiment pair:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-shared-plaincell-pair-serial.json`
    - `00012389.xlsx`: `1758 ms`
    - `testRecordSizeExceeded.xlsx`: `2806 ms`
    - total: `4564 ms`
  - retained baseline pair:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-cap-pair-serial.json`
    - total: `4574 ms`
- but the decisive 21-file `.xlsx` keyset regressed clearly:
  - retained baseline:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-cap-xlsx-keyset-serial.json`
    - `.xlsx millis = 20827`
  - experiment:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-shared-plaincell-xlsx-keyset-serial.json`
    - `.xlsx millis = 23712`
  - parity remained clean:
    - `textBytes = 279385415`
    - `images = 1`
    - `ok = 21`

Interpretation:
- this is the closest shared-string markdown-only candidate so far:
  - it improved the local shared-string benches strongly
  - it even passed the representative serial pair
- but the broader retained `.xlsx` keyset still moved the wrong way by a wide margin
- under the current retention bar, that broader regression outweighs the narrow wins

Decision:
- Reverted.

## 2026-07-06 rejected: one-pass marker probe in `maybeInlineHiddenOfficeReferenceMarker(...)`

- experiment:
  - rewrite only the first probe stage of `maybeInlineHiddenOfficeReferenceMarker(...)`
  - replace the initial cluster of:
    - `IndexByte('=')`
    - `IndexByte(':')`
    - `IndexByte('/')` / `IndexByte('\\')`
    - `ContainsAny("<>{}[]()")`
    - `IndexByte('%')`
    - unconditional `containsRIDFold(...)`
    - unconditional `containsASCIIFold(..., "url(")`
  - with one byte scan that records:
    - marker families present
    - cheap `rid` / `url` hints
  - keep every downstream marker-family check and acceptance rule unchanged
- rationale:
  - the new shared-string markdown clean benchmark and profile showed that:
    - `cleanMarkdownTableCellValue(...)`
    - `stripInlineHiddenOfficeReferences(...)`
    - `maybeInlineHiddenOfficeReferenceMarker(...)`
    still contribute visible cost on the `00012389.xlsx` markdown-clean path
  - this looked like a very small helper-only attempt to reduce repeated precheck scans without
    changing the actual hidden-reference filtering logic

Validation:
- targeted regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run 'TestCleanMarkdownTableCellValue|TestPrepareMarkdownTableCellValueSingleLineFastPath|TestExtractXLSX|TestExtractWorksheet|Test.*Text|Test.*Visible' ./`
- shared-string markdown-stage benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXSharedStringMarkdownClean00012389$|BenchmarkXLSXSharedStringMarkdownPrepare00012389$|BenchmarkXLSXSharedStringPrepared00012389$|BenchmarkXLSXSharedStringPreparedNoMarkdown00012389$' -benchmem -benchtime=1x ./`
- integrated shared-string hotspot checks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkExtractXLSXHotspots/00012389.xlsx$' -benchmem -benchtime=1x ./`
- repeat-aware serial pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-inline-marker-onepass-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-inline-marker-onepass-pair-serial.csv testdata/web-samples/samples/xlsx/00012389.xlsx testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- targeted regression stayed green before revert:
  - `ok officeread 35.379s`
- shared-string markdown-focused benchmarks improved materially:
  - `BenchmarkXLSXSharedStringPrepared00012389`: `675977600 ns/op -> 610426500 ns/op`
  - `BenchmarkXLSXSharedStringPreparedNoMarkdown00012389`: `281461300 ns/op -> 260723900 ns/op`
  - `BenchmarkXLSXSharedStringMarkdownClean00012389`: `519055600 ns/op -> 273243800 ns/op`
  - `BenchmarkXLSXSharedStringMarkdownPrepare00012389`: `144604600 ns/op -> 85533100 ns/op`
- integrated `00012389.xlsx` hotspot checks also looked positive:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `1666754000 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `866545500 ns/op`
- but the decisive serial pair rerun regressed badly because the simple-inline representative moved
  the wrong way:
  - experiment pair:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-inline-marker-onepass-pair-serial.json`
    - `00012389.xlsx`: `1662 ms`
    - `testRecordSizeExceeded.xlsx`: `3786 ms`
    - total: `5448 ms`
  - retained baseline pair:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-cap-pair-serial.json`
    - total: `4574 ms`
    - `00012389.xlsx`: `1758 ms`
    - `testRecordSizeExceeded.xlsx`: `2816 ms`
- output parity stayed clean on the pair:
  - `textBytes = 192482085`
  - `images = 0`
  - `ok = 2`

Interpretation:
- this helper-only change genuinely helped the shared-string markdown path in isolation
- but it still damaged the broader retained representative pair badly enough to fail the keep bar
- the result reinforces the current lesson from this repo: even apparently local marker-probe
  rewrites can perturb the simple-inline workload in ways the isolated shared-string benches do not
  predict

Decision:
- Reverted.

## 2026-07-06 retained: shared-string micro-benchmark harness

- change:
  - extend `extract_bench_test.go` with shared-string-specific benchmarks for `00012389.xlsx`:
    - `BenchmarkXLSXSharedStringCandidate00012389`
    - `BenchmarkXLSXSharedStringPrepared00012389`
  - add a shared helper that loads the first visible worksheet plus `sharedStrings.xml`, then skips
    cleanly if the sample does not hit the shared-string fast path
- rationale:
  - several recent experiments showed the same failure mode:
    - focused end-to-end worksheet benchmarks looked attractive
    - repeat-aware pair / `.xlsx` keyset reruns then moved the wrong way
  - the existing benchmark harness did not isolate the two most important remaining shared-string
    components:
    - candidate admission cost
    - prepared shared-string worksheet extraction cost
  - separating those costs makes the next round of profiling and experiment triage much more direct

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract_bench_test.go`
- benchmark compilation / execution:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXSharedStringCandidate00012389$|BenchmarkXLSXSharedStringPrepared00012389$' -benchmem -benchtime=1x ./`

Observed results:
- `BenchmarkXLSXSharedStringCandidate00012389`:
  - `257169300 ns/op`
  - `2048 B/op`
  - `1 allocs/op`
- `BenchmarkXLSXSharedStringPrepared00012389`:
  - `631530600 ns/op`
  - `149341400 B/op`
  - `1502185 allocs/op`

Interpretation:
- the shared-string prepared extraction path remains materially larger than the candidate gate, but
  the candidate still costs enough to justify a dedicated benchmark when testing future admission
  changes
- this harness is retained as measurement infrastructure; it does not change user-visible behavior

Decision:
- Retained.

## 2026-07-06 retained: shared-string markdown-stage micro-benchmark harness

- change:
  - extend `extract_bench_test.go` with two more shared-string markdown-stage benchmarks for
    `00012389.xlsx`:
    - `BenchmarkXLSXSharedStringMarkdownClean00012389`
    - `BenchmarkXLSXSharedStringMarkdownPrepare00012389`
  - add a worksheet-scoped collector that gathers the markdown-eligible shared-string cell values
    from the first visible worksheet so the clean / prepare stages can be benchmarked separately
- rationale:
  - the previous harness split:
    - candidate admission
    - prepared extraction with markdown
    - prepared extraction without markdown
  - but it still did not isolate the two markdown-stage functions that keep showing up in profiles:
    - `cleanMarkdownTableCellValue(...)`
    - `prepareMarkdownTableCellValue(...)`
  - separating those stages makes the next markdown-path experiments much less guessy

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract_bench_test.go`
- benchmark execution:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXSharedStringCandidate00012389$|BenchmarkXLSXSharedStringPrepared00012389$|BenchmarkXLSXSharedStringPreparedNoMarkdown00012389$|BenchmarkXLSXSharedStringMarkdownClean00012389$|BenchmarkXLSXSharedStringMarkdownPrepare00012389$' -benchmem -benchtime=1x ./`
- stage-specific profile capture:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXSharedStringMarkdownClean00012389$' -benchtime=1x -cpuprofile testdata/web-samples/reports/bench-sharedstring-markdown-clean-00012389.pprof ./`
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXSharedStringMarkdownPrepare00012389$' -benchtime=1x -cpuprofile testdata/web-samples/reports/bench-sharedstring-markdown-prepare-00012389.pprof ./`

Observed results:
- `BenchmarkXLSXSharedStringCandidate00012389`:
  - `268501200 ns/op`
  - `2048 B/op`
  - `1 allocs/op`
- `BenchmarkXLSXSharedStringPrepared00012389`:
  - `675977600 ns/op`
  - `149155400 B/op`
  - `1502111 allocs/op`
- `BenchmarkXLSXSharedStringPreparedNoMarkdown00012389`:
  - `281461300 ns/op`
  - `50352968 B/op`
  - `646837 allocs/op`
- `BenchmarkXLSXSharedStringMarkdownClean00012389`:
  - `519055600 ns/op`
  - `7541400 B/op`
  - `113020 allocs/op`
  - rerun for profile capture:
    - `307114300 ns/op`
    - `7414512 B/op`
    - `112975 allocs/op`
- `BenchmarkXLSXSharedStringMarkdownPrepare00012389`:
  - `144604600 ns/op`
  - `13221160 B/op`
  - `650347 allocs/op`
  - rerun for profile capture:
    - `73205600 ns/op`
    - `13221160 B/op`
    - `650347 allocs/op`
- profile files:
  - `testdata/web-samples/reports/bench-sharedstring-markdown-clean-00012389.pprof`
  - `testdata/web-samples/reports/bench-sharedstring-markdown-prepare-00012389.pprof`

Interpretation:
- the shared-string markdown side is now decomposed enough to state the current priority clearly:
  - `cleanMarkdownTableCellValue(...)` is materially heavier than
    `prepareMarkdownTableCellValue(...)`
  - the no-markdown split plus the new stage splits point to markdown cleaning as the more plausible
    remaining shared-string target
- this retained harness does not change behavior; it sharpens the evidence for the next round

Decision:
- Retained.

## 2026-07-06 retained: shared-string no-markdown micro-benchmark harness

- change:
  - extend `extract_bench_test.go` again with:
    - `BenchmarkXLSXSharedStringPreparedNoMarkdown00012389`
  - this benchmark reuses the shared-string worksheet loader and runs
    `appendSharedStringWorksheetTextPrepared(...)` with `md=nil`, so the prepared extraction path can
    be measured without markdown cleanup / table preparation mixed in
- rationale:
  - the first shared-string harness split candidate admission from prepared extraction, but the
    prepared benchmark still included markdown cleanup and markdown row building
  - recent profiles suggested the remaining shared-string work was a blend of:
    - candidate admission
    - plain text append / clean path
    - markdown cleanup and markdown row preparation
  - adding a no-markdown benchmark makes that split explicit and gives future experiments a cleaner
    signal for whether a change actually helps text extraction or only the markdown side

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract_bench_test.go`
- benchmark execution:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXSharedStringCandidate00012389$|BenchmarkXLSXSharedStringPrepared00012389$|BenchmarkXLSXSharedStringPreparedNoMarkdown00012389$' -benchmem -benchtime=1x ./`
- no-markdown profile capture:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXSharedStringPreparedNoMarkdown00012389$' -benchtime=1x -cpuprofile testdata/web-samples/reports/bench-sharedstring-prepared-nomd-00012389.pprof ./`

Observed results:
- `BenchmarkXLSXSharedStringCandidate00012389`:
  - `258142000 ns/op`
  - `2048 B/op`
  - `1 allocs/op`
- `BenchmarkXLSXSharedStringPrepared00012389`:
  - `627688300 ns/op`
  - `149498160 B/op`
  - `1502217 allocs/op`
- `BenchmarkXLSXSharedStringPreparedNoMarkdown00012389`:
  - `243425100 ns/op`
  - `50390200 B/op`
  - `646844 allocs/op`
  - rerun for profile capture:
    - `244445600 ns/op`
    - `50434536 B/op`
    - `646855 allocs/op`
- no-markdown profile file:
  - `testdata/web-samples/reports/bench-sharedstring-prepared-nomd-00012389.pprof`

Interpretation:
- the shared-string prepared path is now split into three measurable layers:
  - candidate gate: about `258 ms`
  - prepared extraction with markdown: about `628 ms`
  - prepared extraction without markdown: about `244 ms`
- this shows the dominant remaining shared-string cost is still on the markdown side rather than the
  plain text append path alone
- the no-markdown profile also confirms that once markdown is removed, the biggest remaining items
  are still candidate admission and the text-clean / append path, not worksheet-cell parsing by
  itself

Decision:
- Retained.

## 2026-07-06 rejected: shared-string candidate one-pass tag scan

- experiment:
  - rewrite `sharedStringWorksheetCandidate(...)` to use one sequential XML-tag scan over the
    worksheet bytes
  - keep only the existing `DOCTYPE` / `CDATA` whole-file guards
  - replace the long chain of whole-file `bytes.Contains(...)` checks with per-tag rejection while
    scanning:
    - reject tags with local names `f`, `is`, `t`, `rPh`, `hyperlink`, `dataValidation`,
      `oddHeader`, `oddFooter`, `evenHeader`, `evenFooter`, `firstHeader`, `firstFooter`
    - reject any start tag carrying `hidden="..."` / `hidden='...'`
    - keep the same `<c ... t="...">` admission rules for shared-string cells
- rationale:
  - profile evidence on `00012389.xlsx` still showed `sharedStringWorksheetCandidate(...)` at about
    `17.84%` cumulative CPU, with much of that attributed to repeated whole-file `bytes.Contains`
    passes
  - this looked like the most direct way to reduce full-worksheet pre-scan cost without changing
    extraction or markdown semantics

Validation:
- targeted regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run 'TestExtractXLSX|TestExtractWorksheet|TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue|Test.*Text|Test.*Visible' ./`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkExtractXLSXHotspots/00012389.xlsx$' -benchmem -benchtime=1x ./`
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- repeat-aware serial pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-sharedscan-onepass-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-sharedscan-onepass-pair-serial.csv testdata/web-samples/samples/xlsx/00012389.xlsx testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- targeted regression stayed green before revert:
  - `ok officeread 33.222s`
- focused hotspot benchmarks improved on the shared-string representative sample:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: about `1762633400 ns/op -> 1568831100 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: about `937539700 ns/op -> 753659900 ns/op`
- the simple-inline representative stayed roughly flat to slightly worse:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: about `2991825000 ns/op -> 3006932600 ns/op`
- but the decisive serial pair rerun still regressed against the retained baseline:
  - retained baseline pair:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-cap-pair-serial.json`
    - total: `4574 ms`
    - `00012389.xlsx`: `1758 ms`
    - `testRecordSizeExceeded.xlsx`: `2816 ms`
  - experiment pair:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-sharedscan-onepass-pair-serial.json`
    - total: `4720 ms`
    - `00012389.xlsx`: `1527 ms`
    - `testRecordSizeExceeded.xlsx`: `3193 ms`
- output parity stayed clean on the pair:
  - `textBytes = 192482085`
  - `images = 0`
  - `ok = 2`

Interpretation:
- the one-pass candidate scan did speed up the shared-string-heavy representative, but it
  simultaneously dragged the simple-inline representative enough to lose on the same retained pair
- under the current evidence standard, that tradeoff is not good enough to keep

Decision:
- Reverted.

## 2026-07-06 rejected: shared-string markdown row lazy allocation

- experiment:
  - make `appendSharedStringWorksheetTextPrepared(...)` allocate `markdownRowValues` lazily
  - instead of preallocating `make([]string, 0, 16)` for every markdown-eligible row, keep it `nil`
    until the row actually produces a non-empty prepared markdown cell
  - when the first prepared cell appears, allocate once with `max(16, cellCol)` capacity and then
    keep the existing row build / compact logic unchanged
- rationale:
  - the shared-string worksheet path on `00012389.xlsx` still spends meaningful time in markdown
    cell cleanup / preparation, and the function was also paying one slice allocation per eligible
    row even if a row never emitted a markdown cell
  - this looked like a conservative structural allocation reduction that did not alter text
    extraction, markdown cleanup semantics, or dedupe behavior

Validation:
- targeted regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run 'TestExtractXLSX|TestExtractWorksheet|TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue|Test.*Text|Test.*Visible' ./`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkExtractXLSXHotspots/00012389.xlsx$' -benchmem -benchtime=1x ./`
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- repeat-aware serial pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-shared-mdrow-lazy-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-shared-mdrow-lazy-pair-serial.csv testdata/web-samples/samples/xlsx/00012389.xlsx testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`
- decisive `.xlsx` keyset rerun on the retained 21-file input set:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata/web-samples/reports/perf-exp-ai-assistant-shared-mdrow-lazy-xlsx-keyset-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-shared-mdrow-lazy-xlsx-keyset-serial.csv <same-21-xlsx-inputs>`

Observed results:
- targeted regression stayed green before revert:
  - `ok officeread 35.153s`
- focused hotspot benchmarks were only modestly positive:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: about `1762633400 ns/op -> 1756117000 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: about `937539700 ns/op -> 910847800 ns/op`
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: about `2991825000 ns/op -> 2948654700 ns/op`
  - allocations on the shared-string hotspot actually drifted upward:
    - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `618263952 B/op, 3266656 allocs/op -> 631475480 B/op, 3317692 allocs/op`
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `149075128 B/op, 1502096 allocs/op -> 157178048 B/op, 1533465 allocs/op`
- repeat-aware serial pair rerun looked positive:
  - retained baseline pair:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-cap-pair-serial.json`
    - total: `4574 ms`
  - experiment pair:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-shared-mdrow-lazy-pair-serial.json`
    - total: `4532 ms`
    - `00012389.xlsx`: `1713 ms`
    - `testRecordSizeExceeded.xlsx`: `2819 ms`
- but the decisive 21-file `.xlsx` keyset regressed clearly against the retained baseline:
  - retained baseline:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-cap-xlsx-keyset-serial.json`
    - `.xlsx millis = 20827`
  - experiment:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-shared-mdrow-lazy-xlsx-keyset-serial.json`
    - `.xlsx millis = 23446`
  - parity remained clean:
    - `textBytes = 279385415`
    - `images = 1`
    - `ok = 21`

Interpretation:
- lazy row allocation produced a small representative-pair win, but it did not survive the broader
  retained `.xlsx` keyset
- the upward allocation drift in the focused shared-string benchmark was an early warning, and the
  keyset rerun confirmed that this structural change moved the full workload in the wrong direction

Decision:
- Reverted.

## 2026-07-06 rejected: maybeControlFragmentText helper gates

- experiment:
  - add cheap first-byte / marker gates before the heavier helper calls inside
    `maybeControlFragmentText(...)`
  - narrow `looksLikeOLEIdentifierFragment(...)` to strings whose first byte and punctuation profile
    make a GUID / assignment shape plausible
  - narrow `looksLikeOLEWrapperStreamName(...)` to `c/C/o/O` starts
  - narrow `looksLikeOOXMLMarkupNameFragment(...)` to `r/c/p/t` starts or wrapper-style leading
    punctuation
- rationale:
  - fresh profile inspection on `profile-current-00012389.pprof` showed
    `cleanTextFastPathControlFragment(...) -> maybeControlFragmentText(...)` still consuming a
    meaningful share of the remaining shared-string worksheet cost
  - within that stack, the most expensive inner checks were:
    - `looksLikeOLEIdentifierFragment(...)`
    - `looksLikeOOXMLMarkupNameFragment(...)`
    - `looksLikeOLEWrapperStreamName(...)`
  - the goal was to avoid paying those helpers for obviously ordinary short text values while
    keeping the same accepted / rejected control-fragment set

Validation:
- targeted text / worksheet regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run 'Test.*Text|Test.*OLE|Test.*Noise|Test.*Legacy|Test.*Unicode|Test.*XLSX|Test.*Xlsx|Test.*Worksheet|Test.*Markdown|Test.*Visible' -count=1 .`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkExtractXLSXHotspots/00012389.xlsx$' -benchmem -benchtime=1x ./`
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- repeat-aware serial pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-control-gates-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-control-gates-pair-serial.csv testdata/web-samples/samples/xlsx/00012389.xlsx testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- targeted regression stayed green before revert:
  - `ok officeread 35.219s`
- focused hotspot benchmarks improved strongly in isolation:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: about `1762633400 ns/op -> 1517839100 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: about `937539700 ns/op -> 752552800 ns/op`
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: about `2991825000 ns/op -> 2591751800 ns/op`
- but the decisive serial pair rerun regressed badly against the retained baseline:
  - retained baseline pair:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-cap-pair-serial.json`
    - total: `4574 ms`
    - `00012389.xlsx`: `1758 ms`
    - `testRecordSizeExceeded.xlsx`: `2816 ms`
  - experiment pair:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-control-gates-pair-serial.json`
    - total: `5537 ms`
    - `00012389.xlsx`: `2012 ms`
    - `testRecordSizeExceeded.xlsx`: `3525 ms`
- output parity stayed clean on the pair:
  - `textBytes = 192482085`
  - `images = 0`
  - `ok = 2`

Interpretation:
- the helper gating made the focused benchmark shape look much better, but it did not hold under the
  repeat-aware end-to-end pair rerun
- this reinforces the current working rule for this repo: isolated cleanText / control-fragment wins
  are not trustworthy unless the representative serial pair also moves the right way

Decision:
- Reverted.

## 2026-07-06 rejected: cleanTextFastPath one-pass filter scan

- experiment:
  - collapse the retained `cleanTextFastPath(...)` prechecks into the main byte loop
  - remove the separate `strings.ContainsAny(s, "/\\%[]()")` and `containsRIDFold(s)` scans
  - reject the same marker families inline while walking the already-trimmed string once:
    - path / bracket / percent markers
    - `rId` / `RID` relationship-id fragments
    - existing backslash / angle / `#` / `_xNNNN_` guards
- rationale:
  - fresh 2026-07-06 profiles still showed `cleanTextFastPath(...)` / `containsRIDFold(...)` in the
    remaining `.xlsx` hotspot mix
  - this looked like a conservative way to reduce repeated scans on the hottest worksheet text path
    without changing any downstream text-cleaning behavior

Validation:
- targeted text / worksheet regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run 'Test.*Text|Test.*OLE|Test.*Noise|Test.*Legacy|Test.*Unicode|Test.*XLSX|Test.*Xlsx|Test.*Worksheet|Test.*Markdown|Test.*Visible' -count=1 .`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkExtractXLSXHotspots/00012389.xlsx$' -benchmem -benchtime=1x ./`
  - `& 'C:\Program Files\Go\bin\go.exe' test -buildvcs=false -run '^$' -bench 'BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- repeat-aware serial pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-cleantext-onepass-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-cleantext-onepass-pair-serial.csv testdata/web-samples/samples/xlsx/00012389.xlsx testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- targeted regression stayed green before revert:
  - `ok officeread 37.362s`
- focused hotspot benchmarks looked directionally positive versus the current retained tree:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: about `1762633400 ns/op -> 1702512100 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: about `937539700 ns/op -> 904801100 ns/op`
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: about `2991825000 ns/op -> 2891787100 ns/op`
- but the decisive serial pair rerun regressed against the retained baseline:
  - retained baseline pair:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-cap-pair-serial.json`
    - total: `4574 ms`
    - `00012389.xlsx`: `1758 ms`
    - `testRecordSizeExceeded.xlsx`: `2816 ms`
  - experiment pair:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-cleantext-onepass-pair-serial.json`
    - total: `4894 ms`
    - `00012389.xlsx`: `1804 ms`
    - `testRecordSizeExceeded.xlsx`: `3090 ms`
- output parity stayed clean on the pair:
  - `textBytes = 192482085`
  - `images = 0`
  - `ok = 2`

Interpretation:
- collapsing the scans helped the isolated benchmark view but did not survive the repeat-aware
  end-to-end pair rerun
- this is the same pattern as several earlier attempts: a plausible micro-level reduction around
  `cleanTextFastPath(...)` still lost to the real integrated worksheet workload

Decision:
- Reverted.

## 2026-07-06 rejected: parse local tag name once inside `simpleInlineWorksheetCandidate(...)`

- experiment:
  - change `simpleInlineWorksheetCandidate(...)` so each XML tag computes its local start-tag name
    once and then reuses that result for the allowed/disallowed element checks
  - introduce no-allocation byte-slice helpers for:
    - extracting the local tag name from the current tag
    - comparing that local name against the candidate allow/deny list
    - checking header/footer element membership without converting to `string`
- rationale:
  - the current profile still showed `simpleInlineWorksheetCandidate(...)` as a major cost center
  - the existing implementation was reparsing the same tag repeatedly through many
    `xmlStartTagNameIs(...)` calls, which looked like avoidable repeated work

Validation:
- targeted tests:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'TestExtractXLSX|TestExtractWorksheet|TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue' ./`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineTextOnlyHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkExtractXLSXHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- repeat-aware pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-candidate-localname-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-candidate-localname-pair-serial.csv testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata/web-samples/samples/xlsx/00012389.xlsx`

Observed results:
- focused hotspot view was mixed and initially looked plausible on full extract:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `3079023600 ns/op -> 2621251100 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2295981300 ns/op -> 2440455600 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1229985100 ns/op -> 1277236700 ns/op`
- but the repeat-aware pair rerun regressed badly:
  - `00012389.xlsx`: `1758 ms -> 3746 ms`
  - `testRecordSizeExceeded.xlsx`: `2816 ms -> 4105 ms`
  - parity still held: `text=11148655 / 181333430`, `images=0`

Interpretation:
- despite the attractive single-benchmark full-extract number, this local-name rewrite changed the
  real candidate-scan balance in the wrong direction on both representative `.xlsx` files
- under the current bar, the pair rerun is decisive enough to reject without broader keyset work

Decision:
- Reverted.

## 2026-07-06 rejected: `r/R` pre-gate before `containsRIDFold(...)` in `cleanTextFastPath(...)`

- experiment:
  - keep the existing `cleanTextFastPath(...)` structure and control-fragment ordering unchanged
  - only narrow the `rid` rejection check so `containsRIDFold(...)` runs when the value contains
    at least one `'r'` or `'R'`
  - implementation shape:
    - keep `strings.ContainsAny(s, "/\\%[]()")` as-is
    - replace unconditional `containsRIDFold(s)` with
      `(strings.IndexByte(s, 'r') >= 0 || strings.IndexByte(s, 'R') >= 0) && containsRIDFold(s)`
- rationale:
  - the current CPU profile on the retained tree still showed `containsRIDFold(...)` as a visible
    flat cost inside the simple-inline fast-path common case
  - this was intended as a much narrower alternative to the previously rejected merged single-loop
    `cleanTextFastPath(...)` rewrite

Validation:
- targeted tests:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'TestExtractXLSX|TestExtractWorksheet|TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue' ./`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineTextOnlyHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkExtractXLSXHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- repeat-aware pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-ridgate-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-ridgate-pair-serial.csv testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata/web-samples/samples/xlsx/00012389.xlsx`

Observed results:
- focused hotspot evidence stayed mixed:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `3079023600 ns/op -> 2897673000 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2295981300 ns/op -> 2319897600 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1229985100 ns/op -> 1221459600 ns/op`
- but the repeat-aware pair rerun regressed badly:
  - `00012389.xlsx`: `1758 ms -> 2310 ms`
  - `testRecordSizeExceeded.xlsx`: `2816 ms -> 3888 ms`
  - parity still held: `text=11148655 / 181333430`, `images=0`

Interpretation:
- the extra `IndexByte` gate looked cheap in isolation, but it changed the real workload balance in
  the wrong direction on both representative `.xlsx` files
- under the current evidence bar, this is a clear no-keep candidate and does not need broader
  keyset validation

Decision:
- Reverted.

## 2026-07-06 rejected: skip final trim in `appendWorksheetValue(...)`

- experiment:
  - keep `appendWorksheetValue(...)` behavior the same up through `cleanText(...)`
  - after `cleanText(...)` returns a non-empty value, append it directly with
    `appendTrimmedTextBlock(...)` instead of routing through `appendCleanedTextBlock(...)`
  - rationale: `cleanText(...)` already returns a trimmed string on both the fast and slow paths, so
    the extra `appendCleanedTextBlock(...)` call looked like a redundant `strings.TrimSpace(...)`
    pass in a very hot text path

Validation:
- targeted tests:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'TestExtractXLSX|TestExtractWorksheet|TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue' ./`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineTextOnlyHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkExtractXLSXHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- repeat-aware pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-appendtrimmed-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-appendtrimmed-pair-serial.csv testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata/web-samples/samples/xlsx/00012389.xlsx`
- decisive `.xlsx` keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-appendtrimmed-xlsx-keyset-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-appendtrimmed-xlsx-keyset-serial.csv <xlsx-keyset-21>`
- parity / regression:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata/web-samples/reports/compat-xlsx-20260706-after-appendtrimmed.json <xlsx-keyset-21>`
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused hotspot evidence was consistently favorable:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `3079023600 ns/op -> 2979648700 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2295981300 ns/op -> 2056557600 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1229985100 ns/op -> 1179032700 ns/op`
- repeat-aware pair rerun also improved:
  - `00012389.xlsx`: `1758 ms -> 1690 ms`
  - `testRecordSizeExceeded.xlsx`: `2816 ms -> 2786 ms`
- but the decisive 21-file `.xlsx` keyset rerun still regressed against the current retained tree:
  - current retained tree: `20827 ms`
  - experiment: `21598 ms`
  - parity held: `21/21 ok`, `textBytes=279385415`, `images=1`
- repository regression stayed green before revert:
  - `ok officeread 141.393s`

Interpretation:
- this is another case where a very plausible local cleanup improves both the hotspot bench and the
  representative pair rerun
- but it still loses to the current retained tree on the decisive 21-file `.xlsx` workload, so it
  does not clear the present retention bar

Decision:
- Reverted.

## 2026-07-06 retained: stop simple-inline markdown coordinate parsing after markdown row cap

- change:
  - in `appendSimpleInlineWorksheetTextPrepared(...)`, keep collecting worksheet markdown only until
    `maxMarkdownTableRows` is actually reached
  - once the capped markdown rows are flushed, switch the remaining simple-inline worksheet scan to
    text-only handling:
    - skip `r=` attribute extraction
    - skip `cellRefIndexes(...)`
    - skip markdown cell cleanup / preparation
    - keep visible text extraction unchanged
- rationale:
  - `testRecordSizeExceeded.xlsx` is a `200000`-row simple-inline worksheet, but markdown retains at
    most `50000` rows
  - after the first `50000` visible rows, continuing to parse per-cell coordinates only serves the
    discarded markdown tail, not the extracted text

Validation:
- targeted tests:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'TestExtractXLSX|TestExtractWorksheet|TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue' ./`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkXLSXSimpleInlineTextOnlyHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkExtractXLSXHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- repeat-aware pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-cap-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-cap-pair-serial.csv testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata/web-samples/samples/xlsx/00012389.xlsx`
- decisive `.xlsx` keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-cap-xlsx-keyset-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-cap-xlsx-keyset-serial.csv <xlsx-keyset-21>`
- parity / regression:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata/web-samples/reports/compat-xlsx-20260706-after-simpleinline-cap.json <xlsx-keyset-21>`
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused hotspot view was mixed but net-positive on the most relevant integrated paths:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: about `3130030700 ns/op -> 3079023600 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: about `1931084000 ns/op -> 2295981300 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: about `1423902100 ns/op -> 1229985100 ns/op`
  - allocs dropped materially on the integrated extract path:
    - `2249931048 B/op`, `12052922 allocs/op`
    - `-> 1974691216 B/op`, `7602921 allocs/op`
- repeat-aware pair rerun improved strongly with parity preserved:
  - `00012389.xlsx`: `2056 ms -> 1758 ms`
  - `testRecordSizeExceeded.xlsx`: `6162 ms -> 2816 ms`
- decisive 21-file `.xlsx` keyset rerun improved clearly with identical output totals:
  - retained prior keyset total: `23930 ms`
  - current candidate total: `20827 ms`
  - parity held: `textBytes=279385415`, `images=1`, `21/21 ok`
- repository regression stayed green:
  - `ok officeread 140.403s`

Interpretation:
- this optimization specifically removes useless per-cell coordinate work from the discarded
  markdown tail of very large simple-inline worksheets
- although one local markdown-heavy benchmark regressed, the repeat-aware end-to-end evidence on the
  actual retained `.xlsx` workload improved decisively, and parity held throughout

Decision:
- Retained.

## 2026-07-06 rejected: byte-write plain `<t>` segments in `simpleInlineCellText(...)`

- experiment:
  - keep the current `simpleInlineCellText(...)` scan shape, but avoid per-segment string
    allocation when a `<t>...</t>` fragment contains no XML entity
  - for plain fragments, write the raw `[]byte` slice directly into the `strings.Builder`
  - only materialize a Go string and call `html.UnescapeString(...)` for segments that actually
    contain `&`
- rationale:
  - after the retained markdown-cap optimization, the remaining `testRecordSizeExceeded.xlsx`
    hotspot still spends a large share of time in `simpleInlineCellText(...)`
  - the dominant simple-inline worksheet text appears to be plain text, so skipping one short-lived
    string allocation per `<t>` fragment looked like a plausible allocation win

Validation:
- targeted tests:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'TestExtractXLSX|TestExtractWorksheet|TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue' ./`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineTextOnlyHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkExtractXLSXHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- repeat-aware pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-bytewrite-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-bytewrite-pair-serial.csv testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata/web-samples/samples/xlsx/00012389.xlsx`
- decisive `.xlsx` keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-bytewrite-xlsx-keyset-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-bytewrite-xlsx-keyset-serial.csv <xlsx-keyset-21>`
- parity / regression:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -json testdata/web-samples/reports/compat-xlsx-20260706-after-simpleinline-bytewrite.json <xlsx-keyset-21>`
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused hotspot evidence was mixed:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `3079023600 ns/op -> 2910727000 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2295981300 ns/op -> 2482575400 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1229985100 ns/op -> 1358590100 ns/op`
  - allocs improved materially on the hotspot:
    - `1974691216 B/op`, `7602921 allocs/op`
    - `-> 1782686208 B/op`, `4602916 allocs/op`
- repeat-aware pair rerun was slightly favorable:
  - `00012389.xlsx`: `1758 ms -> 1753 ms`
  - `testRecordSizeExceeded.xlsx`: `2816 ms -> 2675 ms`
- decisive 21-file `.xlsx` keyset rerun regressed against the current retained tree even though it
  still beat the older pre-cap baseline:
  - current retained tree: `20827 ms`
  - experiment: `21939 ms`
  - parity held: `21/21 ok`, `textBytes=279385415`, `images=1`
- repository regression stayed green before revert:
  - `ok officeread 153.088s`

Interpretation:
- the byte-write change clearly reduced allocations and slightly helped the two-file rerun
- but the decisive 21-file `.xlsx` workload moved in the wrong direction versus the current retained
  tree, so this is not keepable under the present bar

Decision:
- Reverted.

## 2026-07-06 rejected: plain single-`<t>` direct read for simple-inline worksheet cells

- experiment:
  - add a worksheet-level guard for highly regular inline-string worksheets:
    - no `&` entities
    - no rich-text `<r>` runs
    - no phonetic tags
  - when the guard matches, try to read each cell's text by directly slicing the first `<t>...</t>`
    payload instead of building it through `simpleInlineCellText(...)`
- rationale:
  - fresh shape evidence on `testRecordSizeExceeded.xlsx` showed the dominant worksheet is perfectly
    regular for this path:
    - `cells=3000000`
    - `inline=3000000`
    - `tTags=3000000`
    - `richRuns=0`
    - `entityAmp=0`
    - `rows=200000`
  - that made the generic `<t>` scan loop plus `strings.Builder` inside `simpleInlineCellText(...)`
    look like avoidable per-cell overhead

Validation:
- correctness:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'TestExtractXLSX|TestExtractWorksheet|TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue' ./`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkXLSXSimpleInlineTextOnlyHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkExtractXLSXHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- serial focused pair rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-plaincell-pair-serial.json`
- serial `.xlsx` keyset rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-plaincell-xlsx-keyset-serial.json`
- repository regression after revert:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- correctness stayed green before revert:
  - `ok officeread 4.887s`
- focused benchmark signal regressed sharply:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `3130030700 ns/op -> 3192521100 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `1931084000 ns/op -> 2505165400 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1423902100 ns/op -> 1649195800 ns/op`
- focused pair rerun looked superficially promising:
  - `00012389.xlsx`: `1858 ms`
  - `testRecordSizeExceeded.xlsx`: `3046 ms`
  - pair total: `4904 ms`
- but the broader retained `.xlsx` keyset did not improve enough to justify keeping it:
  - retained shared-scan `.xlsx` keyset baseline:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-sharedscan-xlsx-keyset-serial.json`
    - total: `23930 ms`
  - experiment:
    - `testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-plaincell-xlsx-keyset-serial.json`
    - total: `24007 ms`
  - keyset parity still held:
    - `textBytes=279385415`
    - `images=1`
    - `ok=21`

Interpretation:
- this is another case where a local intuition about a very uniform XML shape was correct, but the
  integrated retained keyset still failed to beat the current baseline
- because the focused benchmark signal was also clearly worse, the pair rerun is best treated as an
  encouraging but noisy outlier rather than enough evidence to keep the change

Decision:
- Reverted.

## 2026-07-06 rejected: conditional markdown-cell prepare escapes / RTF prefix check in `.xlsx`

- experiment:
  - narrow `prepareMarkdownTableCellValue(...)` so the single-line fast path only performs
    backslash and pipe escaping when the normalized cell text actually contains `\` or `|`
  - replace the unconditional `strings.ToLower(...)` + `HasPrefix(...)` RTF guard with a small
    ASCII case-folded prefix helper
- rationale:
  - current `00012389.xlsx` worksheet profiling still shows markdown-cell preparation on the hot
    path after `dec.RawToken()`
  - the active worksheet shape evidence gathered earlier in the session showed tens of thousands of
    single-line shared-string markdown cells, almost all plain text, so the unconditional escaping
    and lowercasing looked like reusable duplicate work

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused correctness:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue' ./`
- focused serial hotspot bench:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`

Observed results:
- correctness stayed green before revert:
  - `ok officeread 2.967s`
- the decisive serial hotspot bench regressed instead of improving:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2334292400 ns/op -> 2302511300 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2002855100 ns/op -> 2595168900 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1367983300 ns/op -> 1799955900 ns/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`: `832965300 ns/op -> 1168554700 ns/op`
- the allocation picture improved in places, but wall time clearly moved the wrong way on the
  retained `.xlsx` hotspots

Interpretation:
- even a very narrow “only escape when needed” / “only fold case when needed” change in markdown
  cell preparation did not survive focused retained hotspot measurement
- this is another example where helper-level duplicate-work intuition is not enough by itself on
  the `.xlsx` path; the surrounding execution balance is sensitive enough that the change should
  not be kept

Decision:
- Reverted.

## 2026-07-06 note: reverted-baseline serial `.xlsx` pair rerun is slower under current machine state

- validation:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-current-xlsx-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-current-xlsx-pair-serial.csv testdata/web-samples/samples/xlsx/00012389.xlsx testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx`

Observed results:
- reverted current code still preserved output parity:
  - `00012389.xlsx`: `ok=true`, `text=11148655`, `images=0`
  - `testRecordSizeExceeded.xlsx`: `ok=true`, `text=181333430`, `images=0`
- current run times were materially slower than the retained clean baseline recorded earlier in the
  session:
  - current rerun total: `6747 ms`
  - current per-file:
    - `00012389.xlsx`: `2791 ms` with `runs=[2211 2791 3162]`
    - `testRecordSizeExceeded.xlsx`: `3956 ms` with `runs=[4450 3956 3899]`
  - retained clean baseline total: `4933 ms`
  - retained clean baseline per-file:
    - `00012389.xlsx`: `2037 ms`
    - `testRecordSizeExceeded.xlsx`: `2896 ms`

Interpretation:
- because this rerun used the reverted retained code yet still measured much slower than the
  earlier clean baseline, it is best treated as environment/noise drift rather than as a real code
  regression signal
- keep using the earlier clean serial pair baseline for optimization retention decisions unless a
  new clean rerun reproduces the slower numbers

## 2026-07-06 retained: conservative shared-string worksheet scan for regular `.xlsx` sheets

- change:
  - add a guarded `appendSharedStringWorksheetText(...)` fast path ahead of the generic
    `encoding/xml.Decoder` worksheet reader in [extract.go](D:/workprj/officeread/extract.go)
  - the new path only activates for regular shared-string worksheets that avoid the risky features
    that previously made worksheet shortcuts fragile:
    - no formula cells
    - no `inlineStr`, boolean, string-error cell types
    - no worksheet `<t>` / phonetic run text
    - no hyperlinks / data-validations
    - no hidden row / column markers
    - no non-empty header/footer text sections
  - when active, it scans `<row>` / `<c>` / `<v>` byte slices directly, still preserving:
    - normal extracted text
    - large-value / shared-string duplicate suppression
    - markdown row collection via the retained prepared-row machinery
- rationale:
  - fresh shape evidence on `00012389.xlsx` showed the main hotspot worksheet is highly regular:
    - `816985` cells
    - `537209` shared-string cells
    - `0` formulas
    - `0` inline strings
    - `0` booleans
    - `0` hyperlinks
    - `0` data-validations
  - pprof kept pointing at `dec.RawToken()` as the dominant cost in `appendWorksheetText(...)`, so
    a conservative structural bypass was finally justified

Validation:
- focused correctness:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue|TestExtractXLSX|TestExtractWorksheet' ./`
- focused hotspots:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$' -benchmem -benchtime=1x ./`
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- serial focused pair rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-sharedscan-xlsx-pair-serial.json`
- serial `.xlsx` keyset rerun:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-sharedscan-xlsx-keyset-serial.json`
- same-session reverted baseline rerun for the exact same `.xlsx` keyset:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-current-baseline-xlsx-keyset-serial-rerun.json`
- full `.xlsx` compatibility rerun:
  - `testdata/web-samples/reports/compat-xlsx-20260706-after-sharedscan-candidate.json`
- repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused worksheet/extract hotspots improved strongly:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1367983300 ns/op -> 961121100 ns/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`: `832965300 ns/op -> 536413900 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2334292400 ns/op -> 1803280200 ns/op`
- focused pair rerun preserved outputs and improved the main sheet materially:
  - pair output parity:
    - `textBytes=192482085`
    - `images=0`
  - `00012389.xlsx`: `2791 ms -> 1973 ms`
  - `testRecordSizeExceeded.xlsx`: `3956 ms -> 4188 ms` on that rerun shape
  - pair total: `6747 ms -> 6161 ms`
- the broader `.xlsx` decision came from same-session, same-input, same-machine serial reruns:
  - reverted baseline keyset total: `27461 ms`
  - shared-scan keyset total: `23930 ms`
  - net improvement: `3531 ms`
  - keyset parity stayed exact on reported fields:
    - `.xlsx total=21`
    - `ok=21`
    - `textBytes=279385415`
    - `images=1`
- full `.xlsx` compatibility stayed perfect:
  - `1000 / 1000`
  - `errors=0`
  - `panics=0`
  - `empty=0`
- repository regression stayed green:
  - `ok officeread (cached)`

Interpretation:
- unlike the recent helper-level `.xlsx` attempts, this structural worksheet-path bypass survives
  both the hotspot view and the broader same-session serial `.xlsx` keyset rerun
- the explicit same-session reverted-baseline rerun matters here: it removes the earlier machine
  drift ambiguity and shows the keyset gain is real on the current host state
- because the guard is intentionally narrow and the full 1000-file `.xlsx` sweep remained clean,
  the optimization is worth keeping

Decision:
- Retained.

## 2026-07-06 follow-up: broader mixed-keyset and full compatibility after retained shared-string scan

Additional validation:
- broader 30-file mixed keyset rerun on the retained code:
  - `testdata/web-samples/reports/perf-exp-ai-assistant-sharedscan-mixed-keyset-serial.json`
- full cross-format compatibility rerun on the retained code:
  - `testdata/web-samples/reports/compat-all-20260706-after-sharedscan.json`

Observed results:
- mixed 30-file rerun stayed fully correct:
  - `.docx`: `ok=2`, `errors=0`, `panics=0`, `images=43`
  - `.xls`: `ok=7`, `errors=0`, `panics=0`
  - `.xlsx`: `ok=21`, `errors=0`, `panics=0`, `textBytes=279385415`, `images=1`
- mixed 30-file timings on this host state:
  - `.docx`: `6937 ms`
  - `.xls`: `15326 ms`
  - `.xlsx`: `21002 ms`
  - total mixed keyset: `43265 ms`
- the mixed `.xlsx` subset stayed close to the earlier retained isolated baseline while preserving
  the hotspot wins:
  - earlier isolated `.xlsx` keyset baseline: `20801 ms`
  - current retained mixed-run `.xlsx` subset: `21002 ms`
  - `00012389.xlsx` inside the mixed rerun: `1910 ms`
  - `testRecordSizeExceeded.xlsx` inside the mixed rerun: `3075 ms`
- full cross-format compatibility rerun remained green after retaining the optimization:
  - `.doc`: `1000 / 1000`, `errors=0`, `panics=0`
  - `.docx`: `1000 / 1000`, `errors=0`, `panics=0`
  - `.ppt`: `1000 / 1000`, `errors=0`, `panics=0`
  - `.pptx`: `1008 / 1008`, `errors=0`, `panics=0`
  - `.xls`: `1000 / 1000`, `errors=0`, `panics=0`
  - `.xlsx`: `1000 / 1000`, `errors=0`, `panics=0`
  - current full sweep total: `6008 / 6008`

Interpretation:
- the retained optimization is now backed by three levels of evidence:
  - focused hotspot benches
  - same-session reverted-baseline `.xlsx` keyset comparison
  - full cross-format compatibility rerun
- the broader mixed keyset remains noisy on non-`.xlsx` formats, so it is not a clean retention
  signal by itself
- however, because the code change is strictly in the `.xlsx` worksheet path, the decisive
  retention evidence remains the same-session `.xlsx`-only baseline-vs-candidate rerun plus the
  clean 6008-file compatibility sweep

## 2026-07-06 rejected: byte-oriented `cellRefIndexes(...)` parser

- experiment:
  - replace `cellRefIndexes(...)` internals with a byte-oriented ASCII parser
  - remove:
    - `strings.TrimSpace(...)`
    - `strings.ReplaceAll(..., "$", "")`
    - two rune-based loops
  - instead, scan the original string by byte, skip `$` / ASCII whitespace inline, parse the column
    letters in one loop, then parse row digits in one loop
  - keep the function signature and all call sites unchanged
- rationale:
  - `cellRefIndexes(...)` sits in several worksheet hot paths, including:
    - `appendWorksheetText(...)`
    - `appendSimpleInlineWorksheetTextPrepared(...)`
    - worksheet markdown helpers
  - fresh `pprof -list appendWorksheetText` still showed measurable time at the main worksheet call
    site around line `11098`
  - this made the parser itself a reasonable low-level target that could help both markdown and
    no-markdown worksheet extraction without adding caches

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- serial pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-cellref-parser-xlsx-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-cellref-parser-xlsx-pair-serial.csv <xlsx-pair>`

Observed results:
- focused benchmark signal looked broadly positive:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    about `2.43 s/op -> 2.36 s/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    about `1.436 s/op -> 1.396 s/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`:
    about `0.864 s/op -> 0.815 s/op`
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`:
    about `2.181 s/op -> 2.025 s/op`
- but the decisive serial pair rerun still regressed overall:
  - pair total: `4933 ms -> 5110 ms`
  - `00012389.xlsx`: `2037 ms -> 2323 ms`
  - `testRecordSizeExceeded.xlsx`: `2896 ms -> 2787 ms`
  - parity stayed stable:
    - `textBytes` unchanged
    - `images=0`
    - `ok=2/2`, `errors=0`, `panics=0`

Interpretation:
- this is one of the strongest recent focused candidates: it helped markdown, no-markdown, and the
  simple-inline hotspot view at the same time
- even so, the retained serial pair still got worse because the `00012389` whole-file path regressed
  more than `testRecord` improved
- under the current bar, that means the parser rewrite is still not keepable

Decision:
- Reverted.

## 2026-07-06 rejected: skip markdown cell cleanup for empty worksheet cells

- experiment:
  - in `appendWorksheetText(...)`, when closing a markdown-collected worksheet cell, first check
    `markdownCellText.Len() > 0`
  - if the builder is empty, skip
    `cleanMarkdownTableCellValue(markdownCellText.String())` entirely instead of sending an empty
    string through the markdown cleanup pipeline
  - leave all non-empty cells and all other code paths unchanged
- rationale:
  - this is a very narrow attempt to remove pure empty-cell overhead in the worksheet markdown path
  - if visible worksheets contain many empty cells within the collected markdown column range, the
    existing code still pays a cleanup call per such cell even though the result is always empty

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- serial pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-skip-empty-mdcell-xlsx-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-skip-empty-mdcell-xlsx-pair-serial.csv <xlsx-pair>`

Observed results:
- focused benchmark signal looked promising:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    about `2.43 s/op -> 2.38 s/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    about `1.436 s/op -> 1.359 s/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`:
    about `0.864 s/op -> 0.837 s/op`
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`:
    about `2.181 s/op -> 2.152 s/op`
- but the decisive serial pair rerun still regressed against the retained baseline:
  - pair total: `4933 ms -> 5190 ms`
  - `00012389.xlsx`: `2037 ms -> 2307 ms`
  - `testRecordSizeExceeded.xlsx`: `2896 ms -> 2883 ms`
  - parity stayed stable:
    - `textBytes` unchanged
    - `images=0`
    - `ok=2/2`, `errors=0`, `panics=0`

Interpretation:
- this is another strong example of a focused hotspot-local win that does not survive the real
  serial retained workload
- skipping empty markdown-cell cleanup helped the direct benchmark view, but not enough to improve
  the actual pair end-to-end

Decision:
- Reverted.

## 2026-07-06 rejected: reuse worksheet markdown row slice backing across rows

- experiment:
  - in `appendWorksheetText(...)`, stop allocating a fresh `make([]string, 0, 16)` for every
    markdown-collected worksheet row
  - instead, when a new row starts and markdown collection is active, reuse the existing
    `markdownRowValues` backing array with `markdownRowValues = markdownRowValues[:0]`
  - rely on the existing row-close copy in `compactPreparedWorksheetMarkdownRow(...)` to keep
    previously appended rows independent
- rationale:
  - the current row-slice worksheet path is already retained, but it still allocates a small new
    slice header/backing for every collected row
  - because row close already copies out the compact row, slice reuse looked like a low-risk way to
    shave some allocation churn without changing parsing logic

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`

Observed results:
- the focused signal again missed the bar:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    about `2.43 s/op -> 2.42 s/op` (noise-level)
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    about `1.463 s/op -> 1.555 s/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`:
    about `0.829 s/op -> 0.973 s/op`
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`:
    about `2.181 s/op -> 2.512 s/op`
- allocation counts did not improve in a meaningful way on the target hotspot:
  - worksheet path stayed around `9.82M allocs/op`

Interpretation:
- even this small row-slice reuse change was not enough to produce a usable focused win
- under the current retention bar, there is no reason to promote it into serial reruns

Decision:
- Reverted.

## 2026-07-06 rejected: inline `len(hiddenCols)` guard before worksheet hidden-column checks

- experiment:
  - keep the existing hidden-column logic unchanged
  - in the three worksheet cell-entry paths that currently do:
    - `hiddenColumnCell(cellRef, hiddenCols)`
    - `columnHidden(cellCol, hiddenCols)`
  - first check `len(hiddenCols) > 0`, and only call the two helpers when there are actually hidden
    column ranges to inspect
- rationale:
  - many worksheets appear to have no hidden columns at all
  - the current code still pays two function calls per cell in that common case, even though both
    helpers immediately return after seeing an empty slice
  - unlike earlier hidden-range gating experiments, this change adds no worksheet-wide pre-scan and
    only trims obviously empty per-cell work

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`

Observed results:
- the focused view stayed mixed and missed the retention bar:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    about `2.43 s/op -> 2.28 s/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    about `1.463 s/op -> 1.526 s/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`:
    about `0.829 s/op -> 0.915 s/op`
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`:
    about `2.181 s/op -> 2.449 s/op`
- because the worksheet and no-markdown hotspot views both regressed, this was not promoted to
  serial reruns

Interpretation:
- avoiding the empty hidden-column helper calls was not enough to improve the actual target paths
- even with zero worksheet-wide pre-scan overhead, the end result still moved the wrong way on the
  direct `00012389` hotspot

Decision:
- Reverted.

## 2026-07-06 rejected: marker-gated special worksheet element checks in `appendWorksheetText(...)`

- experiment:
  - pre-scan worksheet bytes for a few rarely-used marker families:
    - system-cell text: `<extLst`, `<ext `
    - attribute text: `<hyperlink`, `<dataValidation`
    - header/footer: `<headerFooter`, `<oddHeader`, `<oddFooter`, `<evenHeader`, `<evenFooter`, `<firstHeader`, `<firstFooter`
    - phonetic: `<rPh`, `<phoneticPr`
  - use those booleans to skip whole branches inside the `RawToken` loop when the corresponding
    element family is absent from the worksheet
- rationale:
  - temporary sheet-marker probing showed that on the active hotspot worksheets:
    - `00012389.xlsx` `sheet1.xml`: `extLst=0`, `ext=0`, `hyperlinks=0`, `dataValidation=0`,
      `phonetic=1`, `headerFooter=1`
    - `testRecordSizeExceeded.xlsx`: the sampled sheet also had no `ext` / hyperlink /
      data-validation markers
  - this made the “absent element families” look like a plausible place to reduce per-token branch
    work around `RawToken`

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- temporary marker probe:
  - `& 'C:\Program Files\Go\bin\go.exe' run tmp-shape/probe_xlsx_token_markers.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`

Observed results:
- the focused view was mixed, but not keepable:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    about `2.43 s/op -> 2.43 s/op` (no useful change)
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    about `1.463 s/op -> 1.595 s/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`:
    about `0.829 s/op -> 0.965 s/op`
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`:
    about `2.181 s/op -> 2.164 s/op`
- because the target `00012389` worksheet and no-markdown paths both regressed, this was not
  promoted to serial reruns

Interpretation:
- gating absent element families did help one non-target hotspot slightly
- but it hurt the actual retained `00012389` worksheet path, so the branch savings were not worth
  the added pre-scan and control flow

Decision:
- Reverted.

## 2026-07-06 rejected: single-pass `cleanTextFastPath(...)` scan with deferred control-fragment check

- experiment:
  - refactor `cleanTextFastPath(...)` to collapse several checks into one byte loop:
    - fold `strings.ContainsAny(s, "/\\%[]()")`
    - fold `containsRIDFold(s)`
    - keep the existing ASCII/control/underscore/double-space checks in the same pass
  - move `cleanTextFastPathControlFragment(s)` from the top of the function to the end, so only
    strings that survive the cheap byte scan pay the control-fragment helper cost
- rationale:
  - temporary probing on `00012389.xlsx` `sheet1.xml` showed the fast path is overwhelmingly the
    common case:
    - `total=567787`
    - `fastOK=562778`
    - `fastReject=5009`
    - `controlCandidate=1`
  - that made the current multi-pass structure look like an attractive target: many strings were
    likely paying for several scans even though almost all of them were plain fast-path successes

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`

Observed results:
- the focused signal was uniformly negative:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    about `2.43 s/op -> 2.53 s/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    about `1.463 s/op -> 1.863 s/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`:
    about `0.829 s/op -> 1.198 s/op`
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`:
    about `2.181 s/op -> 2.723 s/op`
- because even the narrow focused layer regressed clearly, this was not promoted to serial reruns

Interpretation:
- despite the strong shape evidence in favor of the fast path, collapsing the scans and deferring
  the control-fragment helper changed branch behavior in a way that hurt the real workload
- under the current bar, this is another structural candidate that fails before integrated
  validation

Decision:
- Reverted.

## 2026-07-06 rejected: digit-leading early return in `maybeControlFragmentText(...)`

- experiment:
  - add a very narrow early return at the top of `maybeControlFragmentText(...)`
  - for digit-leading strings:
    - immediately return `false` for `0...`
    - for `1...`, keep only the special `1table` case and a guarded GUID-shaped fallback
  - leave every non-digit-leading path unchanged
- rationale:
  - temporary probing on `00012389.xlsx` `sheet1.xml` showed:
    - `total=567787`
    - `fastOK=562778`
    - `controlCandidate=1`
    - `first[digit]=246158`
  - that meant a very large share of strings were digit-leading while almost none became true
    control-fragment candidates, making digit-first early return look like a low-risk way to avoid
    repeated helper work

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- temporary shape probe:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpProbeXLSX00012389FastPathShape$' -v ./`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`

Observed results:
- the direct benchmark signal was clearly negative:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    about `2.43 s/op -> 2.40 s/op` (noise-level change)
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    about `1.463 s/op -> 1.741 s/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`:
    about `0.829 s/op -> 1.094 s/op`
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`:
    about `2.181 s/op -> 2.606 s/op`
- because even the focused layer regressed materially, this change was not promoted to serial reruns

Interpretation:
- despite the strong shape evidence, the extra branching and special-case handling for digit-leading
  inputs did not help the real hotspot path
- under the current standard, this is another small structural candidate that fails before it even
  reaches integrated rerun validation

Decision:
- Reverted.

## 2026-07-06 rejected: plain shared-string markdown fast path in `appendWorksheetText(...)`

- experiment:
  - keep the worksheet XML decoder path unchanged
  - on markdown-collected shared-string cells, try a narrower fast path than the earlier builder
    bypass experiment:
    - only trigger when the raw shared string has no newline and no obvious special markers such as
      `\\`, `|`, `/`, `%`, `=`, `:`, brackets, `<`, `>`, `&`, or `rid`
    - run a combined preparation helper that performs `cleanText(...)`, discard checks, truncation,
      and `normalizeMarkdownTextLine(...)`
    - if that helper succeeds, skip the normal
      `cleanMarkdownTableCellValue(...) + prepareMarkdownTableCellValue(...)` path for the cell
  - leave non-shared-string cells and the simple-inline pipeline untouched
- rationale:
  - follow-up shape probing on `00012389.xlsx` `xl/worksheets/sheet4.xml` refined the earlier
    shared-string picture:
    - `sharedCells=30804`
    - `singleLine=30804`
    - marker counts were very low relative to the shared-string pool:
      - `slash=380`
      - `pipe=1`
      - `angle=0`
      - `amp=287`
      - `colon=161`
      - `rid=26`
    - this suggested that most shared-string markdown values might be eligible for a very plain,
      single-line fast path

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- marker probe:
  - `& 'C:\Program Files\Go\bin\go.exe' run tmp-shape/probe_xlsx_00012389_markers.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- serial pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-shared-plainfast-xlsx-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-shared-plainfast-xlsx-pair-serial.csv <xlsx-pair>`

Observed results:
- focused benchmark signal again looked attractive on the target hotspot:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    about `2.430 s/op -> 2.218 s/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    about `1.463 s/op -> 1.411 s/op`
  - allocations dropped materially:
    - extract path: `~13487961 allocs/op -> 12166299 allocs/op`
    - worksheet path: `~9823043 allocs/op -> 8665906 allocs/op`
- but the decisive serial pair rerun regressed badly:
  - pair total: `4933 ms -> 5672 ms`
  - `00012389.xlsx`: `2037 ms -> 2583 ms`
  - `testRecordSizeExceeded.xlsx`: `2896 ms -> 3089 ms`
  - parity stayed stable:
    - `textBytes` unchanged
    - `images=0`
    - `ok=2/2`, `errors=0`, `panics=0`

Interpretation:
- the plain shared-string fast path was well-aligned with the observed worksheet shape and again
  produced a real hotspot-local win
- but it still failed the retained end-to-end standard once measured serially on the actual pair
- under the current bar, this remains another “good benchmark, bad rerun” candidate and cannot be
  kept

Decision:
- Reverted.

## 2026-07-06 rejected: shared-string markdown direct pass in `appendWorksheetText(...)`

- experiment:
  - keep the existing worksheet XML decoder path
  - only on markdown-collected shared-string cells (`t="s"`), avoid appending the shared string into
    `markdownCellText`
  - instead hold the shared string directly and feed it to
    `cleanMarkdownTableCellValue(...)` at cell close when no other markdown text segments were
    collected
  - leave non-shared-string cells, plain text extraction, and all simple-inline logic unchanged
- rationale:
  - temporary shape probing on the actual benchmark worksheet for `00012389.xlsx`
    (`xl/worksheets/sheet4.xml`) showed:
    - `totalCells=55411`
    - `markdownCells=55411`
    - `sharedCells=30804` (`55.59%`)
    - `sharedSingleSegment=30804` (`100%` of shared)
    - `sharedSingleSegmentPlain=29077` (`94.39%` of shared)
  - that made the markdown builder write for shared-string cells look like a plausible structural
    cost in the remaining worksheet hotspot

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- shape probe:
  - `& 'C:\Program Files\Go\bin\go.exe' run tmp-shape/probe_xlsx_00012389_shape.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- serial pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-shared-direct-xlsx-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-shared-direct-xlsx-pair-serial.csv <xlsx-pair>`

Observed results:
- focused benchmark signal looked encouraging on the targeted worksheet hotspot:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    about `1.463 s/op -> 1.413 s/op`
  - allocations also dropped materially:
    - `572 MB -> 563 MB`
    - `9823043 allocs/op -> 9285369 allocs/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx` also improved modestly:
    - about `2.431 s/op -> 2.430 s/op` with lower allocs
- but the decisive serial pair rerun regressed against the retained baseline:
  - pair total: `4933 ms -> 5389 ms`
  - `00012389.xlsx`: `2037 ms -> 2201 ms`
  - `testRecordSizeExceeded.xlsx`: `2896 ms -> 3188 ms`
  - parity stayed stable:
    - `textBytes` unchanged
    - `images=0`
    - `ok=2/2`, `errors=0`, `panics=0`

Interpretation:
- this is another case where a very plausible worksheet-local optimization improved the direct
  hotspot view and even cut allocations substantially
- but it still failed the retained end-to-end bar once rerun serially on the real pair
- because the integrated pair regressed clearly, it is not keepable

Decision:
- Reverted.

## 2026-07-06 rejected: tighter hidden-range markers plus hidden-check gate in `appendWorksheetText(...)`

- experiment:
  - tighten `worksheetMayHaveHiddenRanges(...)` so it only looks for more specific hidden markers:
    - ` hidden="1"`, ` hidden="true"` and quote variants
    - ` width="0`, ` ht="0` and quote variants, with a leading space
  - use that gate inside `appendWorksheetText(...)` to skip:
    - hidden column range collection
    - row hidden checks
    - per-cell hidden-column checks
    - `worksheetElementHiddenByRef(...)`
    when the worksheet bytes appear to have no hidden ranges
- rationale:
  - direct inspection of the current hotspot worksheets showed the old broad markers were producing
    false positives:
    - `00012389.xlsx` visible worksheet: no real `hidden`, no `width="0"`, but the broad `ht="0`
      probe falsely matched inside page-margin/footer attributes such as `right="0.75"` /
      `footer="0.5"`
    - `testRecordSizeExceeded.xlsx` showed the same kind of false positive
  - that suggested `appendWorksheetText(...)` might be paying hidden-range overhead on worksheets
    that actually have no hidden rows or columns at all

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- marker sanity check:
  - a temporary Go probe confirmed the tightened gate returned `false` for the active hotspot
    worksheets:
    - `00012389.xlsx` `xl/worksheets/sheet4.xml`
    - `testRecordSizeExceeded.xlsx` `xl/worksheets/sheet1.xml`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`

Observed results:
- despite the gate correctly classifying the hotspot worksheets as having no hidden ranges, the
  focused performance signal was still negative or too weak:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    about `1.299 s/op -> 1.709 s/op`
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`:
    about `2.038 s/op -> 2.498 s/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    about `2.413 s/op -> 2.364 s/op`
- because the direct worksheet hotspot regressed materially, this was not promoted to serial
  integrated reruns

Interpretation:
- the hidden-range false-positive theory was real, but the control flow added around it did not
  improve the retained hotspot path
- under the current standard, a slight whole-extract movement is not enough when the direct
  worksheet hotspot gets worse

Decision:
- Reverted.

## 2026-07-06 rejected: byte-parsed cell refs in simple-inline worksheet extraction

- experiment:
  - narrow the change to `appendSimpleInlineWorksheetTextPrepared(...)`
  - replace `cellRefIndexes(string(cellRef))` with a byte-oriented parser to avoid allocating a
    temporary string for every inline-string cell ref
  - keep all downstream row/column behavior unchanged
- rationale:
  - the simple-inline path is still one of the dominant `.xlsx` hotspots on
    `testRecordSizeExceeded.xlsx`
  - that path parses millions of cell refs, and the existing implementation always converts the
    `r=` attribute bytes into a Go string before parsing
  - this looked like a clean way to reduce per-cell overhead without adding caches or changing
    markdown logic

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkXLSXSimpleInlineHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkExtractXLSXHotspots/00012389.xlsx$' -benchmem -benchtime=1x ./`

Observed results:
- the direct hotspot signal was negative enough to stop there:
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`:
    about `2.181 s/op -> 2.738 s/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    about `1.463 s/op -> 1.870 s/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx` stayed around `2.41 s/op`
- allocation counts did not improve enough to justify a broader rerun:
  - `testRecordSizeExceeded.xlsx`: `7600115 allocs/op -> 7600078 allocs/op`
  - `00012389.xlsx` extract path: `13487961 allocs/op -> 13487878 allocs/op`

Interpretation:
- the string allocation being removed here is real, but it is evidently too small relative to the
  added byte-parsing work
- under the current retention bar, this is not close enough to promote into serial reruns

Decision:
- Reverted.

## 2026-07-06 rejected: gated helper checks inside `maybeControlFragmentText(...)`

- experiment:
  - keep the existing control-fragment decisions unchanged
  - add cheap marker gates before the more expensive helper recognizers inside
    `maybeControlFragmentText(...)`
  - examples:
    - only try legacy-object parsing when `!` is present
    - only try OLE identifier recognition for likely starters such as `{`, quotes, or
      `appid/classid/clsid/guid` initials
    - only try wrapper / OOXML fragment helpers for likely leading bytes
  - also reuse `asciiLower(s[0])` for the later first-byte switch
- rationale:
  - fresh `pprof` on the current retained `.xlsx` hotspots showed control-fragment recognition still
    taking meaningful time inside `cleanTextFastPath(...)`:
    - `00012389.xlsx` worksheet hotspot:
      - `maybeControlFragmentText(...)`: `11.98%` cumulative
      - `looksLikeOLEIdentifierFragment(...)`: `5.39%` cumulative
      - `looksLikeOOXMLMarkupNameFragment(...)`: `2.40%` cumulative
    - `testRecordSizeExceeded.xlsx` simple-inline hotspot:
      - `maybeControlFragmentText(...)`: `12.65%` cumulative
      - `looksLikeOLEIdentifierFragment(...)`: `5.25%` cumulative
      - `looksLikeOOXMLMarkupNameFragment(...)`: `4.01%` cumulative
  - this made cheap plausibility gates look like a reasonable way to reduce false-positive helper
    work without introducing caches

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$' -benchmem -benchtime=1x ./`
- serial pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-controlgate-xlsx-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-controlgate-xlsx-pair-serial.csv <xlsx-pair>`
- serial `.xlsx` keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-controlgate-xlsx-keyset-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-controlgate-xlsx-keyset-serial.csv <xlsx-keyset-21>`

Observed results:
- focused benchmark signal was mixed but initially encouraging on the direct worksheet hotspots:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: about `1.463 s/op -> 1.169 s/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`: about `0.829 s/op -> 0.787 s/op`
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`: about
    `2.181 s/op -> 1.874 s/op`
  - but `BenchmarkExtractXLSXHotspots/00012389.xlsx` moved the wrong way in the same validation run:
    about `2.409 s/op -> 3.600 s/op`
- serial pair rerun improved slightly against the retained baseline:
  - pair total: `4933 ms -> 4829 ms`
  - `00012389.xlsx`: `2037 ms -> 2134 ms`
  - `testRecordSizeExceeded.xlsx`: `2896 ms -> 2695 ms`
- decisive serial `.xlsx` keyset rerun regressed against the retained baseline:
  - retained baseline total: `20801 ms`
  - experiment total: `20954 ms`
  - compatibility/parity stayed stable:
    - `ok=21/21`
    - `errors=0`
    - `panics=0`
    - `textBytes=279385415`
    - `images=1`

Interpretation:
- the gating idea really did help some narrow `.xlsx` paths, especially the two most visible
  worksheet hotspots
- but under the retained standard, the integrated serial `.xlsx` subset result still matters more
  than those local wins
- because the 21-file `.xlsx` keyset regressed overall, this change does not qualify as a keepable
  optimization

Decision:
- Reverted.

## 2026-07-06 rejected: direct-string fast path for single-segment `.xlsx` worksheet markdown cells

- experiment:
  - in `appendWorksheetText(...)`, try to avoid `strings.Builder` churn for the common case where a
    markdown-collected worksheet cell receives only one text segment
  - keep the first segment in a direct string
  - only materialize the builder if a second segment arrives
  - keep cleanup and markdown preparation behavior unchanged at cell close
- rationale:
  - the remaining `.xlsx` hotspot is still concentrated in worksheet markdown work on
    `00012389.xlsx`
  - many cells appear to be plain shared-string or inline values with a single payload segment, so
    skipping builder writes for the first segment looked like a low-risk structural reduction

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$' -benchmem -benchtime=1x ./`

Observed results:
- allocation counts improved in the markdown path:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `9823043 allocs/op -> 9254988 allocs/op`
- but wall time regressed sharply:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: around `2.285 s/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1.396 s/op -> 2.014 s/op`
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`: also drifted upward in the same run to about
    `1.137 s/op`

Interpretation:
- this change reduced some allocation churn, but the additional branching and bookkeeping did not
  translate into usable latency
- the signal was negative enough at the focused benchmark layer that it was not worth promoting to
  serial integrated reruns

Decision:
- Reverted.

## 2026-07-06 rejected: `.xlsx` markdown shared-string-index prepared-cell cache

- experiment:
  - inside `appendWorksheetText(...)`, add a narrow cache only for markdown preparation of `.xlsx`
    shared-string cells
  - cache key: shared-string index
  - cache value: prepared markdown cell text after
    `cleanMarkdownTableCellValue(...) + prepareMarkdownTableCellValue(...)`
  - leave non-shared-string cells and the plain text extraction path unchanged
- rationale:
  - temporary shared-string analysis on `00012389.xlsx` showed substantial shared-index reuse in
    markdown-collected cells:
    - `totalMarkdownCells=567787`
    - `sharedMarkdownCells=537209`
    - `uniqueSharedIndexes=102206`
    - `repeatedSharedIndexes=26646`
    - `repeatedSharedIndexUses=461649`
  - top repeated values such as `Not OA`, `Journal`, `Active`, `Social Sciences`, and
    `Health Sciences` appeared often enough that shared-index-scoped reuse looked promising

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx$|BenchmarkXLSXWorksheetTextNoMarkdown00012389$' -benchmem -benchtime=1x ./`
- serial pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-sharedidx-cache-xlsx-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-sharedidx-cache-xlsx-pair-serial.csv <xlsx-pair>`
- serial `.xlsx` keyset rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-sharedidx-cache-xlsx-keyset-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-sharedidx-cache-xlsx-keyset-serial.csv <xlsx-keyset>`
- full `.xlsx` compatibility + performance rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 1 -json testdata/web-samples/reports/compat-xlsx-20260706-after-sharedidx-cache.json -csv testdata/web-samples/reports/compat-xlsx-20260706-after-sharedidx-cache.csv <xlsx-1000>`

Observed results:
- focused benchmark allocs improved:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: about `12.746M allocs/op` vs prior retained
    baseline around `13.488M allocs/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: about `9.202M allocs/op` vs prior
    retained baseline around `9.822M allocs/op`
- serial pair rerun regressed:
  - pair total: `4933 ms -> 5746 ms`
  - `00012389.xlsx`: `2037 ms -> 2085 ms`
  - `testRecordSizeExceeded.xlsx`: `2896 ms -> 3661 ms`
- serial `.xlsx` keyset rerun improved and kept parity:
  - `.xlsx` subset total: `20801 ms -> 19972 ms`
  - output parity: `NO_DIFF`
- but the decisive integrated `.xlsx` 1000-file rerun regressed:
  - retained baseline total: `55540 ms`
  - experiment total: `57218 ms`
  - retained baseline max: `3158 ms`
  - experiment max: `3439 ms`
  - compatibility stayed perfect: `1000/1000`, `errors=0`, `panics=0`

Interpretation:
- the shared-index reuse hypothesis was real enough to reduce allocations in the narrow hotspot
- but the added cache bookkeeping did not hold up across the full retained `.xlsx` workload
- this is another case where a plausible local markdown optimization improved one view while making
  the integrated serial rerun worse

Decision:
- Reverted.

## 2026-07-06 current hotspot split: `00012389.xlsx` markdown cells are mostly plain single-line text

Experiment:
- temporarily instrument the representative worksheet markdown-cell corpus from `00012389.xlsx`
  to count how often hidden-reference and control-fragment branches actually trigger
- temporarily split the `cleanMarkdownTableCellValue(...)` chain into stage-level benchmarks

Observed results:
- representative worksheet markdown-cell corpus:
  - `values=4102`
  - `singleLine=4071`
  - `multiLine=31`
  - `markerHits=0`
  - `discardHits=23`
  - `controlHits=0`
  - `lineDiscardHits=54`
  - `truncatedHits=0`
- temporary stage split with `-benchtime=10000x`:
  - `cleanText`: `201.2 ns/op`
  - `stripInlineHiddenOfficeReferences`: `65.84 ns/op`
  - `maybeDiscardableHiddenOfficeText`: `137.5 ns/op`
  - `maybeControlFragmentText`: `84.87 ns/op`

Interpretation:
- on the current representative `00012389.xlsx` workload, the markdown-cell common path is
  overwhelmingly plain single-line text
- inline hidden-reference stripping and control-fragment matches are rare on this corpus
- future retained work should continue to optimize the common path first, rather than assuming the
  marker-heavy defense paths dominate

Decision:
- Record as guidance only; temporary analysis code removed after measurement.

## 2026-07-06 rejected: trim-free prepared-row compaction for `.xlsx` worksheet markdown

Experiment:
- add a specialized worksheet-markdown row compactor that skips the repeated `TrimSpace` scan used
  by the generic `compactWorksheetMarkdownRow(...)`
- only populate that path from already-prepared worksheet markdown cell values

Validation:
- `BenchmarkXLSXMarkdownRowCompact00012389 -count=5`
- `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx -count=5`

Observed results:
- isolated row-compaction microbenchmark stayed noisy:
  - `1125 / 1283 / 1745 / 1690 / 1691 ns/op`
- real worksheet hotspot did not improve convincingly:
  - `1245719200 / 1215814200 / 1794226800 / 1747040300 / 1625623500 ns/op`

Interpretation:
- the helper-level idea may remove a little redundant work locally, but it did not survive the
  real retained worksheet hotspot well enough to keep

Decision:
- Reverted.

## 2026-07-06 rejected: `cleanTextFastPath(...)` control-fragment check reorder

Experiment:
- keep `cleanTextFastPath(...)` semantics unchanged, but defer the
  `cleanTextFastPathControlFragment(...)` call until after the cheap ASCII / marker screening loop

Validation:
- `TestCleanTextFastPathReorderedMatches00012389`
- `BenchmarkXLSXMarkdownCleanFastPathVariants00012389 -count=5`

Observed results:
- representative-corpus equivalence check passed
- reordered fast path was slower:
  - current: `117.8 / 133.6 / 152.4 / 133.8 / 169.7 ns/op`
  - reordered: `165.5 / 164.0 / 161.7 / 160.6 / 151.7 ns/op`

Interpretation:
- the existing check order is already the better one for the current representative markdown-cell
  workload

Decision:
- Not applied.

## 2026-07-06 rejected: `.xlsx` single-line markdown common-path hidden-check gating

Experiment:
- restructure the single-line branch of `cleanMarkdownTableCellValue(...)` so the common path:
  - checks `maybeControlFragmentText(...)` first
  - only runs `maybeDiscardableHiddenOfficeText(...)` when the cleaned cell still contains
    hidden-office marker characters such as `=`, `:`, `/`, `\`, `%`, `<`, `>`, `#`, or `rid`
- keep representative-corpus behavior equivalent before applying

Validation:
- temporary representative-corpus equivalence check
- isolated cell-clean benchmark on `00012389.xlsx` worksheet markdown values
- `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx -count=5`
- `cmd/compatcheck -repeat 5 ... 00012389.xlsx`

Observed results:
- isolated cell-clean benchmark improved materially:
  - current: `472.5 / 508.4 / 484.4 / 481.8 / 474.5 ns/op`
  - candidate: `385.1 / 388.8 / 408.6 / 394.9 / 377.7 ns/op`
- output parity stayed intact on the repeat-aware single-file rerun:
  - `text=11148655`
  - `images=0`
- but the integrated signals did not hold:
  - worksheet hotspot before apply:
    `1243286400 / 1204917700 / 1532021200 / 1565506400 / 1538926400 ns/op`
  - worksheet hotspot after apply:
    `1192275100 / 1345777500 / 1621469000 / 1593249600 / 1648980400 ns/op`
  - repeat-aware single-file rerun before apply:
    `ms=2150 min=2053 max=2565 runs=[2060 2150 2526 2565 2053]`
  - repeat-aware single-file rerun after apply:
    `ms=2484 min=2012 max=3076 runs=[2012 2484 3076 3006 2423]`

Interpretation:
- this was a legitimate common-path candidate with a real microbenchmark win
- but it still failed the retained end-to-end standard once the actual worksheet and compatcheck
  workloads were rerun

Decision:
- Reverted.

## 2026-07-06 rejected: `bytes.Contains`-driven `simpleInlineWorksheetCandidate(...)`

Experiment:
- try replacing the current single-pass candidate loop with a detector built from:
  - one positive `bytes.Contains(..., []byte(\`t="inlineStr"\`))`
  - multiple negative `bytes.Contains(...)` checks for forbidden worksheet markers such as
    `<!DOCTYPE`, `<f`, `<v`, `<cols`, `<rPh`, `<phoneticPr`, `t="s"`, `t="b"`, `t="str"`,
    `hidden`, `hyperlink`, `dataValidation`, `Header`, `Footer`, and `ht=`

Why:
- current profile still shows `simpleInlineWorksheetCandidate(...)` dominating the
  `testRecordSizeExceeded.xlsx` simple-inline path

Validation:
- temporary parity check on:
  - positive hotspot worksheet: `testRecordSizeExceeded.xlsx`
  - negative hotspot worksheet: `00012389.xlsx`
- `BenchmarkXLSXSimpleInlineWorksheetCandidateVariants -count=3`

Observed results:
- parity matched on the hotspot positive/negative worksheets
- but the performance result was much worse:
  - positive `testRecordSizeExceeded.xlsx`:
    - current: `451517533 / 456994833 / 569939900 ns/op`
    - candidate: `1861008000 / 1917337600 / 1729359400 ns/op`
  - negative `00012389.xlsx`:
    - current: `1325 / 1217 / 1216 ns/op`
    - candidate: `5762842 / 5753632 / 5738494 ns/op`

Interpretation:
- repeated whole-slice `bytes.Contains(...)` passes are decisively worse than the current
  single-pass loop on both the positive and early-reject negative hotspot worksheets

Decision:
- Not applied.

## 2026-07-06 rejected: sparse interesting-byte scan for `simpleInlineWorksheetCandidate(...)`

Experiment:
- keep the candidate rules unchanged, but replace the current byte-by-byte loop with a sparse scan
  that jumps between only the interesting lead bytes using `bytes.IndexAny(...)`
- interesting bytes:
  - `<`, `t`, `h`, `d`, `H`, `F`, and whitespace

Why:
- after the failed multi-`bytes.Contains(...)` experiment, this was the next most conservative way
  to reduce the candidate scanner's per-byte work without changing its rule set

Validation:
- temporary parity check on:
  - positive hotspot worksheet: `testRecordSizeExceeded.xlsx`
  - negative hotspot worksheet: `00012389.xlsx`
- `BenchmarkXLSXSimpleInlineWorksheetCandidateSparseVariants -count=3`

Observed results:
- parity matched on the hotspot positive/negative worksheets
- but the sparse scan was still slower:
  - positive `testRecordSizeExceeded.xlsx`:
    - current: `444585333 / 530123050 / 506335700 ns/op`
    - sparse: `1108017700 / 1101259800 / 1100491400 ns/op`
  - negative `00012389.xlsx`:
    - current: `1409 / 1210 / 1299 ns/op`
    - sparse: `2126 / 2133 / 2055 ns/op`

Interpretation:
- repeated `bytes.IndexAny(...)` restarts are still more expensive than the current straight-line
  single-pass loop

Decision:
- Not applied.

## 2026-07-06 rejected: `cleanTextFastPath(...)` reorder on simple-inline testRecord values

Experiment:
- collect representative simple-inline cell texts from `testRecordSizeExceeded.xlsx`
- re-test the earlier `cleanTextFastPath(...)` reorder idea on this workload only:
  - current order
  - reordered version with `cleanTextFastPathControlFragment(...)` deferred until after the cheap
    ASCII / marker screening loop

Why:
- after re-profiling the simple-inline text-only hotspot with `-benchtime=3x`, `cleanTextFastPath`
  still showed up as a real flat cost, so this was worth checking on the correct workload even
  though the general-worksheet `00012389.xlsx` result had already rejected it

Validation:
- temporary parity check on collected simple-inline texts
- `BenchmarkXLSXCleanTextFastPathVariantsTestRecord -count=5`

Observed results:
- parity matched on the collected simple-inline values
- performance stayed too mixed:
  - current: `208.6 / 224.1 / 228.1 / 239.0 / 250.5 ns/op`
  - reordered: `219.9 / 217.5 / 229.6 / 249.6 / 263.6 ns/op`

Interpretation:
- this workload does not produce a stable enough win to justify touching the actual code path

Decision:
- Not applied.

## 2026-07-06 component split: `testRecordSizeExceeded.xlsx` simple-inline values all hit `cleanTextFastPath(...)`

Experiment:
- collect representative simple-inline cell texts from `testRecordSizeExceeded.xlsx`
- measure:
  - how many values hit `cleanTextFastPath(...)`
  - how many exceed `maxRepeatedTextPartBytes`
  - component timings for `simpleInlineCellText`, `cellRefParse`, `appendWorksheetValue`,
    `cleanText`, and `appendCleanedTextBlock`

Observed results:
- collected simple-inline cells:
  - `cells=3000000`
  - `large=0`
  - `fastPath=3000000`
  - `emptyAfterClean=0`
- component timings:
  - `simpleInlineCellText`: about `72-106 ns/op`
  - `cellRefParse`: about `38-48 ns/op`
  - `appendWorksheetValue`: about `263-315 ns/op`
  - `cleanText`: about `222-342 ns/op`
  - `appendCleanedTextBlock`: about `36-52 ns/op`

Interpretation:
- on this workload, the simple-inline common path is extremely uniform:
  - every collected cell text hits `cleanTextFastPath(...)`
  - none use the large-value de-duplication map
- this sharply narrows where future retained work should look

Decision:
- Record as guidance only; temporary analysis code removed after measurement.

## 2026-07-06 rejected: direct fast-path append inside `appendWorksheetValue(...)`

Experiment:
- in `appendWorksheetValue(...)`, detect `cleanTextFastPath(...)` explicitly
- if it hits, append the already-trimmed value directly with `appendTrimmedTextBlock(...)`
- keep the old path for the non-fast-path fallback

Validation:
- temporary parity check on collected simple-inline cell texts
- same-tree component benchmark comparing current code path vs previous helper
- integrated checks after applying:
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx -count=5`
  - `cmd/compatcheck -repeat 5 ... testRecordSizeExceeded.xlsx`

Observed results:
- representative parity passed
- integrated checks after applying looked plausible:
  - text-only hotspot after apply:
    `1317174900 / 1335036500 / 1275913600 / 1269446900 / 1288140700 ns/op`
  - repeat-aware single-file rerun after apply:
    `ms=2739 min=2557 max=3010`, `text=181333430`, `images=0`
- but the decisive same-tree component comparison was too mixed:
  - current:
    `297.0 / 271.2 / 263.0 ns/op`
  - previous:
    `260.1 / 265.3 / 289.1 ns/op`

Interpretation:
- this candidate is close enough to be interesting, but not strong enough yet to retain

Decision:
- Reverted.

## 2026-07-06 rejected: merged single-loop `cleanTextFastPath(...)`

Experiment:
- rewrite `cleanTextFastPath(...)` into a merged single-loop variant that:
  - folds the `ContainsAny` rejection characters into the main byte scan
  - tracks a cheap `rid` hint during the loop
  - only calls `containsRIDFold(...)` if the hint is present
  - moves `cleanTextFastPathControlFragment(...)` to the end

Why:
- the current simple-inline component split shows `testRecordSizeExceeded.xlsx` is effectively a
  `100% cleanTextFastPath(...)` workload
- temporary representative microbenchmarks made this look promising enough to test

Validation:
- temporary parity check on representative values from:
  - `testRecordSizeExceeded.xlsx`
  - `00012389.xlsx`
- temporary microbenchmark:
  - `BenchmarkCleanTextFastPathMerged -count=5`
- integrated checks after applying:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx -count=5`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx -count=5`
  - `cmd/compatcheck -repeat 5 ... 00012389.xlsx`
  - `cmd/compatcheck -repeat 5 ... testRecordSizeExceeded.xlsx`

Observed results:
- representative parity passed
- representative microbench looked encouraging for `00012389`:
  - current: `99.01 / 111.9 / 120.3 / 104.8 / 106.7 ns/op`
  - merged: `84.23 / 89.63 / 93.80 / 92.25 / 93.95 ns/op`
- but integrated signals regressed:
  - `00012389.xlsx` worksheet hotspot:
    `1443830000 / 1555045300 / 1979124700 / 2045242800 / 2065070300 ns/op`
  - `testRecordSizeExceeded.xlsx` simple-inline text-only hotspot:
    `1317224900 / 1947406800 / 1818922200 / 1852340000 / 1782194700 ns/op`
  - `00012389.xlsx` repeat-aware rerun:
    `ms=3238 min=2461 max=3346`, `text=11148655`, `images=0`
  - `testRecordSizeExceeded.xlsx` repeat-aware rerun:
    `ms=3887 min=2984 max=3985`, `text=181333430`, `images=0`

Interpretation:
- once again, a local fast-path win did not survive the real extraction workloads

Decision:
- Reverted.

## 2026-07-06 component split: `testRecordSizeExceeded.xlsx` simple-inline markdown cells are uniform plain-text values

Experiment:
- collect representative simple-inline cell texts from `testRecordSizeExceeded.xlsx`
- measure markdown-stage shape and component cost:
  - `cleanMarkdownTableCellValue(...)`
  - `prepareMarkdownTableCellValue(...)`
- count how many values are multiline or control-like after trimming

Observed results:
- representative markdown-value shape:
  - `values=3000000`
  - `cleaned=3000000`
  - `multiline=0`
  - `controlLike=0`
- component timings:
  - `cleanMarkdownTableCellValue`: about `637-1073 ns/op`
  - `prepareMarkdownTableCellValue`: about `247-266 ns/op`

Interpretation:
- the markdown branch for this workload is also extremely uniform
- `cleanMarkdownTableCellValue(...)` is clearly the heavier markdown-stage component, and the
  plausible future target is its single-line plain-text common path

Decision:
- Record as guidance only; temporary benchmark removed after measurement.

## 2026-07-06 no-op probe: `testRecordSizeExceeded.xlsx` markdown plain path has no `>80` fast window

Experiment:
- test whether the single-line plain-text branch of `cleanMarkdownTableCellValue(...)` could use a
  very conservative length-based shortcut:
  - only skip the expensive checks when `len(value) > 80`
  - and when the value has no obvious marker characters

Observed results:
- representative value shape:
  - `values=3000000`
  - `single=3000000`
  - `over80=0`
  - `over80NoMarkers=0`
- parity check passed for the temporary candidate
- microbenchmark numbers looked better:
  - current: `810.0 / 827.7 / 762.5 / 762.4 / 889.2 ns/op`
  - candidate: `662.3 / 764.2 / 651.4 / 649.7 / 649.2 ns/op`

Interpretation:
- because the `len > 80` branch never actually fires on the representative values, this candidate
  does not reveal a real causal optimization opportunity for the current workload
- the useful outcome is negative information: the easiest length-based shortcut does not exist here

Decision:
- Not applied; record as narrowing evidence only.

## 2026-07-06 rejected: no-marker single-line fast return in `cleanMarkdownTableCellValue(...)`

Experiment:
- in the single-line branch of `cleanMarkdownTableCellValue(...)`, add an early return when the
  cleaned value has:
  - no `/ \ % [ ] ( ) < > # = :`
  - no `rid` fragment
- keep the old hidden/control checks for every other case

Why:
- representative probes showed the `testRecordSizeExceeded.xlsx` simple-inline markdown values are
  all single-line and all marker-free on this discriminator
- representative parity also held on sampled `00012389.xlsx` worksheet values

Validation:
- representative parity on:
  - `testRecordSizeExceeded.xlsx`
  - `00012389.xlsx`
- representative microbenchmark on `testRecord` values
- integrated checks after applying:
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx -count=5`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx -count=5`
  - `cmd/compatcheck -repeat 5 ... testRecordSizeExceeded.xlsx`
  - `cmd/compatcheck -repeat 5 ... 00012389.xlsx`

Observed results:
- representative parity passed
- representative microbenchmark looked strong:
  - current: `606.5 / 880.4 / 803.5 / 768.2 / 848.3 ns/op`
  - candidate: `429.0 / 482.4 / 442.7 / 463.2 / 454.4 ns/op`
- but integrated results still regressed or stayed too weak:
  - `testRecordSizeExceeded.xlsx` simple-inline markdown hotspot:
    `2739514200 / 3057031300 / 2876826900 / 4471219700 / 2620501000 ns/op`
  - `00012389.xlsx` worksheet hotspot:
    `1525336200 / 1965797100 / 1945455300 / 1892368200 / 1888481900 ns/op`
  - `testRecordSizeExceeded.xlsx` repeat-aware rerun:
    `ms=3830 min=3462 max=5269`, `text=181333430`, `images=0`
  - `00012389.xlsx` repeat-aware rerun:
    `ms=3141 min=2426 max=3339`, `text=11148655`, `images=0`

Interpretation:
- this is another clear example of a strong local markdown-cell win failing to survive integrated
  extraction validation

Decision:
- Reverted.

## 2026-07-06 current simple-inline markdown hotspot split for `testRecordSizeExceeded.xlsx`

Experiment:
- re-profile `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx` with
  `-benchtime=3x` after the recent markdown common-path experiments

Observed results:
- focused benchmark:
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx`: `2329455067 ns/op`
- profile highlights:
  - `appendSimpleInlineWorksheetTextPrepared(...)`: `75.21%` cumulative
  - `appendWorksheetValue(...)`: `24.55%` cumulative
  - `cleanMarkdownTableCellValue(...)`: `14.96%` cumulative
  - `prepareMarkdownTableCellValue(...)`: `6.86%` cumulative
  - `simpleInlineCellText(...)`: `12.15%` cumulative
- `simpleInlineWorksheetCandidate(...)` still shows `15.54%` flat in this profile, but that is
  partly benchmark-setup pollution rather than timed-loop work

Interpretation:
- the markdown path is still an aggregate pipeline problem inside
  `appendSimpleInlineWorksheetTextPrepared(...)`
- the most plausible remaining structural costs are still the text append path plus the two
  markdown-cell stages

Decision:
- Record as hotspot guidance only.

## 2026-07-06 rejected: plain-text direct return in `prepareMarkdownTableCellValue(...)`

Experiment:
- for single-line `prepareMarkdownTableCellValue(...)` inputs with no:
  - backslash
  - pipe
  - newline
  - RTF prefix
- return `normalizeMarkdownTextLine(trimmed)` directly and skip the two escape `ReplaceAll` calls

Why:
- representative shape on `testRecordSizeExceeded.xlsx` was perfectly uniform:
  - `inputs=3000000`
  - `noSpecial=3000000`
  - `hasSlash=0`
  - `hasPipe=0`
  - `hasNewline=0`
  - `hasRTF=0`
- representative parity also held on sampled `00012389.xlsx` prepare-stage values

Validation:
- representative parity on:
  - `testRecordSizeExceeded.xlsx`
  - `00012389.xlsx`
- representative microbenchmark on `testRecord`
- integrated checks after applying:
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx -count=5`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx -count=5`
  - `cmd/compatcheck -repeat 5 ... testRecordSizeExceeded.xlsx`
  - `cmd/compatcheck -repeat 5 ... 00012389.xlsx`

Observed results:
- representative parity passed
- representative microbenchmark looked somewhat favorable:
  - current: `324.0 / 272.2 / 253.9 / 249.4 / 261.0 ns/op`
  - candidate: `244.0 / 248.7 / 297.5 / 242.8 / 240.4 ns/op`
- but integrated results regressed:
  - `testRecordSizeExceeded.xlsx` simple-inline markdown hotspot:
    `4822266600 / 3355737000 / 3236953000 / 2938747300 / 2391371100 ns/op`
  - `00012389.xlsx` worksheet hotspot:
    `2011261400 / 3265431300 / 1996105300 / 2180943500 / 2158418600 ns/op`
  - `testRecordSizeExceeded.xlsx` repeat-aware rerun:
    `ms=4246 min=3563 max=5623`, `text=181333430`, `images=0`
  - `00012389.xlsx` repeat-aware rerun:
    `ms=3412 min=2758 max=5034`, `text=11148655`, `images=0`

Interpretation:
- even a prepare-stage shortcut with near-perfect representative shape failed integrated
  validation, which reinforces how cautious the retention bar needs to be here

Decision:
- Reverted.

## 2026-07-06 rejected: row-slice markdown row builder for simple-inline `testRecord`

Experiment:
- replace the simple-inline markdown row accumulator in
  `appendSimpleInlineWorksheetTextPrepared(...)` from `map[int]string` to `[]string`
- append prepared cells by `cellCol-1` directly and trim trailing empty cells only at row flush

Why:
- representative worksheet shape on `testRecordSizeExceeded.xlsx` was extremely regular:
  - `rows=200000`
  - `cells=3000000`
  - `maxCol=15`
  - `maxRowWidth=15`
  - `parsedRefs=3000000`
  - `rowStartsAt1=200000`
  - `sequentialWithinRow=2800000`

Validation:
- temporary parity check against the current implementation
- same-tree benchmark:
  - `BenchmarkXLSXSimpleInlineRowSliceVariants -count=3`
- integrated checks after applying:
  - `BenchmarkXLSXSimpleInlineHotspots/...testRecordSizeExceeded.xlsx -count=5`
  - `cmd/compatcheck -repeat 5 ... testRecordSizeExceeded.xlsx`
  - `go test ./...`

Observed results:
- parity passed
- same-tree benchmark was promising:
  - current:
    `2381263500 / 3238469000 / 2584495000 ns/op`
    `1542 MB`, `7.900M allocs/op`
  - row-slice:
    `2272514900 / 2587540500 / 2547998800 ns/op`
    `1472 MB`, `7.600M allocs/op`
- integrated benchmark after applying remained mixed:
  - `2401889200 / 3072494800 / 3012996800 / 4022670400 / 2890919300 ns/op`
  - `1472 MB`, `7.600M allocs/op`
- repeat-aware single-file rerun still did not beat the retained current baseline strongly enough:
  - `ms=3957 min=2969 max=4342`, `text=181333430`, `images=0`
- `go test ./...` passed after revert:
  - `ok officeread 182.484s`

Interpretation:
- this is the strongest recent structural candidate, but the integrated `compatcheck` result is
  still not convincing enough to keep

Decision:
- Reverted.

## 2026-07-06 rejected: reusable row buffer for `appendWorksheetText(...)` markdown rows

Idea:
- replace the per-row `map[int]string` used by `appendWorksheetText(...)` markdown preparation with
  a reusable fixed-width row buffer
- try to remove repeated map allocation and `mapassign_fast64` churn from the `00012389.xlsx`
  markdown-prep path

Why:
- after confirming that markdown preparation is a major extra cost on `00012389.xlsx`, a reusable
  row buffer looked like a more structural attack than helper-only micro-tuning

Validation:
- focused worksheet benchmark on `00012389.xlsx` with `-count=5`
- isolated repeat-aware rerun:
  - `cmd/compatcheck -repeat 5 ... 00012389.xlsx`
- repository regression after revert:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused worksheet benchmark again improved allocations, but wall time stayed mixed:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    `1204844100 / 1200706100 / 1529469600 / 1532947200 / 1597046600 ns/op`
  - alloc profile improved versus the current-baseline markdown-prep benchmark:
    - about `515-516 MB`, `~9.73M allocs/op`
    - earlier baseline was about `562-563 MB`, `~9.94M allocs/op`
- isolated rerun still did not beat the earlier current-baseline evidence convincingly:
  - `perf-exp-ai-assistant-00012389-rowbuffer.json`
  - `00012389.xlsx`: `2241 ms`, `min=2050`, `max=2592`, `textBytes=11148655`, `images=0`
- repository regression stayed green after revert:
  - `ok officeread (cached)`

Interpretation:
- the row-buffer idea is probably structurally closer to the real cost than the earlier helper
  tweaks, since alloc pressure dropped again in a meaningful way
- but by the current retention bar, the wall-time evidence is still too unstable

Decision:
- Reverted.

## 2026-07-06 current finding: markdown preparation is the dominant extra cost on `00012389.xlsx`

Observation:
- a direct benchmark comparison of the same worksheet path with and without markdown preparation
  shows that markdown work is now a dominant extra cost on `00012389.xlsx`

Evidence:
- worksheet benchmark with markdown preparation enabled:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    `1256466300 / 1242155900 / 1320314700 ns/op`
- temporary isolated benchmark on the same path with `appendWorksheetText(..., nil)`:
  - `BenchmarkXLSXWorksheetTextNoMarkdown00012389`:
    `800303200 / 863024900 / 820619900 ns/op`
- allocation delta:
  - with markdown prep: about `562-563 MB`, `~9.94M allocs/op`
  - without markdown prep: about `456-457 MB`, `~8.37M allocs/op`

Interpretation:
- for `00012389.xlsx`, markdown preparation inside `appendWorksheetText(...)` is a first-order
  cost now
- that explains why recent XML-loop micro-optimizations looked weak: they were pushing on a smaller
  piece than the markdown-prep work sitting on top of the same path
- next useful work for this file likely needs to target markdown row / cell preparation directly, or
  provide a semantically acceptable way to avoid paying that cost when structured markdown is not
  required

Decision:
- Record as guidance for subsequent rounds; temporary benchmark removed after measurement.

## 2026-07-06 rejected: case-fold prefix check shortcut in `prepareMarkdownTableCellValue(...)`

Idea:
- replace the per-cell `strings.ToLower(trimmed)` allocation in the RTF-prefix check with
  `hasPrefixFold(...)`
- keep the same `{\rtf` / `\rtf` semantics and change nothing else in markdown cell preparation

Why:
- after confirming that markdown preparation is a major extra cost on `00012389.xlsx`, this was a
  low-risk attempt to trim obvious helper allocation churn inside that path

Validation:
- focused worksheet benchmark on `00012389.xlsx` with `-count=5`
- isolated repeat-aware rerun:
  - `cmd/compatcheck -repeat 5 ... 00012389.xlsx`
- repository regression after revert:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused worksheet benchmark showed lower allocations but mixed timing:
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    `1222076400 / 1208177300 / 1447019500 / 1441599500 / 1468807700 ns/op`
  - alloc profile improved versus the earlier current-baseline markdown-prep benchmark:
    - about `550-551 MB`, `~9.29M allocs/op`
    - earlier baseline was about `562-563 MB`, `~9.94M allocs/op`
- isolated rerun did not improve convincingly:
  - `perf-exp-ai-assistant-00012389-rtf-foldcheck.json`
  - `00012389.xlsx`: `2255 ms`, `min=1990`, `max=2434`, `textBytes=11148655`, `images=0`
- repository regression stayed green after revert:
  - `ok officeread (cached)`

Interpretation:
- the change clearly helped allocation behavior, but the throughput story was too noisy and too weak
  to retain under the current bar
- this is a useful data point: markdown-prep helpers can move alloc counts meaningfully, but we
  still need a cleaner wall-time win before keeping one

Decision:
- Reverted.

## 2026-07-06 current hotspot split: `00012389.xlsx` and `testRecordSizeExceeded.xlsx` are now different `.xlsx` problems

Observation:
- a fresh current-baseline repro shows the two retained `.xlsx` hotspot files are no longer led by
  the same path

Evidence:
- focused current baseline measurements:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2201956800 ns/op`
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `2908301300 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1228603400 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `2372883900 ns/op`
- `00012389.xlsx` profile:
  - top cumulative frame is `appendWorksheetText` at about `1.60s cum`
  - most cost is in `encoding/xml.Decoder`, XML tokenization, and the general worksheet reader
- `testRecordSizeExceeded.xlsx` profile:
  - `appendSimpleInlineWorksheetTextPrepared(...)` remains the core path at about `1.94s cum`
  - `simpleInlineWorksheetCandidate(...)` is still the top flat frame at about `0.45s`
  - `cleanTextFastPath(...)`, `maybeControlFragmentText(...)`, and `simpleInlineCellText(...)`
    are still material inside the simple-inline path

Interpretation:
- `testRecordSizeExceeded.xlsx` is still the retained giant-inlineStr simple-inline workload
- `00012389.xlsx` has largely become a general XML-decoder worksheet workload again
- so future `.xlsx` work should stop assuming one micro-optimization can move both files together
- next useful directions likely split cleanly:
  - simple-inline-specific work for `testRecordSizeExceeded.xlsx`
  - XML-decoder / worksheet-reader work for `00012389.xlsx`

Decision:
- Record as guidance for the next rounds; no code change from this finding alone.

## 2026-07-06 rejected: remove duplicate `hiddenColumnCell(...)` check from `appendWorksheetText(...)`

Idea:
- in the general worksheet XML path, remove `hiddenColumnCell(cellRef, hiddenCols)` from
  `skipCell` once `cellCol` is already known
- leave only `rowHidden || columnHidden(cellCol, hiddenCols)` for the common cell skip decision

Why:
- `appendWorksheetText(...)` already parses `cellRef` and computes `cellCol`
- for `00012389.xlsx`, which is now mainly a general XML-decoder worksheet workload, this looked
  like a more structural duplicate check than the recent parser micro-optimizations

Validation:
- focused `00012389.xlsx` benchmarks with `-count=3`
- isolated repeat-aware rerun:
  - `cmd/compatcheck -repeat 3 ... 00012389.xlsx`
- repository regression after revert:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- isolated rerun looked slightly better:
  - `perf-exp-ai-assistant-00012389-no-hiddenColumnCell.json`
  - `00012389.xlsx`: `2122 ms`, `min=2048`, `max=2518`, `textBytes=11148655`, `images=0`
- but the focused benchmark signal stayed mixed:
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    `2121202300 / 2302127400 / 2340548000 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    `1234948500 / 1284801800 / 1317203100 ns/op`
- repository regression stayed green after revert:
  - `ok officeread (cached)`

Interpretation:
- this is exactly the kind of change that can feel plausible and even look slightly better in an
  isolated rerun, but still fail the stronger repeat-aware benchmark bar
- the evidence was not strong enough to justify keeping the code change

Decision:
- Reverted.

## 2026-07-06 rejected: byte-specialized `cellRefIndexesBytes(...)` in the simple-inline `.xlsx` path

Idea:
- add a `[]byte`-specialized parser for cell references in the simple-inline worksheet path
- avoid converting `xmlAttrBytes(tag, "r")` to `string` before parsing

Why:
- the latest simple-inline profiles still showed `cellRefIndexes(...)` plus
  `runtime.slicebytetostring` on the hot path
- this looked like a neat low-risk way to remove a conversion on every inline cell

Validation:
- focused `.xlsx` hotspot benchmarks with `-benchtime=1x`
- repeat-aware focused pair rerun with
  `cmd/compatcheck -repeat 3`
- repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused benchmark view regressed:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `2768071800 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2726203100 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `3483308200 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1749171500 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1786622000 ns/op`
- repeat-aware pair rerun also regressed, although output parity held:
  - `perf-exp-ai-assistant-cellref-bytes-xlsx-pair.json`
  - `00012389.xlsx`: `2599 ms`, `textBytes=11148655`, `images=0`
  - `testRecordSizeExceeded.xlsx`: `3559 ms`, `textBytes=181333430`, `images=0`
  - total `.xlsx`: `6158 ms`
- repository regression stayed green:
  - `ok officeread 131.562s`
  - after revert: `ok officeread (cached)`

Interpretation:
- the specialized byte parser ended up slower than the existing `string` path on the real retained
  workload
- this is another reminder that shaving an allocation-shaped operation is not automatically a win
  once the whole loop shape changes

Decision:
- Reverted.

## 2026-07-06 rejected: `<t>` tag fast path in `simpleInlineCellText(...)`

Idea:
- add a narrow fast path in `simpleInlineCellText(...)` for the common non-namespaced `<t>` tag
- only fall back to `xmlStartTagNameIs(...)` for namespaced or suspicious cases

Why:
- recent `.xlsx` simple-inline profiles still showed `simpleInlineCellText(...)` as a visible cost
- `xmlStartTagNameIs(...)` itself was still taking measurable time, so this looked like a small
  mechanical parser win

Validation:
- focused `.xlsx` hotspot benchmarks with `-benchtime=1x`
- repeat-aware focused pair rerun with
  `cmd/compatcheck -repeat 3`
- repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused benchmark view regressed:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `2742973800 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2688273900 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `3568474200 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1763604600 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1618000500 ns/op`
- repeat-aware pair rerun also regressed, even though output parity held:
  - `perf-exp-ai-assistant-inlinecell-tagfast-xlsx-pair.json`
  - `00012389.xlsx`: `2201 ms`, `textBytes=11148655`, `images=0`
  - `testRecordSizeExceeded.xlsx`: `3510 ms`, `textBytes=181333430`, `images=0`
  - total `.xlsx`: `5711 ms`
- repository regression stayed green:
  - `ok officeread 126.086s`
  - after revert: `ok officeread (cached)`

Interpretation:
- the fast path was too clever for the gain it tried to buy
- the extra branch shape cost more than the saved generic tag check on the real retained workload
- this is a clean semantics-safe rejection: the change simply did not help

Decision:
- Reverted.

## 2026-07-06 rejected: first-byte gating for `maybeControlFragmentText(...)`

Idea:
- add a first-byte gate inside `maybeControlFragmentText(...)`
- only call the expensive OLE / OOXML fragment helpers for a narrower set of leading-byte families,
  while leaving the existing keyword switch in place

Why:
- the latest `.xlsx` simple-inline profile still showed
  `cleanTextFastPathControlFragment(...)` and `maybeControlFragmentText(...)` as a meaningful cost
- most worksheet values are ordinary text, so paying for all of the fragment helpers on every
  5-80 byte candidate looked wasteful

Validation:
- focused `.xlsx` hotspot benchmarks with `-benchtime=1x`
- repeat-aware focused pair rerun with
  `cmd/compatcheck -repeat 3`
- CPU profile on
  `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx`
- repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- the performance signal looked strong at first:
  - one-shot focused view:
    - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `2532887600 ns/op`
    - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2147824100 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `2993073600 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1397696300 ns/op`
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1126569100 ns/op`
  - repeat-aware focused rerun:
    - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`:
      `2464512900 / 2609855300 / 2860446000 ns/op`
    - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
      `2397538500 / 2455047400 / 2429942900 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`:
      `2818320200 / 2689491300 / 2732958700 ns/op`
    - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
      `1495584300 / 1461816400 / 1525629500 ns/op`
    - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`:
      `1370466300 / 1333364900 / 1392958600 ns/op`
  - repeat-aware pair rerun kept output parity and improved total time:
    - `perf-exp-ai-assistant-controlfragment-gate-xlsx-pair.json`
    - `00012389.xlsx`: `1978 ms`, `textBytes=11148655`, `images=0`
    - `testRecordSizeExceeded.xlsx`: `3097 ms`, `textBytes=181333430`, `images=0`
    - total `.xlsx`: `5075 ms`
  - post-change profile also looked better:
    - `cleanTextFastPathControlFragment(...)`: about `0.07s cum`
    - `maybeControlFragmentText(...)`: about `0.06s cum`
- but repository tests proved the change was wrong:
  - `TestOLEClassFragmentsAreFilteredWithoutDroppingProse` failed because `CompObj` was no longer
    filtered
  - `TestOLEIdentifierFragmentsAreFilteredConservatively` failed because
    `CLSID={00020906-0000-0000-C000-000000000046}` was no longer filtered
  - failing run:
    - `FAIL officeread 129.224s`
- after reverting the experiment, repository regression returned green:
  - `ok officeread (cached)`

Interpretation:
- the benchmark win came from making the filter too weak
- since it leaked exactly the OLE/control fragments the filter exists to suppress, this experiment
  cannot be kept regardless of the speedup
- this is a nice example of why the acceptance bar has to include semantic regression tests, not
  just hotspot timing

Decision:
- Reverted.

## 2026-07-06 retained: cheaper byte checks in `simpleInlineWorksheetCandidate(...)` for large inlineStr `.xlsx`

Idea:
- keep the retained single-pass `simpleInlineWorksheetCandidate(...)` scan, but replace the hot
  `bytes.HasPrefix(...)` checks with direct fixed-width byte comparisons
- keep the same rejection markers and the same `foundInlineStr` success condition

Why:
- a fresh `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx` profile still
  showed `simpleInlineWorksheetCandidate(...)` at the top:
  - before the change: `1.16s flat`, `2.32s cum`
  - `bytes.HasPrefix` / `bytes.Equal` / `memeqbody` were still taking a large share of the scan
- the retained single-pass scan had already removed the extra pass, so the remaining obvious cost
  was the generic prefix machinery itself

Validation:
- focused `.xlsx` hotspot benchmarks with `-benchtime=1x`
- repeat-aware focused pair rerun with
  `cmd/compatcheck -repeat 3`
- CPU profile on
  `BenchmarkXLSXSimpleInlineTextOnlyHotspots/...testRecordSizeExceeded.xlsx`
- repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- same-turn focused baseline before the change:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `3814173000 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `2513019300 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `3780165300 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1403364200 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1639419400 ns/op`
- repeat-aware focused rerun after the change:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`:
    `2726812800 / 3230328300 / 3328964500 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    `2588378900 / 2622103500 / 2620755000 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`:
    `3053558500 / 3374688400 / 2709811100 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    `1248280500 / 1232240400 / 1217979500 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`:
    `1252283200 / 1216756100 / 1238279800 ns/op`
- repeat-aware pair rerun kept output parity and improved total time:
  - previous retained pair report:
    `perf-exp-ai-assistant-worksheetvalue-xlsx-pair.json`
    - total `.xlsx`: `7586 ms`
    - `00012389.xlsx`: `2535 ms`, `textBytes=11148655`, `images=0`
    - `testRecordSizeExceeded.xlsx`: `5051 ms`, `textBytes=181333430`, `images=0`
  - current byte-check rerun:
    `perf-exp-ai-assistant-candidate-bytechecks-xlsx-pair.json`
    - total `.xlsx`: `5765 ms`
    - `00012389.xlsx`: `2463 ms`, `textBytes=11148655`, `images=0`
    - `testRecordSizeExceeded.xlsx`: `3302 ms`, `textBytes=181333430`, `images=0`
- post-change profile:
  - `simpleInlineWorksheetCandidate(...)`: `0.94s flat`, `0.94s cum`
- repository regression stayed green:
  - `ok officeread 119.878s`

Interpretation:
- the retained single-pass scan was still paying too much for generic prefix helpers
- direct byte checks keep the same semantics but materially reduce the candidate-scan cost on the
  large inlineStr path
- `00012389.xlsx` extraction-level timing moved a little in the focused benchmark, but the
  repeat-aware compat rerun and worksheet-level timings stayed net-positive

Decision:
- Retained.

## 2026-07-06 rejected: skip the second trim in `appendWorksheetValue(...)`

Idea:
- narrow the worksheet-only path by assuming `cleanText(...)` already returns trimmed output
- after `appendWorksheetValue(...)` calls `cleanText(value)`, write the result directly with
  `appendTrimmedTextBlock(...)` instead of `appendCleanedTextBlock(...)`

Why:
- after the retained candidate-scan optimization, the next `.xlsx` hotspot cluster was
  `appendWorksheetValue(...)` plus `cleanText(...)`
- this looked like a cheap way to remove one more `strings.TrimSpace(...)` from the hottest
  simple-inline worksheet path

Validation:
- focused `.xlsx` hotspot benchmarks with `-benchtime=1x`
- repeat-aware focused pair rerun with
  `cmd/compatcheck -repeat 3`
- repository regression after revert:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- the focused pair rerun kept output parity and looked faster:
  - retained pair report:
    `perf-exp-ai-assistant-candidate-bytechecks-xlsx-pair.json`
    - total `.xlsx`: `5765 ms`
    - `00012389.xlsx`: `2463 ms`, `textBytes=11148655`, `images=0`
    - `testRecordSizeExceeded.xlsx`: `3302 ms`, `textBytes=181333430`, `images=0`
  - trim-skip rerun:
    `perf-exp-ai-assistant-worksheet-trimskip-xlsx-pair.json`
    - total `.xlsx`: `5415 ms`
    - `00012389.xlsx`: `2177 ms`, `textBytes=11148655`, `images=0`
    - `testRecordSizeExceeded.xlsx`: `3238 ms`, `textBytes=181333430`, `images=0`
- but the decisive repeat-aware hotspot benchmarks regressed:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`:
    `2735872500 / 2929788700 / 3210416300 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`:
    `2723029000 / 2643780000 / 2673465600 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`:
    `3043613900 / 3037627700 / 3182032000 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`:
    `1594783500 / 1658658300 / 1983508600 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`:
    `2024724800 / 1979255100 / 1993959500 ns/op`
- repository regression stayed green after revert:
  - `ok officeread 136.639s`

Interpretation:
- the pair rerun was not strong enough evidence on its own
- the more targeted repeat-aware hotspot harness showed this change moved the retained `.xlsx`
  path in the wrong direction, especially on the `simpleInlineTextOnly` hotspot
- this is a good example of why the decision standard has to prefer the focused benchmark harness
  over a small end-to-end pair that can drift with noise

Decision:
- Reverted.

## 2026-07-06 revisited and retained: single-pass `simpleInlineWorksheetCandidate(...)` scan for huge inlineStr worksheets

- idea:
  - keep the existing `.xlsx` simple-inline fast path semantics unchanged
  - only replace the worksheet candidate precheck:
    - before: repeated `bytes.Contains(...)` scans over the full worksheet XML for each allowed and
      forbidden marker
    - after: a single byte scan that detects:
      - required marker: `t="inlineStr"`
      - forbidden markers: `<f`, `<v`, `<cols`, `hidden`, `hyperlink`, `dataValidation`,
        `Header`, `Footer`, `<rPh`, `<phoneticPr`, ` ht=`, ` ht =`, `\tht=`, `\nht=`, `\rht=`,
        `t="s"`, `t="b"`, `t="str"`, and `<!DOCTYPE`

- rationale:
  - fresh profile on the current retained baseline for
    `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx` showed that
    `simpleInlineWorksheetCandidate(...)` was still dominating the text-only inline path:
    - previous candidate profile: about `3.10 s` cumulative inside
      `simpleInlineWorksheetCandidate(...)`
    - current retained baseline before this change: about `2.30 s` cumulative, still the largest
      single hot frame in the profile
  - that meant the worksheet candidate check itself was still spending too much time rescanning the
    same giant XML payload
  - this revisits an earlier rejected 2026-07-05 direction with a newer baseline and fresh profile
    evidence; the current retained result is based on the latest integrated reruns, not the older
    rejection context

- validation:
  - focused `.xlsx` hotspot benchmarks:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSXHotspots|BenchmarkXLSXSimpleInlineHotspots|BenchmarkXLSXWorksheetTextHotspots|BenchmarkXLSXSimpleInlineTextOnlyHotspots' -benchmem -benchtime=1x ./`
  - profile:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench '^BenchmarkXLSXSimpleInlineTextOnlyHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchtime=1x -cpuprofile tmp-shape/xlsx-simpleinline-textonly-candidate-fast.prof ./`
    - `& 'C:\Program Files\Go\bin\go.exe' tool pprof -top ./officeread.test.exe tmp-shape/xlsx-simpleinline-textonly-candidate-fast.prof`
  - repeat-aware mixed keyset rerun:
    - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xlsx-inline-candidate-scan-keyset.json -csv testdata/web-samples/reports/perf-exp-xlsx-inline-candidate-scan-keyset.csv <current-keyset>`
  - repository regression:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

- focused benchmark results:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: `4474042700 ns/op -> 3553883100 ns/op`
  - `BenchmarkExtractXLSXHotspots/00012389.xlsx`: `3189833500 ns/op -> 2209982100 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: `2039197300 ns/op -> 2035989800 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/testRecordSizeExceeded.xlsx`: `3755042400 ns/op -> 3274889600 ns/op`
  - `BenchmarkXLSXWorksheetTextHotspots/00012389.xlsx`: `1548595600 ns/op -> 1293209700 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: `1178737800 ns/op -> 1158989000 ns/op`

- profile result:
  - `simpleInlineWorksheetCandidate(...)` remained hot, but its cumulative cost dropped further to
    about `2.30 s` in the new profile instead of the earlier `~3.10 s` range

- repeat-aware mixed keyset results:
  - `.docx`: `6931 ms -> 6214 ms`
  - `.xls`: `17830 ms -> 11725 ms`
  - `.xlsx`: `23756 ms -> 21647 ms`
  - output parity held on all 30 inputs

- repository regression:
  - `ok officeread 124.196s`

Interpretation:
- this is a good example of a low-risk `.xlsx` optimization: it only changes how the candidate fast
  path is recognized, not how worksheet text is extracted once the path is chosen
- the strongest direct benefit is on huge inlineStr worksheets and on the full `.xlsx` extract path
  that depends on recognizing them efficiently
- the `.docx` and `.xls` improvements in the mixed rerun should be treated as ordinary run variance,
  because this change only touches `.xlsx`

Decision:
- Retained.

## 2026-07-06 measured: coverage-side cell entries are mostly real additions on the heavy `008055/016161` `.xls` samples

- measurement:
  - temporary coverage contribution probe:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSCoverageCellContribution$' -v ./`
  - temporary length-bucket probe:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSCoverageCellContributionByLength$' -v ./`

- observed coverage contribution:
  - `006087.xls`:
    - `rawAdds=48582`
    - `visibleAdds=0`
    - `cellTotal=235659`
    - `cellNew=37952`
    - `cellDup=197707`
    - `finalCoverage=130515`
  - `008055.xls`:
    - `rawAdds=16917`
    - `visibleAdds=0`
    - `cellTotal=1313143`
    - `cellNew=1288825`
    - `cellDup=24318`
    - `cellComparableNew=2`
    - `cellComparableDup=74`
    - `finalCoverage=1322621`
  - `016161.xls`:
    - `rawAdds=45630`
    - `visibleAdds=0`
    - `cellTotal=551263`
    - `cellNew=520599`
    - `cellDup=30664`
    - `cellComparableNew=2`
    - `finalCoverage=611643`

- observed length split:
  - `006087.xls`:
    - `<=8`: `37653 new / 127487 dup`
    - `9-16`: `164 new / 19594 dup`
    - `17-32`: `74 new / 50607 dup`
  - `008055.xls`:
    - `<=8`: `47813 new / 6920 dup`
    - `9-16`: `604586 new / 15520 dup`
    - `17-32`: `636338 new / 1781 dup`
  - `016161.xls`:
    - `<=8`: `49616 new / 10394 dup`
    - `9-16`: `9744 new / 673 dup`
    - `17-32`: `461239 new / 19597 dup`

Interpretation:
- this clearly separates coverage from exact-set:
  - exact-set had enough short-cell duplication to justify a narrow short-cell cache
  - coverage does not look like that on the heavy `008055/016161` samples
- on those two samples, most coverage-side table-cell entries are genuinely new keys, especially in
  the `9..32` byte range
- that means any attempt to skip or coarsely prune coverage-side cells by short-length or generic
  duplicate heuristics is likely to break the useful part of coverage rather than just removing
  redundant work

Decision:
- No code change from this measurement.

## 2026-07-06 measured: only `006087.xls` reaches the second stage, and its remaining queries are purely `visible` hits

- measurement:
  - query-resolution probe:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSQueryResolution$' -v ./`
  - `006087.xls` second-stage timing probe:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLS006087SecondStageTiming$' -v ./`

- observed query-resolution split:
  - `006087.xls`:
    - `totalAllowed=41339`
    - `exact=40935`
    - `coverage=0`
    - `table=0`
    - `visible=404`
    - `miss=0`
  - `008055.xls`:
    - `totalAllowed=1181169`
    - `exact=1181169`
    - `coverage=0`
    - `table=0`
    - `visible=0`
    - `miss=0`
  - `016161.xls`:
    - `totalAllowed=461603`
    - `exact=461603`
    - `coverage=0`
    - `table=0`
    - `visible=0`
    - `miss=0`

- observed `006087.xls` second-stage timing:
  - full `coverage + containment` build: `469 ms`
  - remaining query resolution: `20 ms`

Interpretation:
- this is the first strong evidence that the retained `.xls` workload is split into two different
  post-exact shapes:
  - `008055.xls` and `016161.xls` finish completely in exact precheck
  - `006087.xls` is the only retained heavy sample that falls through, and when it does, the
    remaining queries resolve entirely via `visibleLineContainsLine(...)`
- so if another optimization in this area is going to survive, it likely has to be a very narrow
  `006087`-shape fallback, not a generic change to the shared second-stage path

Decision:
- No code change from this measurement.

## 2026-07-06 corrected and retained: `006087`-shape exact+visible-only fallback before full second-stage build

- experiment:
  - after exact precheck fails, add a very narrow `.xls`-only fallback for the `006087`-style shape:
    - at least `30000` normalized source lines
    - source text at most `512 KiB`
  - on that path, first try `exact OR visible-only` resolution:
    - keep the existing exact-set
    - build only visible containment
    - if every remaining line is covered by exact or visible-only matching, return immediately
    - otherwise fall back to the full retained `coverage + table + visible` second stage

- rationale:
  - fresh query-resolution measurement showed that `006087.xls` was the only retained heavy sample
    reaching the second stage, and its remaining `404` queries were all visible hits
  - a direct timing probe then showed that the expensive part on that sample was building the full
    second-stage structures, not answering the queries
  - the visible-only precheck itself covered `006087.xls` successfully in about `152 ms`

- validation:
  - formatting:
    - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
  - focused hotspot benchmarks:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/006087.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
  - repeat-aware `.xls6` rerun:
    - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xls-visible-only-fallback-xls6.json -csv testdata/web-samples/reports/perf-exp-xls-visible-only-fallback-xls6.csv <xls6>`
  - repeat-aware mixed keyset rerun:
    - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xls-visible-only-fallback-keyset.json -csv testdata/web-samples/reports/perf-exp-xls-visible-only-fallback-keyset.csv <current-keyset>`
  - repository regression before revert:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

- focused hotspot results:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `645141700 ns/op -> 167294500 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1625965900 ns/op -> 1665857100 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `638316500 ns/op -> 692025000 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `110346100 ns/op -> 104585600 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `699207500 ns/op -> 679812100 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `312707000 ns/op -> 420859900 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `504470600 ns/op -> 616614500 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1304217200 ns/op -> 1410249700 ns/op`

- repeat-aware `.xls6` results:
  - current retained baseline total: `10362 ms`
  - experiment total: `10106 ms`
  - per-file:
    - `002505.xls`: `1373 ms -> 1427 ms`
    - `006087.xls`: `1277 ms -> 1234 ms`
    - `008055.xls`: `4064 ms -> 3961 ms`
    - `016161.xls`: `1711 ms -> 1632 ms`
    - `019088.xls`: `993 ms -> 928 ms`
    - `019089.xls`: `944 ms -> 924 ms`
  - output parity held on all six files

- initial mixed-keyset result was invalid:
  - the first `perf-exp-xls-visible-only-fallback-keyset.json` run was executed in parallel with
    `go test ./...`
  - that concurrent CPU contention inflated the mixed-keyset numbers and made the experiment look
    like a regression outside the `.xls6` slice

- corrected repeat-aware mixed keyset rerun in isolation:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xls-visible-only-fallback-keyset-rerun-clean.json -csv testdata/web-samples/reports/perf-exp-xls-visible-only-fallback-keyset-rerun-clean.csv <current-keyset>`
  - corrected results against the current retained baseline:
    - `.docx`: `6931 ms -> 7035 ms`
    - `.xls`: `17830 ms -> 13237 ms`
    - `.xlsx`: `23756 ms -> 21139 ms`
  - output parity held on all 30 inputs

- repository regression in isolation:
  - `ok officeread 133.940s`

Interpretation:
- the first rejection was based on a polluted mixed-keyset rerun and should not be treated as
  authoritative
- once rerun in isolation, the experiment keeps its strong `.xls6` win and also improves the
  broader mixed keyset materially overall:
  - especially `.xls`, but also `.xlsx`
  - `.docx` regresses only slightly (`+104 ms` across 2 files), which is far smaller than the wins
    in the affected spreadsheet buckets
- this makes the fallback a valid retained optimization after all

Decision:
- Retained.

## 2026-07-06 measured: the visible-only fallback pattern covers all 24 exact-miss files in the 1000-file `.xls` corpus, but only 9 of them are large enough to justify the current gate

- measurement:
  - temporary corpus-scope probe over all `1000` downloaded `.xls` samples:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSVisibleOnlyScope$' -v ./`

- observed corpus split:
  - `total=1000`
  - `exactMiss=24`
  - current gate hits: `9`
  - current gate hits covered by `exact+visible-only`: `9`
  - additional exact-miss files outside the gate that still cover with `exact+visible-only`: `15`

- covered large-shape files under the current gate:
  - `000034-2.xls`
  - `000034.xls`
  - `000055.xls`
  - `001138.xls`
  - `006086.xls`
  - `006087.xls`
  - `007260.xls`
  - `014479.xls`
  - `022223.xls`

- exact-miss files that also cover with `exact+visible-only`, but are currently excluded by the
  size/line gate:
  - `000051.xls`
  - `003995.xls`
  - `004703.xls`
  - `006812.xls`
  - `009715.xls`
  - `009933.xls`
  - `010276.xls`
  - `010480.xls`
  - `010701.xls`
  - `015300.xls`
  - `018541.xls`
  - `022230.xls`
  - `022238.xls`
  - `022239.xls`
  - `022241.xls`

Interpretation:
- the visible-only fallback is not a single-file trick; it matches the full exact-miss subset of the
  current 1000-file `.xls` corpus
- but that does not automatically mean it should run on every exact-miss file, because many of the
  non-gated files are much smaller and therefore far more sensitive to the fallback's own setup cost

Decision:
- No code change from this measurement.

## 2026-07-06 rejected: broaden the visible-only fallback from the large-shape gate to every `.xls` exact miss

- experiment:
  - keep the retained `exact+visible-only` fallback logic itself unchanged
  - remove only the `006087`-style size/line gate
  - on `.xls`, after exact precheck fails, always try `exact+visible-only` before building the full
    retained second stage

- rationale:
  - the corpus-scope measurement showed that all `24` exact-miss files in the 1000-file `.xls`
    corpus can be covered by `exact+visible-only`
  - that made it tempting to let every `.xls` exact miss take the cheaper path

- validation:
  - formatting:
    - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
  - focused hotspot benchmarks:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/006087.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
  - repeat-aware `.xls6` rerun:
    - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xls-visible-only-all-exactmiss-xls6.json -csv testdata/web-samples/reports/perf-exp-xls-visible-only-all-exactmiss-xls6.csv <xls6>`

- focused hotspot results:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `167294500 ns/op -> 164642200 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1665857100 ns/op -> 1840787400 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `692025000 ns/op -> 932711500 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `104585600 ns/op -> 100448400 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `679812100 ns/op -> 641924700 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `420859900 ns/op -> 295391600 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `616614500 ns/op -> 479961000 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1410249700 ns/op -> 1273455200 ns/op`

- repeat-aware `.xls6` results against the current retained gated fallback:
  - retained total: `10106 ms`
  - experiment total: `10520 ms`
  - per-file:
    - `002505.xls`: `1427 ms -> 1433 ms`
    - `006087.xls`: `1234 ms -> 1445 ms`
    - `008055.xls`: `3961 ms -> 4077 ms`
    - `016161.xls`: `1632 ms -> 1708 ms`
    - `019088.xls`: `928 ms -> 948 ms`
    - `019089.xls`: `924 ms -> 909 ms`

Interpretation:
- the broader fallback does exactly what the corpus-scope measurement predicted semantically
- but it loses on runtime because many of the newly included exact-miss files are small enough that
  the visible-only preflight overhead is not repaid
- the current large-shape gate is therefore doing real performance work, not just narrowing coverage
- an additional isolated rerun on only the affected `24` exact-miss files confirmed the same point:
  - gated exact-miss subset: `14778 ms`
  - broadened exact-miss subset: `15418 ms`
  - delta: `+640 ms`
- this isolates the decision from any noise contributed by exact-hit `.xls` files and shows that the
  gate still pays for itself inside the exact-miss subset alone
- a still narrower isolated rerun on only the `15` out-of-gate exact-miss files with `repeat 7`
  showed the opposite extreme:
  - gated out-of-gate subset: `1029 ms`
  - broadened out-of-gate subset: `1024 ms`
  - delta: `-5 ms`
- that means the newly eligible small files do not create a meaningful win either; their aggregate
  benefit is noise-level even when isolated from the gated large-shape files

Decision:
- Reverted.

## 2026-07-06 measured: exact-set/source-line variant usage is too narrow to justify another heading-focused optimization

- measurement:
  - temporary source-line variant probe:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSExactVariantShape$' -v ./`
  - temporary exact-set markdown-shape probe:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSExactSetHashShape$' -v ./`

- observed source-line variant usage:
  - `006087.xls`:
    - `total=41347`
    - `allowed=41339`
    - `visibleAdds=0`
    - `markdownAdds=0`
    - `escapedAdds=0`
  - `008055.xls`:
    - `total=1181198`
    - `allowed=1181169`
    - `visibleAdds=0`
    - `markdownAdds=2`
    - examples:
      - `#ID_REF = -> ID_REF =`
      - `#VALUE = raw signal output from ArrayPro image analysis program -> VALUE = raw signal output from ArrayPro image analysis program`
  - `016161.xls`:
    - `total=461615`
    - `allowed=461603`
    - `visibleAdds=0`
    - `markdownAdds=2`
    - examples:
      - `#1376-2 -> 1376-2`
      - `#581 -> 581`

- observed exact-set markdown shape:
  - `008055.xls`:
    - `total=16917`
    - `plain=0`
    - `nonPlain=16917`
    - `startsHash=4`
    - `startsPipe=16910`
    - `changedMarkdown=7`
    - hash examples:
      - `## Sheet1 -> Sheet1`
      - `## Sheet2 -> Sheet2`
      - `## Sheet3 -> Sheet3`
      - `## Workbook Sheets -> Workbook Sheets`
  - `016161.xls`:
    - `total=45634`
    - `plain=0`
    - `nonPlain=45634`
    - `startsHash=7`
    - `startsPipe=45621`
    - `changedMarkdown=13`
    - hash examples:
      - `## chr1 -> chr1`
      - `## chr1_siRNAs -> chr1_siRNAs`
      - `## chr2 -> chr2`
      - `## chr2_siRNAs -> chr2_siRNAs`
      - `## Workbook Sheets -> Workbook Sheets`

Interpretation:
- this rules out another obvious micro-direction: the remaining exact-stage cost is not coming from
  a large family of heading-like source lines or markdown headings
- on the retained heavy `.xls` samples, source-line variant usage is almost entirely direct, and the
  rare `markdownVisibleLineText(...)` additions are just a couple of leading-`#` cases
- exact-set markdown input is overwhelmingly table rows (`|...|`), with only a handful of heading
  lines, so specializing `#...` handling would shave at most a tiny fraction of the work
- the cost center therefore remains table-row / cell expansion inside `markdownBackfillExactSet(...)`,
  not heading normalization

Decision:
- No code change from this measurement.

## 2026-07-06 retained: short-cell cache for `.xls` exact-set table-cell normalization

- idea:
  - keep the retained `.xls` exact/backfill structure intact
  - target only one repeated cost inside `markdownBackfillExactSet(...)`: repeated normalization of
    very short markdown table cells
  - add a local cache only for raw table-cell fragments whose byte length is `<= 8`
  - leave longer cell fragments uncached so `008055.xls` does not pay map overhead for the huge
    mostly-unique `9..32` byte cell population

- why:
  - fresh table-shape measurement showed that duplicate table cells are real, but heavily skewed
    toward short raw fragments
  - duplicate-cell counts:
    - `006087.xls`: `235659` total cells, `37952` unique, `197707` duplicates
    - `008055.xls`: `1313143` total cells, `1288825` unique, `24318` duplicates
    - `016161.xls`: `551263` total cells, `520599` unique, `30664` duplicates
  - raw-cell length split explained why the earlier broad cache idea was too expensive:
    - `006087.xls`:
      - `<=8`: `620670 dup / 34688 unique`
      - `9-16`: `40709 dup / 3099 unique`
      - `17-32`: `50632 dup / 102 unique`
    - `008055.xls`:
      - `<=8`: `79767 dup / 31039 unique`
      - `9-16`: `16120 dup / 617078 unique`
      - `17-32`: `1771 dup / 640619 unique`
    - `016161.xls`:
      - `<=8`: `95535 dup / 3196 unique`
      - `9-16`: `855 dup / 46582 unique`
      - `17-32`: `20124 dup / 470823 unique`
  - that shape strongly favored caching only the shortest raw cell fragments

- implementation:
  - introduce `markdownVisibleTableCellsWithCache(...)`
  - factor per-cell normalization into `markdownVisibleTableCellVariants(...)`
  - add a local `tableCellCache` inside `markdownBackfillExactSet(...)`
  - cache only when the raw cell fragment length is `<= 8`

- validation:
  - formatting:
    - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
  - focused hotspot benchmarks:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/006087.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
  - repeat-aware `.xls6` rerun:
    - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xls-short-cell-cache8-xls6.json -csv testdata/web-samples/reports/perf-exp-xls-short-cell-cache8-xls6.csv <xls6>`
  - repeat-aware mixed keyset rerun:
    - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-short-cell-cache8-keyset.json -csv testdata/web-samples/reports/perf-exp-short-cell-cache8-keyset.csv <current-keyset>`
  - repository regression:
    - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

- focused hotspot results:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `760822200 ns/op -> 645141700 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `2104182200 ns/op -> 1625965900 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `814307500 ns/op -> 638316500 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `173512500 ns/op -> 110346100 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 699207500 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `278034100 ns/op -> 312707000 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `495409100 ns/op -> 504470600 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1489791600 ns/op -> 1304217200 ns/op`

- repeat-aware `.xls6` results:
  - retained baseline total: `10461 ms`
  - experiment total: `10362 ms`
  - per-file:
    - `002505.xls`: `1387 ms -> 1373 ms`
    - `006087.xls`: `1351 ms -> 1277 ms`
    - `008055.xls`: `3922 ms -> 4064 ms`
    - `016161.xls`: `1665 ms -> 1711 ms`
    - `019088.xls`: `1077 ms -> 993 ms`
    - `019089.xls`: `1059 ms -> 944 ms`
  - output parity held on all six files (`textBytes/images` unchanged; no repeat mismatches)

- repeat-aware mixed keyset results:
  - `.docx`: `6931 ms -> 6478 ms`
  - `.xls`: `17830 ms -> 12063 ms`
  - `.xlsx`: `23756 ms -> 22413 ms`
  - output parity held on all 30 inputs

- repository regression:
  - `ok officeread 140.019s`

Interpretation:
- broad cell caching would have overpaid on `008055.xls`, where most medium-length raw cells are
  unique
- the short-cell-only cache is narrow enough to avoid that trap while still removing a large amount
  of repeated normalization work from the duplicate-heavy short-cell population
- some isolated exact-set / containment benchmarks remain mixed, but both repeat-aware integrated
  reruns stayed green and improved, which is the deciding signal for this code path

Decision:
- Retained.

## 2026-07-06 rejected: extend the short-cell cache into coverage / containment table-cell expansion

- experiment:
  - keep the retained short-cell cache inside `markdownBackfillExactSet(...)`
  - additionally reuse the same `markdownVisibleTableCellsWithCache(...)` helper inside:
    - `markdownBackfillBuildCoverageContainment(...)`
    - `addMarkdownBackfillCoverage(...)`
  - keep the cache scope narrow and identical to the retained exact-set version:
    - local map only
    - only cache raw cell fragments whose byte length is `<= 8`

- rationale:
  - after the retained exact-set short-cell cache, the next natural question was whether the same
    duplicate-heavy short-cell shape would also pay off on the coverage / containment side
  - the duplicate evidence already showed that short raw cells repeat heavily, especially on
    `006087.xls`, so reusing the same narrow cache looked plausible

- validation:
  - formatting:
    - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
  - focused hotspot benchmarks:
    - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/006087.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
  - repeat-aware `.xls6` rerun:
    - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-short-cell-cache8-coverage-xls6.json -csv testdata/web-samples/reports/perf-exp-short-cell-cache8-coverage-xls6.csv <xls6>`

- focused hotspot results:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `645141700 ns/op -> 574004600 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1625965900 ns/op -> 1661609000 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `638316500 ns/op -> 649785600 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `110346100 ns/op -> 140733800 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `699207500 ns/op -> 624811700 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `312707000 ns/op -> 287776800 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `504470600 ns/op -> 444791000 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1304217200 ns/op -> 1365105200 ns/op`

- repeat-aware `.xls6` results:
  - retained baseline total: `10362 ms`
  - experiment total: `10531 ms`
  - per-file:
    - `002505.xls`: `1373 ms -> 1342 ms`
    - `006087.xls`: `1277 ms -> 1186 ms`
    - `008055.xls`: `4064 ms -> 4344 ms`
    - `016161.xls`: `1711 ms -> 1633 ms`
    - `019088.xls`: `993 ms -> 928 ms`
    - `019089.xls`: `944 ms -> 1098 ms`
  - output parity held, but the integrated total still lost

Interpretation:
- the extra coverage-side cache helped `006087.xls` and some isolated coverage / exact-set benches
- but it gave back too much on `008055.xls` and slightly on `019089.xls`, which was enough to lose
  against the retained short-cell-cache baseline
- this is another case where a locally sensible reuse of a proven helper is still not a net win for
  the integrated `.xls` workload

Decision:
- Reverted.

## 2026-07-06 rejected: `.xls` exact precheck direct-hit fast path without duplicate direct lookup

- experiment:
  - revisit the earlier rejected `.xls` direct-first idea, but fix its main confounder
  - on the `.xls` exact precheck only:
    - first do the direct `exact[line]` lookup once
    - on miss, run a new helper that checks only the non-direct variants:
      - `markdownBackfillVisibleText(line)`
      - `markdownVisibleLineText(line)`
      - escaped table-cell visible text
  - unlike the earlier rejected direct-first experiment, this version does not pay the generic
    helper’s direct lookup a second time
- rationale:
  - temporary exact-hit-mode instrumentation confirmed that the retained heavy `.xls` source lines
    were still overwhelmingly direct hits:
    - `006087.xls`: `direct=40935`, `visible=0`, `markdown=0`, `escaped=0`, `miss=404`
    - `008055.xls`: `direct=1181169`, `visible=0`, `markdown=0`, `escaped=0`, `miss=0`
    - `016161.xls`: `direct=461603`, `visible=0`, `markdown=0`, `escaped=0`, `miss=0`
  - fresh exact-stage timing also showed that the line-scan part was large on the retained covered
    samples:
    - `008055.xls`: `exactBuild=634ms`, `exactScan=977ms`
    - `016161.xls`: `exactBuild=301ms`, `exactScan=322ms`
  - that made a non-duplicated direct-first split worth checking once, even after the earlier
    broader direct-first failure

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- temporary exact-hit and stage instrumentation:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSExactDirectHitShape$|^TestTmpXLSExactStageTiming$' -v ./`
- focused hotspot / extract benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSHotspots/006087.xls$|BenchmarkExtractXLSHotspots/008055.xls$|BenchmarkExtractXLSHotspots/016161.xls$|BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$' -benchmem -benchtime=1x ./`

Observed results:
- the exact-stage timing regressed instead of improving:
  - `002505.xls`: `exactBuild=271ms -> 366ms`, `exactScan=122ms -> 170ms`
  - `006087.xls`: `exactBuild=166ms -> 229ms`, `exactScan=18ms -> 19ms`
  - `008055.xls`: `exactBuild=634ms -> 780ms`, `exactScan=977ms -> 1320ms`
  - `016161.xls`: `exactBuild=301ms -> 378ms`, `exactScan=322ms -> 435ms`
- focused benchmarks also regressed sharply:
  - `BenchmarkExtractXLSHotspots/006087.xls`: `1348264600 ns/op -> 1352256600 ns/op`
  - `BenchmarkExtractXLSHotspots/008055.xls`: `4141427600 ns/op -> 4623366400 ns/op`
  - `BenchmarkExtractXLSHotspots/016161.xls`: `1766906500 ns/op -> 2015096000 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 906804600 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 2118399500 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 845987500 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 620281700 ns/op`

Interpretation:
- removing the duplicate direct lookup was not enough; the broader retained `.xls` exact path still
  got slower
- that effectively closes the loop on the direct-first family for this workload: the problem was not
  just “we paid direct twice”

Decision:
- Reverted.

## 2026-07-06 rejected: skip `seen` dedupe in `.xls` exact precheck

- experiment:
  - keep the retained `.xls` exact precheck structure and exact-set logic unchanged
  - only on the `.xls` path, stop maintaining the `seen` map inside
    `markdownBackfillExactLinesCoveredWithExact(...)`
  - rationale for trying this safely:
    - correctness does not depend on the dedupe map; removing it only means potentially rechecking
      duplicate source lines
- rationale:
  - fresh temporary instrumentation on the retained `.xls6` hotspot set showed that duplicate source
    lines in exact precheck were literally zero on every file:
    - `002505.xls`: `lines=215015`, `dups=0`
    - `006087.xls`: `lines=41347`, `dups=0`
    - `008055.xls`: `lines=1181198`, `dups=0`
    - `016161.xls`: `lines=461615`, `dups=0`
    - `019088.xls`: `lines=47549`, `dups=0`
    - `019089.xls`: `lines=51514`, `dups=0`
  - that made the `.xls`-only `seen` map look like pure overhead on the retained hotspot set

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- temporary source-shape instrumentation:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSExactSeenShape$' -v ./`
- focused hotspot / extract benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSHotspots/006087.xls$|BenchmarkExtractXLSHotspots/008055.xls$|BenchmarkExtractXLSHotspots/016161.xls$|BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$' -benchmem -benchtime=1x ./`
- repeat-aware `.xls6` rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xls-no-seen-xls6.json -csv testdata/web-samples/reports/perf-exp-xls-no-seen-xls6.csv <xls6>`
- repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused view was mixed:
  - `BenchmarkExtractXLSHotspots/006087.xls`: `1348264600 ns/op -> 1371871000 ns/op`
  - `BenchmarkExtractXLSHotspots/008055.xls`: `4141427600 ns/op -> 4279031400 ns/op`
  - `BenchmarkExtractXLSHotspots/016161.xls`: `1766906500 ns/op -> 2005530600 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 897860900 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 1487665800 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 605652400 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 856889900 ns/op`
- the decisive repeat-aware `.xls6` rerun still lost heavily against the retained baseline:
  - retained baseline total: `10461 ms`
  - experiment total: `15375 ms`
  - per-file:
    - `002505.xls`: `1387 ms -> 1517 ms`
    - `006087.xls`: `1351 ms -> 1869 ms`
    - `008055.xls`: `3922 ms -> 4738 ms`
    - `016161.xls`: `1665 ms -> 2123 ms`
    - `019088.xls`: `1077 ms -> 1573 ms`
    - `019089.xls`: `1059 ms -> 1555 ms`

Interpretation:
- even though the duplicate-count evidence was absolute, the `seen` path still mattered to the real
  retained `.xls` workload in a way that the raw “dups=0” metric did not predict
- this reinforces that the exact precheck cost is tied up with broader execution shape, not just
  obviously removable local bookkeeping

Decision:
- Reverted.

## 2026-07-06 rejected: plain-source fast return in `markdownBackfillCandidateLine(...)`

- experiment:
  - add a very narrow fast return to `markdownBackfillCandidateLine(...)`
  - after the initial `TrimSpace`, if `cleanTextFastPath(line)` succeeds and returns the original
    string unchanged, return it directly instead of continuing through the generic
    `cleanText(...)` + `stripInlineHiddenOfficeReferences(...)` path
- rationale:
  - new stage timing on the retained `.xls6` hotspots showed that `biffMarkdownWithText(...)` was
    dominated by its backfill stage, not by table assembly:
    - `006087.xls`: `tables=137ms`, `assembly=13ms`, `backfill=772ms`
    - `008055.xls`: `tables=768ms`, `assembly=32ms`, `backfill=1648ms`
    - `016161.xls`: `tables=344ms`, `assembly=14ms`, `backfill=775ms`
  - temporary source-line instrumentation then showed that the retained backfill source lines were
    almost entirely already-plain and unchanged by `markdownBackfillCandidateLine(...)`:
    - `006087.xls`: `lines=41347`, `plain=41336`, `changed=0`
    - `008055.xls`: `lines=1181198`, `plain=1181185`, `changed=0`
    - `016161.xls`: `lines=461615`, `plain=461613`, `changed=0`
  - that made `markdownBackfillCandidateLine(...)` look like a good place to cut generic cleaning
    work without changing coverage semantics

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- temporary stage instrumentation:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpBackfillSourceShape$|^TestTmpXLSMarkdownStageTiming$' -v ./`
- focused hotspot / extract benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$' -benchmem -benchtime=1x ./`

Observed results:
- the temporary stage timing after the change actually got worse inside the real markdown backfill:
  - `006087.xls`: `backfill=772ms -> 1086ms`
  - `008055.xls`: `backfill=1648ms -> 2067ms`
  - `016161.xls`: `backfill=775ms -> 836ms`
- focused benchmarks also regressed:
  - `BenchmarkExtractXLSHotspots/008055.xls`: `4141427600 ns/op -> 4232403500 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 2035422700 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 725233000 ns/op`

Interpretation:
- despite the very strong “plain source line” shape, bypassing the generic candidate-line cleanup
  path made the real retained backfill slower
- this suggests the generic path is doing work that materially helps later exact coverage checks,
  even when the final candidate string is unchanged

Decision:
- Reverted.

## 2026-07-06 rejected: BIFF plain-cell fast path in markdown cell cleanup / prepare

- experiment:
  - add a narrow fast path ahead of `cleanMarkdownTableCellValue(...)` for already-clean single-line
    BIFF table-cell values:
    - no `\r` / `\n`
    - `cleanTextFastPath(value)` succeeds and returns the original string unchanged
    - not flagged by `maybeDiscardableHiddenOfficeText(...)`
    - not flagged by `maybeControlFragmentText(...)`
  - add a matching fast path ahead of `prepareMarkdownTableCellValue(...)` for already-prepared
    plain values:
    - no `\r` / `\n` / `\` / `|`
    - already trimmed
    - not RTF-looking
- rationale:
  - the full `Extract(008055.xls)` profile showed that the end-to-end hotspot was not only
    exact/backfill; earlier BIFF markdown work remained large:
    - `biffMarkdownWithText`
    - `biffTextParts`
    - `biffWorksheetMarkdownTables`
    - `cleanMarkdownTableCellValue`
  - temporary BIFF cell-shape instrumentation showed that all three retained `.xls` hotspots were
    overwhelmingly plain at this stage:
    - `006087.xls`: `raw=235659`, `cleanChanged=0`, `rawShapes nl=0 slash=0 pipe=0 rtf=0`,
      `prepared br=0 escapedPipe=0 escapedBS=0`
    - `008055.xls`: `raw=1313143`, `cleanChanged=0`, `rawShapes nl=0 slash=0 pipe=0 rtf=0`,
      `prepared br=0 escapedPipe=0 escapedBS=0`
    - `016161.xls`: `raw=551263`, `cleanChanged=0`, `rawShapes nl=0 slash=0 pipe=0 rtf=0`,
      `prepared br=0 escapedPipe=0 escapedBS=0`
  - that made the BIFF markdown cell-clean / prepare pair look like a realistic end-to-end target

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- temporary BIFF cell-shape instrumentation:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpBIFFCellShape008055$' -v ./`
- focused hotspot / extract benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkExtractXLSHotspots/006087.xls$|BenchmarkExtractXLSHotspots/008055.xls$|BenchmarkExtractXLSHotspots/016161.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$' -benchmem -benchtime=1x ./`
- repeat-aware `.xls6` rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-biff-cell-fastpath-xls6.json -csv testdata/web-samples/reports/perf-exp-biff-cell-fastpath-xls6.csv <xls6>`
- repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused full-extract benchmarks improved on two of the three retained hotspots:
  - `BenchmarkExtractXLSHotspots/006087.xls`: `1348264600 ns/op -> 1295386300 ns/op`
  - `BenchmarkExtractXLSHotspots/008055.xls`: `4141427600 ns/op -> 3794928300 ns/op`
  - `BenchmarkExtractXLSHotspots/016161.xls`: `1766906500 ns/op -> 1798147600 ns/op`
- but the focused retained backfill-only view did not materially improve:
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: about `1769059600 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: about `645925500 ns/op`
- and the decisive repeat-aware `.xls6` rerun lost heavily against the retained baseline:
  - retained baseline total: `10461 ms`
  - experiment total: `14914 ms`
  - per-file:
    - `002505.xls`: `1387 ms -> 1428 ms`
    - `006087.xls`: `1351 ms -> 1683 ms`
    - `008055.xls`: `3922 ms -> 4862 ms`
    - `016161.xls`: `1665 ms -> 2205 ms`
    - `019088.xls`: `1077 ms -> 1351 ms`
    - `019089.xls`: `1059 ms -> 1385 ms`

Interpretation:
- this was the first rejected candidate in this stretch that genuinely improved the full
  `Extract(008055.xls)` benchmark, not just isolated exact/backfill helpers
- even so, the retained repeat-aware `.xls` rerun still got worse overall, which means the local
  BIFF plain-cell win did not translate into the right integrated balance

Decision:
- Reverted.

## 2026-07-06 rejected: helper early-return guards for markdown exact/backfill cleanup

- experiment:
  - keep the retained `.xls` exact/backfill structure unchanged
  - add only helper-local early returns to skip no-op work:
    - `markdownInlineLinkVisibleText(...)`: return immediately when the line contains no `[`
    - `stripMarkdownInlineWrappers(...)`: return immediately when the line contains no `*_~``
    - `unescapeMarkdownInlineFormattingMarkers(...)`: return immediately when the string contains no `\`
    - `unescapeMarkdownVisibleText(...)`: return immediately when the string contains no `\`
- rationale:
  - fresh temporary variant-shape instrumentation showed that exact precheck traffic on the retained
    `.xls` hotspots was almost entirely raw unique lines with essentially no extra variant growth:
    - `006087.xls`: `allowed=41339`, `visibleAdds=0`, `markdownAdds=0`, `escapedAdds=0`
    - `008055.xls`: `allowed=1181169`, `visibleAdds=0`, `markdownAdds=2`, `escapedAdds=0`
    - `016161.xls`: `allowed=461603`, `visibleAdds=0`, `markdownAdds=2`, `escapedAdds=0`
  - the current profile also still showed helper-string costs like `strings.Replace`,
    `strings.TrimSpace`, `strings.IndexAny`, and builder traffic inside the exact path
  - this made helper-local no-op guards look like the narrowest way to reduce string materialization
    without changing any higher-level `.xls` structure

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- temporary variant-shape instrumentation:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSExactVariantShape$' -v ./`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/006087.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
- repeat-aware `.xls6` rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-helper-early-return-xls6.json -csv testdata/web-samples/reports/perf-exp-helper-early-return-xls6.csv <xls6>`
- repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused hotspot benchmarks improved dramatically:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 408463000 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 1419166500 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 567112100 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `173512500 ns/op -> 83928700 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 377578500 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `278034100 ns/op -> 188349000 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `495409100 ns/op -> 326220200 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1489791600 ns/op -> 888473300 ns/op`
- allocation counts also dropped sharply on the focused benches, for example:
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `215699 allocs/op -> 63459 allocs/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `198340 allocs/op -> 46111 allocs/op`
- but the decisive repeat-aware `.xls6` rerun still regressed against the retained baseline:
  - retained baseline total: `10461 ms`
  - experiment total: `11871 ms`
  - per-file:
    - `002505.xls`: `1387 ms -> 1344 ms`
    - `006087.xls`: `1351 ms -> 1313 ms`
    - `008055.xls`: `3922 ms -> 4632 ms`
    - `016161.xls`: `1665 ms -> 2116 ms`
    - `019088.xls`: `1077 ms -> 1300 ms`
    - `019089.xls`: `1059 ms -> 1166 ms`

Interpretation:
- helper-local no-op guards produced one of the strongest isolated benchmark wins seen so far
- but even that was not enough to improve the retained integrated `.xls` rerun mix, with the losses
  concentrated on `008055.xls` and `016161.xls`
- this is a particularly strong reminder that isolated helper and alloc wins are still not the
  deciding metric for this workload

Decision:
- Reverted.

## 2026-07-06 rejected: shape-gated large exact-set capacity hint for very wide table markdown

- experiment:
  - add a shape-gated initial capacity hint for `markdownBackfillExactSet(...)`
  - estimate capacity from markdown structure only when the table is extremely wide:
    - `lineCount = strings.Count(markdown, "\n") + 1`
    - `pipeCount = strings.Count(markdown, "|")`
    - if `pipeCount >= lineCount * 40`, initialize the exact map with `lineCount + pipeCount`
- rationale:
  - fresh current-baseline profiles on `008055.xls` again showed that the exact-set path was still
    heavily map-bound:
    - `runtime.mapassign_faststr`
    - `internal/runtime/maps.(*table).split`
    - `internal/runtime/maps.(*table).rehash`
  - the older rejected preallocation attempt only used line count, which badly underestimates the
    true exact-set size for wide `.xls` markdown tables
  - temporary shape instrumentation showed:
    - `006087.xls`: `exact=86539`, `lines=54606`, `pipes=799984`, `lines+pipes=854590`
    - `008055.xls`: `exact=1305746`, `lines=16922`, `pipes=1403530`, `lines+pipes=1420452`
    - `016161.xls`: `exact=566236`, `lines=45647`, `pipes=682736`, `lines+pipes=728383`
  - that made `008055.xls` look like the only strong candidate for a large exact-map hint

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- temporary capacity-shape instrumentation:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLSExactCapacityShape$' -v ./`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/006087.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
- repeat-aware `.xls6` rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xls-exactset-capacity-xls6.json -csv testdata/web-samples/reports/perf-exp-xls-exactset-capacity-xls6.csv <xls6>`
- repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused hotspot view looked promising for the targeted wide-table sample:
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 2000567900 ns/op` (regressed whole-path)
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 547445700 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1489791600 ns/op -> 1331312700 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 748556100 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 820065400 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `173512500 ns/op -> 196222100 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `278034100 ns/op -> 297024100 ns/op`
- despite the exact-set-only win on `008055.xls`, the decisive repeat-aware `.xls6` rerun still
  lost badly against the retained baseline:
  - retained baseline total: `10461 ms`
  - experiment total: `12852 ms`
  - per-file:
    - `002505.xls`: `1387 ms -> 1448 ms`
    - `006087.xls`: `1351 ms -> 1657 ms`
    - `008055.xls`: `3922 ms -> 4832 ms`
    - `016161.xls`: `1665 ms -> 2186 ms`
    - `019088.xls`: `1077 ms -> 1304 ms`
    - `019089.xls`: `1059 ms -> 1425 ms`

Interpretation:
- even a shape-gated, width-aware exact-map capacity hint was not enough to improve the retained
  integrated `.xls` flow
- the map-heavy exact-set profile is real, but reducing exact-set rehash pressure in isolation still
  does not line up with the best end-to-end rerun behavior

Decision:
- Reverted.

## 2026-07-06 rejected: skip table-row `markdownVisibleLineText(...)` exact entries except empty-cell-gap shapes

- experiment:
  - keep the generic exact-set builder for non-table lines
  - on markdown lines recognized as table rows, stop adding the row-level
    `markdownVisibleLineText(...)` exact entry by default
  - keep a tiny exception only for rows containing the literal empty-cell-gap shape `|  |`, because
    temporary instrumentation showed that the only retained table-row visible exact additions on
    `006087.xls` came from rows where empty cells collapsed from `|  |` to `| |`
- rationale:
  - fresh exact-set-origin instrumentation showed that row-level visible exact entries were almost
    all duplicates of raw rows or cell entries:
    - `006087.xls`: `tableVisibleAdds=2`, `tableVisibleDup=49997`
    - `008055.xls`: `tableVisibleAdds=0`, `tableVisibleDup=16910`
    - `016161.xls`: `tableVisibleAdds=0`, `tableVisibleDup=45621`
  - on `008055.xls`, the overall exact-set origin split was:
    - `rawAdds=16917`
    - `visibleAdds=4` (all non-table)
    - `cellAdds=1288825`
  - that made the table-row visible exact insert look like nearly pure duplicate work

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- temporary origin instrumentation:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^TestTmpXLS008055ExactSetOrigins$' -v ./`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/006087.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownExactSetHotspots/016161.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`

Observed results:
- the origin evidence really was as narrow as it looked:
  - `006087.xls` sample table-row visible additions were only the two empty-cell-gap rows
  - `008055.xls` and `016161.xls` had zero unique table-row visible exact additions
- despite that, the integrated hotspot mix regressed badly:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 763505800 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 2171014300 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 623884600 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/006087.xls`: `173512500 ns/op -> 184974700 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 631996100 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/016161.xls`: `278034100 ns/op -> 473723000 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `495409100 ns/op -> 492549900 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1489791600 ns/op -> 1334273700 ns/op`

Interpretation:
- even an almost perfectly targeted attempt to remove duplicate table-row visible exact inserts did
  not translate into a usable integrated win
- this is another sign that the retained `.xls` exact stage is sensitive to small structural
  changes, and that duplicate-count evidence alone is not enough to justify altering it

Decision:
- Reverted.

## 2026-07-06 rejected: table-cell helper fast path for markdown exact-set / backfill cleanup

- experiment:
  - refactor the markdown table-cell cleanup path to avoid unconditional helper work on every cell
  - add trigger-based branches so cell cleanup only runs the expensive transforms when the current
    cell still contains the relevant marker family:
    - backslash unescape only when `\` is present
    - footnote stripping only when `[^` is present
    - wrapper / inline-format handling only when `*_~`` is present
    - HTML / autolink handling only when `<` or `&` is present
    - hard-break stripping only when the cleaned cell ends with `\`
  - reuse the same helper in `markdownVisibleTableCells(...)` so the `<br>` split segments also go
    through the narrowed cleanup path
- rationale:
  - current profiling still showed `markdownVisibleTableCells(...)`, `cleanVisibleText`, inline
    cleanup, and string helper churn inside the retained `.xls` exact-set path for `008055.xls`
  - the existing code was still running a full chain of trims and cleanup helpers for every table
    cell, even when the cell was already plain text

Validation:
- formatting:
  - `& 'C:\Program Files\Go\bin\gofmt.exe' -w extract.go`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSMarkdownBackfillHotspots/006087.xls$|BenchmarkXLSMarkdownBackfillHotspots/008055.xls$|BenchmarkXLSMarkdownBackfillHotspots/016161.xls$|BenchmarkXLSMarkdownExactSetHotspots/008055.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls$|BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls$' -benchmem -benchtime=1x ./`
- repeat-aware `.xls6` rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-xls-cell-fastpath-xls6.json -csv testdata/web-samples/reports/perf-exp-xls-cell-fastpath-xls6.csv <xls6>`
- repository regression:
  - `& 'C:\Program Files\Go\bin\go.exe' test ./...`

Observed results:
- focused hotspot view improved materially:
  - `BenchmarkXLSMarkdownBackfillHotspots/006087.xls`: `706136200 ns/op -> 510587000 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/008055.xls`: `1603810600 ns/op -> 1470653300 ns/op`
  - `BenchmarkXLSMarkdownBackfillHotspots/016161.xls`: `678344200 ns/op -> 650053900 ns/op`
  - `BenchmarkXLSMarkdownExactSetHotspots/008055.xls`: `633411500 ns/op -> 412054200 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/006087.xls`: `495409100 ns/op -> 389228900 ns/op`
  - `BenchmarkXLSMarkdownCoverageContainmentHotspots/008055.xls`: `1489791600 ns/op -> 1086864100 ns/op`
- full repository regression stayed green before revert:
  - `ok officeread 125.447s`
- but the decisive repeat-aware `.xls6` rerun regressed against the retained baseline:
  - retained baseline total: `10461 ms`
  - experiment total: `11748 ms`
  - per-file:
    - `002505.xls`: `1387 ms -> 1293 ms`
    - `006087.xls`: `1351 ms -> 1446 ms`
    - `008055.xls`: `3922 ms -> 4573 ms`
    - `016161.xls`: `1665 ms -> 2159 ms`
    - `019088.xls`: `1077 ms -> 1130 ms`
    - `019089.xls`: `1059 ms -> 1147 ms`

Interpretation:
- the helper-level fast path clearly helped the isolated benchmarked exact-set / containment work
- but it changed the integrated `.xls` rerun balance in the wrong direction, especially on
  `008055.xls` and `016161.xls`
- this looks like another case where a local helper win does not survive the real retained
  end-to-end workload

Decision:
- Reverted.
## 2026-07-06 rejected: simple-inline markdown row slice reuse in appendSimpleInlineWorksheetTextPrepared

- experiment:
  - reuse a per-function `rowScratch` slice for markdown row accumulation inside
    `appendSimpleInlineWorksheetTextPrepared(...)`
  - after each row flush, clear the current row in place and reset `rowValues` to `rowScratch[:0]`
    instead of dropping the slice to `nil`
  - allocate the scratch slice once on the first markdown-eligible row transition and keep reusing it
    across up to `maxMarkdownTableRows`
- rationale:
  - `testRecordSizeExceeded.xlsx` has `200000` rows and the markdown builder only retains the first
    `50000`, so repeated `rowValues` slice allocation looked like a plausible structural cost in the
    simple-inline path

Validation:
- targeted tests before decision:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run 'TestExtractXLSX|TestExtractWorksheet|TestPrepareMarkdownTableCellValueSingleLineFastPath|TestCleanMarkdownTableCellValue' ./`
- focused hotspot benchmarks:
  - `& 'C:\Program Files\Go\bin\go.exe' test -run '^$' -bench 'BenchmarkXLSXSimpleInlineHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkXLSXSimpleInlineTextOnlyHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$|BenchmarkExtractXLSXHotspots/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx$' -benchmem -benchtime=1x ./`
- repeat-aware pair rerun:
  - `& 'C:\Program Files\Go\bin\go.exe' run -buildvcs=false ./cmd/compatcheck -repeat 3 -json testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-rowscratch-pair-serial.json -csv testdata/web-samples/reports/perf-exp-ai-assistant-simpleinline-rowscratch-pair-serial.csv testdata/web-samples/samples/xlsx/apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx testdata/web-samples/samples/xlsx/00012389.xlsx`

Observed results:
- repeat-aware pair rerun was directionally positive but modest:
  - `00012389.xlsx`: `2056 ms`
  - `testRecordSizeExceeded.xlsx`: `6162 ms`
- but the focused integrated hotspot benchmark regressed hard versus the retained baseline:
  - `BenchmarkExtractXLSXHotspots/testRecordSizeExceeded.xlsx`: about `3130030700 ns/op -> 6098721400 ns/op`
  - `BenchmarkXLSXSimpleInlineHotspots/testRecordSizeExceeded.xlsx`: about `1931084000 ns/op -> 6063238700 ns/op`
  - `BenchmarkXLSXSimpleInlineTextOnlyHotspots/testRecordSizeExceeded.xlsx`: about `1423902100 ns/op -> 1734609500 ns/op`
- allocation profile also stayed very large and did not justify keeping the change:
  - `BenchmarkExtractXLSXHotspots/...`: `2249931048 B/op`, `12052922 allocs/op`
  - `BenchmarkXLSXSimpleInlineHotspots/...`: `1748064864 B/op`, `12050097 allocs/op`

Interpretation:
- the pair rerun alone was not decisive enough to override the direct hotspot regression
- reusing the row slice changed the simple-inline markdown path in a way that hurt the main end-to-end
  extraction flow, especially when markdown preparation and text extraction are measured together
- this is another case where a plausible allocation-reduction idea does not survive the real workload

Decision:
- Reverted.
