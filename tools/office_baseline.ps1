param(
    [Parameter(Mandatory = $true)]
    [string]$PathsBase64,
    [int]$ExcelMaxCells = 10000,
    # A document that Office cannot open normally can show a modal
    # conversion/repair/File Block surface which never returns to COM.  The
    # recovery auditor may opt out after a normal open fails; regular strict
    # baselines retain Word's own repair and Protected View fallbacks.
    [switch]$NoRecoveryOpen,
    # Identifies this isolated invocation for timeout cleanup. It is only
    # process ownership metadata and does not affect extracted content.
    [string]$RunId = ''
)

$ErrorActionPreference = 'Stop'
$wordApp = $null
$powerPointApp = $null
$excelApp = $null

function Release-ComObject([object]$Object) {
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Normalize-OfficeText([string]$Text) {
    if ($null -eq $Text) { return '' }
    $normalized = $Text -replace "`r`n?", "`n"
    # Word's Content.Text uses several C0 control characters as field and
    # object boundaries.  Deleting them joins otherwise separate visible
    # words (for example "the" + "American" -> "theAmerican") and turns a
    # COM serialization detail into a false extractor mismatch.  They have no
    # visible glyph, but do form a text boundary, so normalize them to space.
    $normalized = $normalized -replace '\p{Cc}', ' '
    return ($normalized -replace '[ \t]+', ' ').Trim()
}

function Test-PasswordProtectedOfficePackage([string]$File) {
    # Agile/standard encrypted OOXML files retain their original extension,
    # but are Compound File Binary containers rather than ZIP packages.  Word
    # opens a password dialog for them; in a non-interactive COM run that dialog
    # cannot be answered and eventually consumes the watchdog timeout.  Detect
    # the documented EncryptionInfo + EncryptedPackage stream names before
    # activating Word.  This does not decrypt, modify, or otherwise inspect
    # user content beyond the container directory names.
    try {
        $bytes = [IO.File]::ReadAllBytes($File)
        if ($bytes.Length -lt 8) { return $false }
        $oleMagic = @(0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
        for ($i = 0; $i -lt $oleMagic.Count; $i++) {
            if ($bytes[$i] -ne $oleMagic[$i]) { return $false }
        }
        $directoryText = [Text.Encoding]::Unicode.GetString($bytes)
        return $directoryText.Contains('EncryptionInfo') -and $directoryText.Contains('EncryptedPackage')
    } catch {
        # A failed read is handled by the normal Office path and reported with
        # its native error; never convert an I/O problem into an encryption
        # claim.
        return $false
    }
}

function New-OfficeTextSegments([System.Collections.IEnumerable]$Parts) {
    $out = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($part in $Parts) {
        $text = [string]$part
        $context = ''
        if ($part -is [System.Collections.IDictionary]) {
            $text = [string]$part['text']
            $context = [string]$part['context']
        }
        $normalized = Normalize-OfficeText $text
        if ($normalized -eq '') { continue }
        $index++
        $out.Add(@{ index = $index; context = $context; text = $normalized })
    }
    return @($out)
}

function Initialize-OfficeBaselineExcelTextBridge {
    # Windows PowerShell's COM adapter performs a costly late-bound dispatch
    # for every property access.  On a formatted legacy XLS with thousands of
    # real cells that overhead alone can consume the per-document watchdog,
    # even though Excel itself has already finished rendering.  Use a tiny
    # in-process C# bridge for the sparse strict path.  It still obtains
    # *only* each Cell.Text value from the same Excel COM object; it merely
    # avoids routing every Row/Column/Text access through PowerShell.
    if ($null -eq ('OfficeBaselineExcelTextBridge' -as [type])) {
        # Use the CLR dynamic COM binder rather than Windows PowerShell's COM
        # adapter for the hot per-cell path.  Reflection's InvokeMember still
        # performs member lookup and argument marshaling for every single
        # populated cell; on a 400k-cell legacy XLS that alone exceeds a
        # multi-minute watchdog.  The dynamic binder caches the IDispatch
        # call-sites while preserving the one permitted acceptance source:
        # Cell.Text.  Microsoft.CSharp is part of the desktop CLR, but is not
        # referenced by Add-Type implicitly on all Windows PowerShell hosts.
        $runtimeDirectory = [Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()
        $systemCoreAssembly = Join-Path $runtimeDirectory 'System.Core.dll'
        $microsoftCSharpAssembly = Join-Path $runtimeDirectory 'Microsoft.CSharp.dll'
        Add-Type -ReferencedAssemblies @($systemCoreAssembly, $microsoftCSharpAssembly) -TypeDefinition @'
using System;
using System.Collections;
using System.Collections.Generic;

public static class OfficeBaselineExcelTextBridge {
    private static void AddText(List<string> texts, dynamic cell) {
        try {
            // Only the displayed Excel property is an acceptance source.
            object value = cell.Text;
            string text = value == null ? "" : Convert.ToString(value);
            if (text.Length != 0) texts.Add(text);
        } finally {
            Release(cell);
        }
    }


    private static void ReadTyped(dynamic populated, List<string> texts) {
        dynamic cells = populated.Cells;
        try {
            if (cells is Array) {
                foreach (object cell in (Array)cells) AddText(texts, cell);
            } else {
                foreach (object cell in (IEnumerable)cells) AddText(texts, cell);
            }
        } finally { Release(cells); }
    }



    private static void ReadWholeRange(dynamic usedRange, List<string> texts) {
        // Excel rejects Range.SpecialCells on protected worksheets, even when
        // the workbook is opened read-only.  Treating that COM policy error as
        // an empty sheet silently removes visible text from the Office
        // reference.  The fallback deliberately reads the same Cell.Text
        // property cell-by-cell; it is slower, but remains an Office-visible
        // baseline and is normally limited to the protected report/template
        // sheets for which SpecialCells is unavailable.
        dynamic cells = usedRange.Cells;
        try {
            if (cells is Array) {
                foreach (object cell in (Array)cells) AddText(texts, cell);
            } else {
                foreach (object cell in (IEnumerable)cells) AddText(texts, cell);
            }
        } finally { Release(cells); }
    }

    public static string[] Read(dynamic usedRange) {
        var texts = new List<string>();
        // xlCellTypeConstants (2) and xlCellTypeFormulas (-4123) overlap for
        // neither cells nor formula results.  They are deliberately read in
        // their worksheet order as two efficient COM ranges.  Do *not* use a
        // successful Constants call as proof that Formulas has nothing to
        // contribute: a formula-only visible sheet is a common report layout
        // and the old code silently produced its sheet name only.
        bool foundRenderedCells = false;
        foreach (int cellType in new[] { 2, -4123 }) {
            try {
                dynamic populated = usedRange.SpecialCells(cellType);
                try {
                    foundRenderedCells = true;
                    // A COM enumerable can invoke IDispatch once per cell merely
                    // to obtain the next object. Ask Excel for its Cells SafeArray
                    // in one call, then obtain only Cell.Text from each element.
                    // This remains the strict rendered-text source, while avoiding
                    // a large portion of the marshaling cost on dense legacy XLS.
                    ReadTyped(populated, texts);
                } finally {
                    Release(populated);
                }
            } catch {
                // Excel raises when a requested SpecialCells class is absent.
                // If both classes are unavailable (not merely empty), this is
                // commonly a protected worksheet; scan its rendered cells.
            }
        }
        if (!foundRenderedCells) ReadWholeRange(usedRange, texts);
        return texts.ToArray();
    }

    private static void Release(object value) {
        if (value == null || !System.Runtime.InteropServices.Marshal.IsComObject(value)) return;
        try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value); } catch { }
    }

    public static string[] ReadAll(object usedRange) {
        // Empty formatted cells have no rendered Text. Reuse the same strict
        // Text source for all range sizes, avoiding a dense sweep of blank
        // formatting-only cells that causes unattended Excel timeouts.
        return Read(usedRange);
    }
}
'@
    }
}

function Get-ExcelSparseRenderedTexts([object]$UsedRange) {
    Initialize-OfficeBaselineExcelTextBridge
    return @([OfficeBaselineExcelTextBridge]::Read($UsedRange))
}

function Get-ExcelRenderedTexts([object]$Range) {
    # Read every cell's Excel-rendered Text, including cells which do not
    # appear in SpecialCells constants/formulas. Keep the loop in the CLR so
    # the strict source does not pay Windows PowerShell's per-property COM
    # adapter overhead on medium dense worksheets.
    Initialize-OfficeBaselineExcelTextBridge
    return @([OfficeBaselineExcelTextBridge]::ReadAll($Range))
}

function Get-WordBaseline([string]$File) {
    $document = $null
    $protectedView = $null
    try {
        if ($null -eq $script:wordApp) {
            $script:wordApp = New-Object -ComObject Word.Application
            $script:wordApp.Visible = $false
            $script:wordApp.DisplayAlerts = 0
            $script:wordApp.AutomationSecurity = 3
        }
        # msoAutomationSecurityForceDisable: macros must not run during tests.
        # Word evaluates DATE fields while opening the document.  Preserve that
        # instant for the strict OOXML extractor; recording a later time after
        # Content.Text returns can be several seconds behind the field value.
        $fieldTime = [DateTime]::Now
        # Use Word's ordinary read-only open as the reference path.  A small
        # number of corpus documents are accepted only after Word's own
        # OpenAndRepair recovery pass; retry that mode only after the ordinary
        # Office open has failed.  This is still a Word-rendered baseline, not
        # an extractor-side repair, and the source records the recovery so it
        # remains auditable in the comparison report.
        $recoveredOpen = $false
        $protectedViewOpen = $false
        try {
            $document = $script:wordApp.Documents.Open($File, $false, $true, $false)
        } catch {
            if ($NoRecoveryOpen) { throw }
            # Documents.Open arguments through OpenAndRepair (14th argument):
            # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
            # PasswordDocument, PasswordTemplate, Revert,
            # WritePasswordDocument, WritePasswordTemplate, Format,
            # Encoding, Visible, OpenConflictDocument, OpenAndRepair.
            try {
                $document = $script:wordApp.Documents.Open($File, $false, $true, $false, '', '', $false, '', '', 0, '', $false, $false, $true, $false)
                $recoveredOpen = $true
            } catch {
                # A Trust Center File Block can deny Documents.Open before Word
                # has parsed any document content.  Protected View is Word's
                # own read-only rendering path for that policy category.  Use
                # it solely as a last-resort reference surface; never call
                # Edit(), save, or change the document/security policy.
                $protectedView = $script:wordApp.ProtectedViewWindows.Open($File, $false, '', $false)
                $document = $protectedView.Document
                if ($null -eq $document) { throw 'Word Protected View did not expose a document reference' }
                $protectedViewOpen = $true
            }
        }
        # Document.Content.Text follows Word's current revision-display mode.
        # An automation-created window can default to "Simple Markup", which
        # exposes deleted text that is absent from the final document.  The
        # extractor intentionally follows final content, so make the Office
        # reference explicit as well (wdRevisionsViewFinal = 0).  Word's
        # original/final enum values are counter-intuitive: 1 is Original,
        # and would make deleted text appear in Document.Content.Text.
        try {
            $document.ActiveWindow.View.RevisionsView = 0
            $document.ActiveWindow.View.ShowRevisionsAndComments = $false
        } catch { }
        $text = $document.Content.Text
        # msoPicture=13 and msoLinkedPicture=11. Shapes also includes text
        # boxes, callouts, and other non-image drawing objects.
        $images = 0
        $groupImages = 0
        $inlineImages = 0
        $floatingImages = 0
        $inlineAnchors = [System.Collections.Generic.List[int]]::new()
        foreach ($shape in $document.InlineShapes) {
            # wdInlineShapePicture=3 and wdInlineShapeLinkedPicture=4.  Word
            # also exposes OLE, chart, SVG, and other inline drawing objects
            # through this collection; they are not Picture Shapes and must
            # not be folded into the image baseline.
            if ($shape.Type -eq 3 -or $shape.Type -eq 4) { $images++; $inlineImages++; $inlineAnchors.Add([int]$shape.Range.Start) }
        }
        # Word returns a top-level msoGroup (6) for grouped floating artwork.
        # Its Picture children are real, visible GroupItems, but are absent
        # from Document.Shapes itself.  Recurse exactly as Word exposes the
        # group so a grouped picture is not silently treated as zero images.
        function Add-WordShapeImages([object]$Shape, [int]$GroupDepth) {
            if ($Shape.Type -eq 6) {
                try {
                    foreach ($child in $Shape.GroupItems) {
                        Add-WordShapeImages $child ($GroupDepth + 1)
                    }
                } catch { }
                return
            }
            if ($Shape.Type -eq 13 -or $Shape.Type -eq 11) {
                $script:wordFloatingImages++
                $script:wordImages++
                if ($GroupDepth -gt 0) { $script:wordGroupImages++ }
                try { $script:wordShapeAnchors.Add([int]$Shape.Anchor.Start) } catch { }
            }
        }
        $script:wordImages = $images
        $script:wordFloatingImages = $floatingImages
        $script:wordGroupImages = $groupImages
        $script:wordShapeAnchors = [System.Collections.Generic.List[int]]::new()
        foreach ($shape in $document.Shapes) {
            Add-WordShapeImages $shape 0
        }
        $images = $script:wordImages
        $floatingImages = $script:wordFloatingImages
        $groupImages = $script:wordGroupImages
        $shapeAnchors = [System.Collections.Generic.List[int]]::new()
		$shapeAnchors.AddRange($script:wordShapeAnchors)
		$segments = [System.Collections.Generic.List[object]]::new()
		# Keep Content.Text as the authoritative aggregate baseline, but expose
		# paragraph and field-code segments for diagnosis. Field codes such as
		# DATE/PAGE have no standalone visible glyphs, yet explain otherwise
		# opaque Content.Text token differences without putting headers/footers
		# outside Word.Document.Content into the comparison scope.
		$paragraphIndex = 0
		try {
			foreach ($paragraph in $document.Content.Paragraphs) {
				$paragraphIndex++
				$segments.Add(@{ context = ('document-paragraph-{0}' -f $paragraphIndex); text = [string]$paragraph.Range.Text })
			}
		} catch { }
		if ($segments.Count -eq 0) { $segments.Add(@{ context = 'document-content'; text = [string]$text }) }
        $source = 'Word.Content'
        if ($recoveredOpen) { $source += '.recovered' }
        if ($protectedViewOpen) { $source = 'Word.ProtectedView.Content' }
        return @{ text = (Normalize-OfficeText $text); textSegments = (New-OfficeTextSegments $segments); images = $images; inlineImages = $inlineImages; floatingImages = $floatingImages; inlineAnchors = @($inlineAnchors); shapeAnchors = @($shapeAnchors); source = $source; fieldTime = $fieldTime.ToString('o') }
    } finally {
        # ProtectedViewWindow owns its Document. Closing both objects can make
        # the second COM call fail (or display a policy prompt), so close the
        # owning surface exactly once and release the document reference.
        if ($null -ne $protectedView) {
            try { $protectedView.Close() } catch { } finally { Release-ComObject $protectedView }
            Release-ComObject $document
        } elseif ($null -ne $document) {
            try { $document.Close(0) } catch { } finally { Release-ComObject $document }
        }
    }
}

function Get-PowerPointBaseline([string]$File) {
    $presentation = $null
    try {
        if ($null -eq $script:powerPointApp) { $script:powerPointApp = New-Object -ComObject PowerPoint.Application }
        $presentation = $script:powerPointApp.Presentations.Open($File, $true, $false, $false)
        $parts = [System.Collections.Generic.List[string]]::new()
		$textSegments = [System.Collections.Generic.List[object]]::new()
        $shapeCounts = @{ images = 0; groupImages = 0 }
        $imageFiles = [System.Collections.Generic.List[string]]::new()
        function Add-PowerPointShapeBaseline([object]$Shape, [string]$ShapePath, [int]$GroupDepth) {
            if ($Shape.Type -eq 6) {
                try {
                    foreach ($child in $Shape.GroupItems) {
                        Add-PowerPointShapeBaseline $child ($ShapePath + '-' + $child.Id) ($GroupDepth + 1)
                    }
                } catch { }
                return
            }
            if ($Shape.HasTextFrame -eq -1 -and $Shape.TextFrame.HasText -eq -1) {
				$shapeText = $Shape.TextFrame.TextRange.Text
                $parts.Add($shapeText)
				$textSegments.Add(@{ context = ('slide-{0}-shape-{1}' -f $slide.SlideIndex, $ShapePath); text = $shapeText })
            }
            # msoPicture=13 and msoLinkedPicture=11. OLE objects (7) are not
            # picture shapes and are deliberately excluded.  Group members are
            # separately exposed by GroupItems and are rendered on the slide.
            if ($Shape.Type -eq 13 -or $Shape.Type -eq 11) {
                $shapeCounts.images++
                if ($GroupDepth -gt 0) { $shapeCounts.groupImages++ }
                # Shape.Export captures the actual visual picture payload
                # PowerPoint exposes (including producer-specific image
                # encoding).  The Go comparator decodes it before hashing,
                # so PNG container metadata/recompression does not affect
                # quality matching.  Export is best-effort: unsupported
                # linked/metafile shapes still participate in count parity.
                try {
                    $directory = Join-Path ([IO.Path]::GetTempPath()) ('officebaseline-' + [guid]::NewGuid().ToString('N'))
                    [void][IO.Directory]::CreateDirectory($directory)
                    $destination = Join-Path $directory ('slide-{0:D4}-shape-{1}.png' -f $slide.SlideIndex, $ShapePath)
                    $Shape.Export($destination, 2)
                    if (Test-Path -LiteralPath $destination) { $imageFiles.Add($destination) }
                } catch { }
            }
        }
        foreach ($slide in $presentation.Slides) {
            # PowerPoint represents a hidden slide as msoTrue (-1).  The OOXML
            # extractor deliberately excludes p:sld@show="0", so use the same
            # slideshow-visible scope for the Office reference result.
            if ($slide.SlideShowTransition.Hidden -eq -1) { continue }
            foreach ($shape in $slide.Shapes) {
                Add-PowerPointShapeBaseline $shape ([string]$shape.Id) 0
            }
        }
		return @{ text = (Normalize-OfficeText ($parts -join "`n")); textSegments = (New-OfficeTextSegments $textSegments); images = $shapeCounts.images; groupImages = $shapeCounts.groupImages; imageFiles = @($imageFiles); source = 'PowerPoint.visible-slides.all-shapes' }
    } finally {
        if ($null -ne $presentation) { $presentation.Close(); Release-ComObject $presentation }
    }
}

function Get-ExcelBaseline([string]$File) {
    $workbook = $null
    try {
        if ($null -eq $script:excelApp) {
            $script:excelApp = New-Object -ComObject Excel.Application
            $script:excelApp.Visible = $false
            $script:excelApp.DisplayAlerts = $false
            $script:excelApp.AutomationSecurity = 3
            $script:excelApp.AskToUpdateLinks = $false
            $script:excelApp.EnableEvents = $false
            try { $script:excelApp.Calculation = -4135 } catch { }
        }
        # Prevent external links, file validation, and calculation prompts from
        # blocking unattended compatibility runs on older workbooks.
        # Corrupt-load=2 (xlExtractData) is a recovery mode, not a normal
        # comparison mode. It can surface worksheets that Excel would not open
        # unchanged and produces an unnecessarily expensive repair path. Open
        # normally first; if Excel rejects a damaged workbook, retry explicitly
        # with extraction mode and retain that fact in the baseline source.
        $recoveredOpen = $false
        try {
            # Do not pass placeholder $null arguments through the COM binder.
            # In Windows PowerShell that overload resolution can fail before
            # Excel even sees the filename, causing every healthy workbook to
            # be opened by the recovery fallback.  The first three arguments
            # are enough to request a normal, read-only, no-link-update open;
            # application-level AutomationSecurity/DisplayAlerts/EnableEvents
            # above provide the remaining unattended safety controls.
            $workbook = $script:excelApp.Workbooks.Open($File, 0, $true)
        } catch {
            if ($NoRecoveryOpen) { throw }
            $workbook = $script:excelApp.Workbooks.Open($File, 0, $true, 5, '', '', $true, 2, $null, $false, $false, $false, $false, $false, $false)
            $recoveredOpen = $true
        }
        $parts = [System.Collections.Generic.List[string]]::new()
        $images = 0
        $usedBulkValues = $false
        # The threshold applies to the entire workbook, not one worksheet.
        # A workbook with twelve individually modest 2,640-cell sheets still
        # requires more than thirty thousand cross-process .Text calls. That
        # is precisely the pattern that caused an unattended Excel COM child
        # to exceed its timeout. First inspect the cheap range dimensions,
        # then choose one consistent source scope for every visible sheet.
        [int64]$visibleUsedCells = 0
        # Do not retain worksheet RCWs across the measurement and rendered
        # passes. On a few Excel builds a later ReleaseComObject can detach an
        # RCW that is still held in a PowerShell generic collection, yielding
        # "RCW separated" on otherwise healthy workbooks. Keep only the
        # stable 1-based worksheet indexes; Item() returns a fresh wrapper for
        # each pass and every wrapper is released in that same pass.
        $visibleSheetIndexes = [System.Collections.Generic.List[int]]::new()
        $worksheets = $workbook.Worksheets
        try {
            for ($sheetIndex = 1; $sheetIndex -le [int]$worksheets.Count; $sheetIndex++) {
                $sheet = $worksheets.Item($sheetIndex)
                # Excel returns XlSheetVisibility as xlSheetVisible=-1,
                # xlSheetHidden=0, xlSheetVeryHidden=2. The COM adapter can
                # also marshal VARIANT_BOOL as System.Boolean. Preserve that
                # distinction: converting a numeric -1 through [bool] changes
                # it to 1 and then loses the actual Excel enum value.
                $rawSheetVisibility = $sheet.Visible
                $sheetVisible = if ($rawSheetVisibility -is [bool]) { [bool]$rawSheetVisibility } else { [int]$rawSheetVisibility -eq -1 }
                if (-not $sheetVisible) { Release-ComObject $sheet; continue }
                $visibleSheetIndexes.Add($sheetIndex)
                $used = $sheet.UsedRange
                if ($null -ne $used) {
                    try {
                        $visibleUsedCells += [int64]$used.Rows.Count * [int64]$used.Columns.Count
                    } finally { Release-ComObject $used }
                }
                Release-ComObject $sheet
            }
        } finally {
            Release-ComObject $worksheets
        }
        # Retain the explicit comparison for the per-sheet variable as well:
        # it documents that either a large individual range or the aggregate
        # workbook total selects the bounded Value2 baseline.  Do not compute
        # an 80% safety margin in strict mode: Int32.MaxValue is the explicit
        # no-Value2 contract and multiplying it in PowerShell produces a
        # floating-point value that is unnecessary for every workbook.
        $renderedTextBudget = [int64]$ExcelMaxCells
        if ($ExcelMaxCells -lt 2147483647) {
            # The cost of Cell.Text is dominated by COM crossings, not just
            # cells. Reserve a margin only for the non-strict coverage mode,
            # where a stored-value result is deliberately marked excluded.
            $renderedTextBudget = [Math]::Max(1, [int64]([Math]::Floor($ExcelMaxCells * 0.8)))
            if ($visibleUsedCells -gt $renderedTextBudget) { $usedBulkValues = $true }
        }
        $worksheets = $workbook.Worksheets
        try {
            foreach ($sheetIndex in $visibleSheetIndexes) {
                $sheet = $worksheets.Item($sheetIndex)
                try {
                $parts.Add("# " + $sheet.Name)
                $used = $sheet.UsedRange
                if ($null -ne $used) {
                # Value2 is the stored value, while Office renders the Text
                # property according to the cell's number format. Read cells
                # by bounds instead of enumerating a COM 2-D SafeArray: the
                # latter is flattened by Windows PowerShell and loses rows.
                $rowCount = [int]$used.Rows.Count
                $columnCount = [int]$used.Columns.Count
                $cellCount = [int64]$rowCount * [int64]$columnCount
                if ($cellCount -gt $ExcelMaxCells -or $cellCount -gt $renderedTextBudget -or $usedBulkValues) {
                    # Crossing the COM boundary once for every cell makes even
                    # a modest legacy workbook take minutes. Read Value2 as one
                    # SafeArray instead; it is still Excel's resolved stored
                    # worksheet value, but avoids an automation timeout.  Do
                    # not reconstruct that SafeArray with GetValue(row,col):
                    # a formatted UsedRange can extend to XFC even when only a
                    # handful of cells contain values, and millions of managed
                    # GetValue calls can itself exceed the COM timeout.  Row
                    # boundaries are immaterial for this explicitly stored-
                    # value comparison scope; retain non-empty values in the
                    # SafeArray's natural (row-major) enumeration order.
                    # Strict runs set ExcelMaxCells to Int32.MaxValue, so this
                    # branch is retained only for the explicitly non-strict
                    # coverage mode.  The strict path below never obtains
                    # stored Value2 values as an acceptance baseline.
                    $values = $used.Value2
                    $usedBulkValues = $true
                    if ($values -is [System.Array]) {
                        foreach ($value in $values) {
                            if ($null -ne $value -and [string]$value -ne '') {
                                $parts.Add([string]$value)
                            }
                        }
                    } else {
                        $parts.Add([string]$values)
                    }
                } else {
                    # Some legacy workbooks carry a UsedRange spanning an
                    # entire formatted sheet while only a small subset has a
                    # value or formula.  A blank formatted cell contributes
                    # no visible Text, so do not spend one COM round trip on
                    # each of those empty positions.  This sparse path still
                    # reads Excel's rendered .Text property for every actual
                    # constant/formula cell; it never substitutes Value2.
                    # A 250k-cell gate still routes a 100k-cell formatted
                    # range through 100k PowerShell COM dispatches.  That is
                    # enough to exceed the watchdog on cold legacy Excel,
                    # despite most of those cells being blank formatting.
                    # Above 10k cells use the same strict sparse Text path;
                    # SpecialCells identifies only constants/formulas and the
                    # bridge reads their rendered .Text values in coordinate
                    # order.  This preserves the visible-content source while
                    # making the timeout bound depend on populated cells.
                    $sparseStrictText = $ExcelMaxCells -ge 2147483647 -and $cellCount -gt 10000
                    if ($sparseStrictText) {
                        # The comparison normalizes whitespace into token
                        # boundaries, so a row/tab reconstruction provides no
                        # visible-content information beyond the rendered cell
                        # texts themselves. Avoid two extra COM calls (Row and
                        # Column) plus PowerShell sorting for every populated
                        # cell. The bridge still gets Cell.Text for each cell;
                        # it merely returns that strict source directly.
                        foreach ($text in (Get-ExcelSparseRenderedTexts $used)) {
                            $parts.Add([string]$text)
                        }
                    } else {
                        # Keep the same Cell.Text source below the sparse
                        # threshold. The in-process bridge avoids one
                        # PowerShell COM-adapter dispatch per property while
                        # preserving every populated rendered value.
                        foreach ($text in (Get-ExcelRenderedTexts $used)) {
                            $parts.Add([string]$text)
                        }
                    }
                }
                    Release-ComObject $used
                }
                # Shapes is a COM collection too.  Enumerating it leaves one
                # RCW per shape alive until Excel.Quit, which can turn a
                # completed strict Text read into a delayed COM timeout.
                $shapes = $sheet.Shapes
                try {
                    for ($shapeIndex = 1; $shapeIndex -le [int]$shapes.Count; $shapeIndex++) {
                        $shape = $shapes.Item($shapeIndex)
                        try {
                            if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $images++ }
                        } finally { Release-ComObject $shape }
                    }
                } finally { Release-ComObject $shapes }
                } finally {
                    Release-ComObject $sheet
                }
            }
        } finally {
            Release-ComObject $worksheets
        }
        # For smaller ranges this is Excel's rendered cell Text.  Large ranges
        # deliberately use one Value2 SafeArray so the audit remains practical:
        # crossing the COM boundary once per cell can take several minutes.
        $source = 'Excel.visible-worksheets.UsedRange.Text'
        if ($usedBulkValues) { $source = 'Excel.visible-worksheets.UsedRange.Value2' }
		if ($recoveredOpen) { $source += '.recovered' }
		return @{ text = (Normalize-OfficeText ($parts -join "`n")); textSegments = (New-OfficeTextSegments $parts); images = $images; source = $source }
    } finally {
        if ($null -ne $workbook) { $workbook.Close($false); Release-ComObject $workbook }
    }
}

try {
    $pathsJson = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($PathsBase64))
    foreach ($onePath in ($pathsJson | ConvertFrom-Json)) {
        $extension = [IO.Path]::GetExtension($onePath).ToLowerInvariant()
        $result = [ordered]@{ path = $onePath; ext = $extension; text = ''; textSegments = @(); images = 0; groupImages = 0; imageFiles = @(); inlineImages = 0; floatingImages = 0; inlineAnchors = @(); shapeAnchors = @(); source = ''; error = '' }
        try {
            if (Test-PasswordProtectedOfficePackage $onePath) {
                throw 'password-protected Office package; no plaintext Office COM baseline is available without a password (password prompt avoided)'
            }
            switch ($extension) {
                '.doc' { $baseline = Get-WordBaseline $onePath }
                '.docx' { $baseline = Get-WordBaseline $onePath }
                '.ppt' { $baseline = Get-PowerPointBaseline $onePath }
                '.pptx' { $baseline = Get-PowerPointBaseline $onePath }
                '.xls' { $baseline = Get-ExcelBaseline $onePath }
                '.xlsx' { $baseline = Get-ExcelBaseline $onePath }
                default { throw "unsupported extension: $extension" }
            }
            $result.text = $baseline.text
			if ($null -ne $baseline.textSegments) { $result.textSegments = @($baseline.textSegments) }
            $result.images = [int]$baseline.images
            if ($null -ne $baseline.groupImages) { $result.groupImages = [int]$baseline.groupImages }
            if ($null -ne $baseline.imageFiles) { $result.imageFiles = @($baseline.imageFiles) }
            if ($null -ne $baseline.inlineImages) { $result.inlineImages = [int]$baseline.inlineImages }
            if ($null -ne $baseline.floatingImages) { $result.floatingImages = [int]$baseline.floatingImages }
            if ($null -ne $baseline.inlineAnchors) { $result.inlineAnchors = @($baseline.inlineAnchors) }
            if ($null -ne $baseline.shapeAnchors) { $result.shapeAnchors = @($baseline.shapeAnchors) }
            $result.source = $baseline.source
            if ($null -ne $baseline.fieldTime) { $result.fieldTime = [string]$baseline.fieldTime }
        } catch {
            $result.error = $_.Exception.Message
        }
        $json = $result | ConvertTo-Json -Compress -Depth 4
        $utf8 = [Text.UTF8Encoding]::new($false).GetBytes($json + "`n")
        $stdout = [Console]::OpenStandardOutput()
        $stdout.Write($utf8, 0, $utf8.Length)
    }
} finally {
    # A timed-out sibling audit can already have terminated an Automation
    # server. Quit then raises RPC_S_SERVER_UNAVAILABLE in this finally block,
    # which used to discard the valid JSON result emitted above. Cleanup must
    # be best-effort; per-file failures are captured in the result stream.
    if ($null -ne $excelApp) { try { $excelApp.Quit() } catch { } finally { Release-ComObject $excelApp } }
    if ($null -ne $powerPointApp) { try { $powerPointApp.Quit() } catch { } finally { Release-ComObject $powerPointApp } }
    if ($null -ne $wordApp) { try { $wordApp.Quit() } catch { } finally { Release-ComObject $wordApp } }
}
