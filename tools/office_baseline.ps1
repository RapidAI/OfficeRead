param(
    [Parameter(Mandatory = $true)]
    [string]$PathsBase64,
    [int]$ExcelMaxCells = 200000
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

function Get-WordBaseline([string]$File) {
    $document = $null
    try {
        if ($null -eq $script:wordApp) {
            $script:wordApp = New-Object -ComObject Word.Application
            $script:wordApp.Visible = $false
            $script:wordApp.DisplayAlerts = 0
            $script:wordApp.AutomationSecurity = 3
        }
        # msoAutomationSecurityForceDisable: macros must not run during tests.
        $document = $script:wordApp.Documents.Open($File, $false, $true, $false)
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
            if ($shape.Type -eq 3 -or $shape.Type -eq 4) { $images++; $inlineImages++; $inlineAnchors.Add([int]$shape.Range.Start) }
        }
        foreach ($shape in $document.Shapes) {
            if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $images++; $floatingImages++ }
        }
        $shapeAnchors = [System.Collections.Generic.List[int]]::new()
        foreach ($shape in $document.Shapes) {
            if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $shapeAnchors.Add([int]$shape.Anchor.Start) }
        }
        return @{ text = (Normalize-OfficeText $text); images = $images; inlineImages = $inlineImages; floatingImages = $floatingImages; inlineAnchors = @($inlineAnchors); shapeAnchors = @($shapeAnchors); source = 'Word.Content' }
    } finally {
        if ($null -ne $document) { $document.Close(0); Release-ComObject $document }
    }
}

function Get-PowerPointBaseline([string]$File) {
    $presentation = $null
    try {
        if ($null -eq $script:powerPointApp) { $script:powerPointApp = New-Object -ComObject PowerPoint.Application }
        $presentation = $script:powerPointApp.Presentations.Open($File, $true, $false, $false)
        $parts = [System.Collections.Generic.List[string]]::new()
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
                $parts.Add($Shape.TextFrame.TextRange.Text)
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
		return @{ text = (Normalize-OfficeText ($parts -join "`n")); images = $shapeCounts.images; groupImages = $shapeCounts.groupImages; imageFiles = @($imageFiles); source = 'PowerPoint.visible-slides.all-shapes' }
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
        $workbook = $script:excelApp.Workbooks.Open($File, 0, $true, 5, '', '', $true, 2, $null, $false, $false, $false, $false, $false, $false)
        $parts = [System.Collections.Generic.List[string]]::new()
        $images = 0
		$usedBulkValues = $false
        foreach ($sheet in $workbook.Worksheets) {
            if ($sheet.Visible -ne -1) { continue }
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
                if ($cellCount -gt $ExcelMaxCells) {
                    # Crossing the COM boundary once for every cell makes even
                    # a modest legacy workbook take minutes. Read Value2 as one
                    # SafeArray instead; it is still Excel's resolved stored
                    # worksheet value, but avoids an automation timeout.
                    $values = $used.Value2
                    $usedBulkValues = $true
                    if ($values -is [System.Array]) {
                        $rowLower = $values.GetLowerBound(0)
                        $columnLower = $values.GetLowerBound(1)
                        for ($row = 0; $row -lt $rowCount; $row++) {
                            $cells = [System.Collections.Generic.List[string]]::new()
                            for ($column = 0; $column -lt $columnCount; $column++) {
                                $cells.Add([string]$values.GetValue($rowLower + $row, $columnLower + $column))
                            }
                            $parts.Add(($cells -join "`t"))
                        }
                    } else {
                        $parts.Add([string]$values)
                    }
                } else {
                    for ($row = 1; $row -le $rowCount; $row++) {
                        $cells = [System.Collections.Generic.List[string]]::new()
                        for ($column = 1; $column -le $columnCount; $column++) {
                            $cell = $used.Cells.Item($row, $column)
                            try { $cells.Add([string]$cell.Text) } finally { Release-ComObject $cell }
                        }
                        $parts.Add(($cells -join "`t"))
                    }
                }
                Release-ComObject $used
            }
            foreach ($shape in $sheet.Shapes) {
                if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $images++ }
            }
            Release-ComObject $sheet
        }
        # For smaller ranges this is Excel's rendered cell Text.  Large ranges
        # deliberately use one Value2 SafeArray so the audit remains practical:
        # crossing the COM boundary once per cell can take several minutes.
        $source = 'Excel.visible-worksheets.UsedRange.Text'
        if ($usedBulkValues) { $source = 'Excel.visible-worksheets.UsedRange.Value2' }
        return @{ text = (Normalize-OfficeText ($parts -join "`n")); images = $images; source = $source }
    } finally {
        if ($null -ne $workbook) { $workbook.Close($false); Release-ComObject $workbook }
    }
}

try {
    $pathsJson = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($PathsBase64))
    foreach ($onePath in ($pathsJson | ConvertFrom-Json)) {
        $extension = [IO.Path]::GetExtension($onePath).ToLowerInvariant()
        $result = [ordered]@{ path = $onePath; ext = $extension; text = ''; images = 0; groupImages = 0; imageFiles = @(); inlineImages = 0; floatingImages = 0; inlineAnchors = @(); shapeAnchors = @(); source = ''; error = '' }
        try {
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
            $result.images = [int]$baseline.images
            if ($null -ne $baseline.groupImages) { $result.groupImages = [int]$baseline.groupImages }
            if ($null -ne $baseline.imageFiles) { $result.imageFiles = @($baseline.imageFiles) }
            if ($null -ne $baseline.inlineImages) { $result.inlineImages = [int]$baseline.inlineImages }
            if ($null -ne $baseline.floatingImages) { $result.floatingImages = [int]$baseline.floatingImages }
            if ($null -ne $baseline.inlineAnchors) { $result.inlineAnchors = @($baseline.inlineAnchors) }
            if ($null -ne $baseline.shapeAnchors) { $result.shapeAnchors = @($baseline.shapeAnchors) }
            $result.source = $baseline.source
        } catch {
            $result.error = $_.Exception.Message
        }
        $json = $result | ConvertTo-Json -Compress -Depth 4
        $utf8 = [Text.UTF8Encoding]::new($false).GetBytes($json + "`n")
        $stdout = [Console]::OpenStandardOutput()
        $stdout.Write($utf8, 0, $utf8.Length)
    }
} finally {
    if ($null -ne $excelApp) { $excelApp.Quit(); Release-ComObject $excelApp }
    if ($null -ne $powerPointApp) { $powerPointApp.Quit(); Release-ComObject $powerPointApp }
    if ($null -ne $wordApp) { $wordApp.Quit(); Release-ComObject $wordApp }
}
