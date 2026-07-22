param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = 'Stop'

function Release-ComObject([object]$Object) {
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Normalize-OfficeText([string]$Text) {
    if ($null -eq $Text) { return '' }
    $normalized = $Text -replace "`r`n?", "`n"
    $normalized = $normalized -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    return ($normalized -replace '[ \t]+', ' ').Trim()
}

function Get-WordBaseline([string]$File) {
    $app = $null
    $document = $null
    try {
        $app = New-Object -ComObject Word.Application
        $app.Visible = $false
        $app.DisplayAlerts = 0
        # msoAutomationSecurityForceDisable: macros must not run during tests.
        $app.AutomationSecurity = 3
        $document = $app.Documents.Open($File, $false, $true, $false)
        $text = $document.Content.Text
        # msoPicture=13 and msoLinkedPicture=11. Shapes also includes text
        # boxes, callouts, and other non-image drawing objects.
        $images = 0
        foreach ($shape in $document.InlineShapes) {
            if ($shape.Type -eq 3 -or $shape.Type -eq 4) { $images++ }
        }
        foreach ($shape in $document.Shapes) {
            if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $images++ }
        }
        return @{ text = (Normalize-OfficeText $text); images = $images; source = 'Word.Content' }
    } finally {
        if ($null -ne $document) { $document.Close(0); Release-ComObject $document }
        if ($null -ne $app) { $app.Quit(); Release-ComObject $app }
    }
}

function Get-PowerPointBaseline([string]$File) {
    $app = $null
    $presentation = $null
    try {
        $app = New-Object -ComObject PowerPoint.Application
        $presentation = $app.Presentations.Open($File, $true, $false, $false)
        $parts = [System.Collections.Generic.List[string]]::new()
        $images = 0
        foreach ($slide in $presentation.Slides) {
            foreach ($shape in $slide.Shapes) {
                if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                    $parts.Add($shape.TextFrame.TextRange.Text)
                }
                # msoPicture=13 and msoLinkedPicture=11. OLE objects (7) are
                # not picture shapes and are deliberately excluded.
                if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $images++ }
            }
        }
        return @{ text = (Normalize-OfficeText ($parts -join "`n")); images = $images; source = 'PowerPoint.visible-slides' }
    } finally {
        if ($null -ne $presentation) { $presentation.Close(); Release-ComObject $presentation }
        if ($null -ne $app) { $app.Quit(); Release-ComObject $app }
    }
}

function Get-ExcelBaseline([string]$File) {
    $app = $null
    $workbook = $null
    try {
        $app = New-Object -ComObject Excel.Application
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $app.AutomationSecurity = 3
        $workbook = $app.Workbooks.Open($File, 0, $true)
        $parts = [System.Collections.Generic.List[string]]::new()
        $images = 0
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
                for ($row = 1; $row -le $rowCount; $row++) {
                    $cells = [System.Collections.Generic.List[string]]::new()
                    for ($column = 1; $column -le $columnCount; $column++) {
                        $cell = $used.Cells.Item($row, $column)
                        try { $cells.Add([string]$cell.Text) } finally { Release-ComObject $cell }
                    }
                    $parts.Add(($cells -join "`t"))
                }
                Release-ComObject $used
            }
            foreach ($shape in $sheet.Shapes) {
                if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $images++ }
            }
            Release-ComObject $sheet
        }
        return @{ text = (Normalize-OfficeText ($parts -join "`n")); images = $images; source = 'Excel.visible-worksheets.UsedRange.Text' }
    } finally {
        if ($null -ne $workbook) { $workbook.Close($false); Release-ComObject $workbook }
        if ($null -ne $app) { $app.Quit(); Release-ComObject $app }
    }
}

$extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
$result = [ordered]@{ path = $Path; ext = $extension; text = ''; images = 0; source = ''; error = '' }
try {
    switch ($extension) {
        '.doc' { $baseline = Get-WordBaseline $Path }
        '.docx' { $baseline = Get-WordBaseline $Path }
        '.ppt' { $baseline = Get-PowerPointBaseline $Path }
        '.pptx' { $baseline = Get-PowerPointBaseline $Path }
        '.xls' { $baseline = Get-ExcelBaseline $Path }
        '.xlsx' { $baseline = Get-ExcelBaseline $Path }
        default { throw "unsupported extension: $extension" }
    }
    $result.text = $baseline.text
    $result.images = [int]$baseline.images
    $result.source = $baseline.source
} catch {
    $result.error = $_.Exception.Message
}

[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$result | ConvertTo-Json -Compress -Depth 4
