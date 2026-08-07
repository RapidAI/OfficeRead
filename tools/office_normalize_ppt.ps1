param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

# This is deliberately a test/diagnostic helper.  It opens the source read-only
# with macros disabled, writes a distinct temporary PPTX, and never saves the
# input presentation.  The production Go extractor does not invoke it.
$ErrorActionPreference = 'Stop'
$app = $null
$presentation = $null

function Release-ComObject([object]$Object) {
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

try {
    $source = [IO.Path]::GetFullPath($InputPath)
    $destination = [IO.Path]::GetFullPath($OutputPath)
    if (-not [IO.Path]::GetExtension($source).Equals('.ppt', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'only legacy .ppt files may be normalized'
    }
    if ($source.Equals($destination, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'normalization output must differ from input'
    }
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
    if (Test-Path -LiteralPath $destination) { Remove-Item -LiteralPath $destination -Force }

    $app = New-Object -ComObject PowerPoint.Application
    try { $app.AutomationSecurity = 3 } catch { }
    # Presentations.Open(FileName, ReadOnly, Untitled, WithWindow).
    $presentation = $app.Presentations.Open($source, $true, $false, $false)
    # ppSaveAsOpenXMLPresentation = 24. SaveCopyAs avoids mutating the opened
    # document's identity or dirty state, even though it was opened read-only.
    $presentation.SaveCopyAs($destination, 24)
    if (-not (Test-Path -LiteralPath $destination) -or (Get-Item -LiteralPath $destination).Length -le 0) {
        throw 'PowerPoint did not create a non-empty PPTX copy'
    }
} finally {
    if ($null -ne $presentation) { $presentation.Close(); Release-ComObject $presentation }
    if ($null -ne $app) { $app.Quit(); Release-ComObject $app }
}
