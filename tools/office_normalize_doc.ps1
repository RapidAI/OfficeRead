param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

# Test/diagnostic helper only: open a legacy DOC read-only with macros disabled
# and create a distinct DOCX copy. The production Go extractor never invokes it.
$ErrorActionPreference = 'Stop'
$app = $null
$document = $null

function Release-ComObject([object]$Object) {
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

try {
    $source = [IO.Path]::GetFullPath($InputPath)
    $destination = [IO.Path]::GetFullPath($OutputPath)
    if (-not [IO.Path]::GetExtension($source).Equals('.doc', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'only legacy .doc files may be normalized'
    }
    if ($source.Equals($destination, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'normalization output must differ from input'
    }
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
    if (Test-Path -LiteralPath $destination) { Remove-Item -LiteralPath $destination -Force }

    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $app.DisplayAlerts = 0
    $app.AutomationSecurity = 3
    # Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles).
    $document = $app.Documents.Open($source, $false, $true, $false)
    try {
        $document.ActiveWindow.View.RevisionsView = 0
        $document.ActiveWindow.View.ShowRevisionsAndComments = $false
    } catch { }
    # wdFormatXMLDocument = 12. Word's SaveCopyAs only supports a subset of
    # conversion paths on older binary documents; SaveAs2 performs the OOXML
    # conversion. The source was opened read-only and Close(0) below suppresses
    # any source save, so the original remains untouched.
    $document.SaveAs2($destination, 12)
    if (-not (Test-Path -LiteralPath $destination) -or (Get-Item -LiteralPath $destination).Length -le 0) {
        throw 'Word did not create a non-empty DOCX copy'
    }
} finally {
    if ($null -ne $document) { $document.Close(0); Release-ComObject $document }
    if ($null -ne $app) { $app.Quit(); Release-ComObject $app }
}
