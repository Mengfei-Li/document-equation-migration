param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$OutputPdf
)

$ErrorActionPreference = "Stop"

$inputPath = (Resolve-Path -LiteralPath $InputDocx).Path
$outputPath = [System.IO.Path]::GetFullPath($OutputPdf)
$outputDir = Split-Path -Parent $outputPath

if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
}

$word = $null
$document = $null
$wdFormatPDF = 17
$wdDoNotSaveChanges = 0

function Open-DocumentWithRepair {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Documents,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        return $Documents.Open($Path, $false, $true)
    }
    catch {
        return $Documents.Open(
            $Path,
            $false,
            $true,
            $false,
            "",
            "",
            $false,
            "",
            "",
            0,
            0,
            $false,
            $false,
            $true
        )
    }
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $document = Open-DocumentWithRepair -Documents $word.Documents -Path $inputPath
    $document.ExportAsFixedFormat($outputPath, $wdFormatPDF)
}
finally {
    if ($document -ne $null) {
        $document.Close([ref]$wdDoNotSaveChanges)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
    }

    if ($word -ne $null) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output $outputPath
