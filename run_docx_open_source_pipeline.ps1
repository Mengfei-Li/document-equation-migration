param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$OutputDocx = '',

    [int]$Limit = 0,

    [switch]$Resume,

    [int]$StartIndex = 0,

    [int]$EndIndex = 0,

    [string]$JavaExe = '',

    [string]$JavacExe = '',

    [string]$MathtypeExtensionDir = '',

    [string]$MathTypeToMathMlDir = '',

    [string]$Mml2OmmlXsl = '',

    [switch]$SkipLatexPreview,

    [switch]$PreserveMathTypeLayout,

    [double]$MathTypeLayoutFactor = 1.01375
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$base = Split-Path -Parent $MyInvocation.MyCommand.Path
$inputPath = (Resolve-Path -LiteralPath $InputDocx).Path

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$outputDirPath = (Resolve-Path -LiteralPath $OutputDir).Path

$stem = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
if ([string]::IsNullOrWhiteSpace($OutputDocx)) {
    $OutputDocx = Join-Path $outputDirPath ($stem + '.omml.docx')
}
else {
    $OutputDocx = [System.IO.Path]::GetFullPath($OutputDocx)
}

$oleDir = Join-Path $outputDirPath 'ole_bins'
$convertedDir = Join-Path $outputDirPath 'converted'
$mapJson = Join-Path $outputDirPath ($stem + '.omml.ole_map.json')
$validationTex = Join-Path $outputDirPath ($stem + '.omml.validation.tex')
$xmlCountsTxt = Join-Path $outputDirPath 'xml_counts.txt'
$summaryJson = Join-Path $outputDirPath 'pipeline_summary.json'
$summaryTxt = Join-Path $outputDirPath 'pipeline_summary.txt'
$replaceScript = Join-Path $base 'replace_docx_ole_with_omml.py'
$mapScript = Join-Path $base 'docx_math_object_map.py'
$probeScript = Join-Path $base 'probe_formula_pipeline.ps1'

function Invoke-Native {
    param(
        [string]$Command,
        [string[]]$Arguments
    )

    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw ($Command + ' failed with exit code ' + $LASTEXITCODE)
    }
}

if ($StartIndex -lt 0) {
    throw 'StartIndex must be greater than or equal to 0.'
}
if ($EndIndex -lt 0) {
    throw 'EndIndex must be greater than or equal to 0.'
}
if (($StartIndex -gt 0) -and ($EndIndex -gt 0) -and ($StartIndex -gt $EndIndex)) {
    throw 'StartIndex must be less than or equal to EndIndex.'
}

foreach ($path in @($oleDir, $convertedDir)) {
    if ((Test-Path -LiteralPath $path) -and (-not $Resume)) {
        Remove-Item -LiteralPath $path -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $path | Out-Null
}

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ('docx_ole_pipeline_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $tempDir | Out-Null

try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($inputPath, $tempDir)

    $embeddingsDir = Join-Path $tempDir 'word\embeddings'
    if (-not (Test-Path -LiteralPath $embeddingsDir)) {
        throw ('Missing word\embeddings in DOCX: ' + $inputPath)
    }

    Get-ChildItem -LiteralPath $embeddingsDir -Filter 'oleObject*.bin' -File |
        Sort-Object @{
            Expression = {
                if ($_.BaseName -match '\d+') {
                    [int]$matches[0]
                }
                else {
                    [int]::MaxValue
                }
            }
        }, Name |
        Copy-Item -Destination $oleDir -Force

    $sourceOleCount = (Get-ChildItem -LiteralPath $oleDir -Filter 'oleObject*.bin' -File).Count
    if ($sourceOleCount -eq 0) {
        throw ('No oleObject*.bin extracted from: ' + $inputPath)
    }

    $probeArgs = @{
        InputDir = $oleDir
        OutputDir = $convertedDir
        Limit = $Limit
        JavaExe = $JavaExe
        JavacExe = $JavacExe
        MathtypeExtensionDir = $MathtypeExtensionDir
        MathTypeToMathMlDir = $MathTypeToMathMlDir
        Mml2OmmlXsl = $Mml2OmmlXsl
    }
    if ($StartIndex -gt 0) {
        $probeArgs.StartIndex = $StartIndex
    }
    if ($EndIndex -gt 0) {
        $probeArgs.EndIndex = $EndIndex
    }
    if ($SkipLatexPreview) {
        $probeArgs.SkipLatexPreview = $true
    }
    if ($Resume) {
        $probeArgs.SkipExisting = $true
    }
    & $probeScript @probeArgs | Out-Null

    $summaryCsv = Join-Path $convertedDir 'summary.csv'
    $rows = @(Import-Csv -LiteralPath $summaryCsv)
    $okRows = @($rows | Where-Object { $_.status -eq 'ok' -and $_.omml_exists -eq 'True' })
    $availableOmmlCount = (Get-ChildItem -LiteralPath $convertedDir -Filter '*.omml.xml' -File).Count
    if ($availableOmmlCount -eq 0) {
        throw 'No OMML output was generated, replacement cannot continue.'
    }

    $replaceArgs = @($replaceScript, $inputPath, $convertedDir, $OutputDocx)
    if ($PreserveMathTypeLayout) {
        $replaceArgs += '--preserve-mathtype-layout'
        $replaceArgs += '--mathtype-layout-factor'
        $replaceArgs += [string]$MathTypeLayoutFactor
    }
    Invoke-Native 'python' $replaceArgs
    Invoke-Native 'python' @($mapScript, $OutputDocx, $mapJson)
    if (-not $SkipLatexPreview) {
        Invoke-Native 'pandoc' @($OutputDocx, '-t', 'latex', '-o', $validationTex)
    }

    $replaceSummaryPath = $OutputDocx + '.replace_summary.json'
    $replaceSummary = Get-Content -LiteralPath $replaceSummaryPath -Raw -Encoding UTF8 | ConvertFrom-Json

    $archive = [System.IO.Compression.ZipFile]::OpenRead($OutputDocx)
    try {
        $entry = $archive.GetEntry('word/document.xml')
        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        try {
            $xmlText = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $archive.Dispose()
    }

    $xmlCounts = [ordered]@{
        ObjectCount = ([regex]::Matches($xmlText, '<w:object\b')).Count
        OMathCount = ([regex]::Matches($xmlText, '<m:oMath\b')).Count
        OMathParaCount = ([regex]::Matches($xmlText, '<m:oMathPara\b')).Count
    }

    $xmlCountLines = @(
        ('ObjectCount=' + $xmlCounts.ObjectCount)
        ('OMathCount=' + $xmlCounts.OMathCount)
        ('OMathParaCount=' + $xmlCounts.OMathParaCount)
    )
    Set-Content -LiteralPath $xmlCountsTxt -Value $xmlCountLines -Encoding UTF8

    $summary = [ordered]@{
        input_docx = $inputPath
        output_dir = $outputDirPath
        output_docx = $OutputDocx
        limit = $Limit
        resume = [bool]$Resume
        start_index = $StartIndex
        end_index = $EndIndex
        source_ole_count = $sourceOleCount
        attempted_count = $rows.Count
        converted_ok_count = $okRows.Count
        converted_error_count = ($rows.Count - $okRows.Count)
        converted_available_count = $availableOmmlCount
        replaced_count = $replaceSummary.replaced_count
        preserve_mathtype_layout = [bool]$PreserveMathTypeLayout
        mathtype_layout_factor = $(if ($PreserveMathTypeLayout) { $MathTypeLayoutFactor } else { $null })
        layout_preservation = $replaceSummary.layout_preservation
        map_json = $mapJson
        validation_tex = $(if ($SkipLatexPreview) { '' } else { $validationTex })
        xml_counts = $xmlCounts
    }
    Set-Content -LiteralPath $summaryJson -Value ($summary | ConvertTo-Json -Depth 5) -Encoding UTF8

    $summaryLines = @(
        ('input_docx=' + $inputPath)
        ('output_docx=' + $OutputDocx)
        ('source_ole_count=' + $sourceOleCount)
        ('resume=' + [bool]$Resume)
        ('start_index=' + $StartIndex)
        ('end_index=' + $EndIndex)
        ('attempted_count=' + $rows.Count)
        ('converted_ok_count=' + $okRows.Count)
        ('converted_error_count=' + ($rows.Count - $okRows.Count))
        ('converted_available_count=' + $availableOmmlCount)
        ('replaced_count=' + $replaceSummary.replaced_count)
        ('validation_tex=' + $(if ($SkipLatexPreview) { '' } else { $validationTex }))
        ('map_json=' + $mapJson)
        ('xml_counts_file=' + $xmlCountsTxt)
    )
    Set-Content -LiteralPath $summaryTxt -Value $summaryLines -Encoding UTF8

    Write-Output $summaryTxt
}
finally {
    if (Test-Path -LiteralPath $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }
}
