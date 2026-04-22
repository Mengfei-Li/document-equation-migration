param(
    [Parameter(Mandatory = $true)]
    [string]$InputDir,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [int]$Limit = 10,

    [int]$StartIndex = 0,

    [int]$EndIndex = 0,

    [string]$JavaExe = '',

    [string]$JavacExe = '',

    [string]$MathtypeExtensionDir = '',

    [string]$MathTypeToMathMlDir = '',

    [string]$Mml2OmmlXsl = '',

    [switch]$SkipLatexPreview,

    [switch]$SkipExisting
)

$ErrorActionPreference = "Stop"

$base = Split-Path -Parent $MyInvocation.MyCommand.Path
$classes = Join-Path $base 'java_bridge\classes'
$mathmlNormalizer = Join-Path $base 'normalize_mathml.py'

function Use-DefaultIfEmpty {
    param(
        [string]$Value,
        [string]$Fallback
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Fallback
    }

    return $Value
}

function Require-Path {
    param(
        [string]$Path,
        [string]$Description
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw ($Description + ' path is empty')
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        throw ($Description + ' not found: ' + $Path)
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

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

function Get-OleObjectIndex {
    param(
        [string]$Name
    )

    if ($Name -match '\d+') {
        return [int]$matches[0]
    }
    return [int]::MaxValue
}

function Get-TextPreview {
    param(
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return ''
    }

    $preview = ((Get-Content -LiteralPath $Path -Encoding utf8 -Raw) -replace '\s+', ' ').Trim()
    if ($preview.Length -gt 120) {
        return $preview.Substring(0, 120)
    }
    return $preview
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

$JavaExe = Use-DefaultIfEmpty $JavaExe (Use-DefaultIfEmpty $env:JAVA_EXE 'java')
$JavacExe = Use-DefaultIfEmpty $JavacExe (Use-DefaultIfEmpty $env:JAVAC_EXE 'javac')
$MathtypeExtensionDir = Use-DefaultIfEmpty $MathtypeExtensionDir (Use-DefaultIfEmpty $env:MATHTYPE_EXTENSION_DIR (Join-Path $base 'third_party\mathtype-extension'))
$MathTypeToMathMlDir = Use-DefaultIfEmpty $MathTypeToMathMlDir (Use-DefaultIfEmpty $env:MATHTYPE_TO_MATHML_DIR (Join-Path $base 'third_party\mathtype_to_mathml'))

if ([string]::IsNullOrWhiteSpace($Mml2OmmlXsl)) {
    $Mml2OmmlXsl = $env:MML2OMML_XSL
}
if ([string]::IsNullOrWhiteSpace($Mml2OmmlXsl)) {
    $officeCandidates = @(
        'C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL',
        'C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL'
    )
    foreach ($candidate in $officeCandidates) {
        if (Test-Path -LiteralPath $candidate) {
            $Mml2OmmlXsl = $candidate
            break
        }
    }
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$rows = New-Object System.Collections.Generic.List[object]
$files = Get-ChildItem -LiteralPath $InputDir -Filter '*.bin' -File |
    Where-Object { $_.BaseName -notlike '*__*' } |
    Sort-Object @{
        Expression = {
            Get-OleObjectIndex $_.BaseName
        }
    }, Name

if ($StartIndex -gt 0) {
    $files = $files | Where-Object { (Get-OleObjectIndex $_.BaseName) -ge $StartIndex }
}
if ($EndIndex -gt 0) {
    $files = $files | Where-Object { (Get-OleObjectIndex $_.BaseName) -le $EndIndex }
}
if ($Limit -gt 0) {
    $files = $files | Select-Object -First $Limit
}
$files = @($files)

$filesToConvert = @(
    $files | Where-Object {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        $ommlPath = Join-Path $OutputDir ($stem + '.omml.xml')
        -not ($SkipExisting -and (Test-Path -LiteralPath $ommlPath))
    }
)

if ($filesToConvert.Count -gt 0) {
    $MathtypeExtensionDir = Require-Path $MathtypeExtensionDir 'mathtype-extension directory'
    $MathTypeToMathMlDir = Require-Path $MathTypeToMathMlDir 'mathtype_to_mathml directory'
    $Mml2OmmlXsl = Require-Path $Mml2OmmlXsl 'MML2OMML.XSL'

    $jar = Require-Path (Join-Path $MathtypeExtensionDir 'jar\MathType2MathML.jar') 'MathType2MathML.jar'
    $jruby = Require-Path (Join-Path $MathtypeExtensionDir 'lib\jruby-complete-9.3.8.0.jar') 'JRuby runtime'
    $rubyBase = Join-Path $MathtypeExtensionDir 'ruby'
    $transformXsl = Require-Path (Join-Path $MathTypeToMathMlDir 'lib\transform.xsl') 'MathType-to-MathML XSLT'

    $cp = @(
        $classes,
        $jar,
        $jruby,
        (Join-Path $rubyBase 'ruby-ole-1.2.12.2\lib'),
        (Join-Path $rubyBase 'nokogiri-1.7.0.1-java\lib'),
        (Join-Path $rubyBase 'bindata-2.3.5\lib'),
        (Join-Path $rubyBase 'mathtype-0.0.7.5\lib')
    ) -join ';'

    $bridgeClass = Join-Path $classes 'Ole2XmlCli.class'
    if (-not (Test-Path -LiteralPath $bridgeClass)) {
        New-Item -ItemType Directory -Force -Path $classes | Out-Null
        $javaSource = Require-Path (Join-Path $base 'java_bridge\Ole2XmlCli.java') 'Ole2XmlCli.java'
        Invoke-Native $JavacExe @('-cp', $cp, '-d', $classes, $javaSource)
    }
}

foreach ($file in $files) {
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $xmlPath = Join-Path $OutputDir ($stem + '.xml')
    $mathmlPath = Join-Path $OutputDir ($stem + '.mathml')
    $ommlPath = Join-Path $OutputDir ($stem + '.omml.xml')
    $texPath = Join-Path $OutputDir ($stem + '.tex')

    $status = 'ok'
    $errorMessage = ''
    $latexPreview = ''
    $mathmlPreview = ''
    $skippedExisting = $false

    try {
        if ($SkipExisting -and (Test-Path -LiteralPath $ommlPath)) {
            $skippedExisting = $true
            $latexPreview = Get-TextPreview $texPath
            $mathmlPreview = Get-TextPreview $mathmlPath
        }
        else {
            Invoke-Native $JavaExe @(
                '--add-opens',
                'java.base/sun.nio.ch=ALL-UNNAMED',
                '--add-opens',
                'java.base/java.io=ALL-UNNAMED',
                '-cp',
                $cp,
                'Ole2XmlCli',
                $file.FullName,
                $xmlPath
            )

            $xslt = New-Object System.Xml.Xsl.XslCompiledTransform
            $xslt.Load($transformXsl)
            $writer = [System.Xml.XmlWriter]::Create($mathmlPath)
            $xslt.Transform($xmlPath, $writer)
            $writer.Close()

            $mathmlRaw = Get-Content -LiteralPath $mathmlPath -Encoding utf8 -Raw
            if ($mathmlRaw -notmatch 'xmlns="http://www\.w3\.org/1998/Math/MathML"') {
                $mathmlRaw = $mathmlRaw -replace '<math(\s|>)', '<math xmlns="http://www.w3.org/1998/Math/MathML"$1'
                Set-Content -LiteralPath $mathmlPath -Value $mathmlRaw -Encoding utf8
            }

            & python $mathmlNormalizer $mathmlPath | Out-Null

            $mml2omml = New-Object System.Xml.Xsl.XslCompiledTransform
            $mml2omml.Load($Mml2OmmlXsl)
            $ommlWriter = [System.Xml.XmlWriter]::Create($ommlPath)
            $mml2omml.Transform($mathmlPath, $ommlWriter)
            $ommlWriter.Close()

            if (-not $SkipLatexPreview) {
                Invoke-Native 'pandoc' @('-f', 'html', '-t', 'latex', $mathmlPath, '-o', $texPath)
                $latexPreview = Get-TextPreview $texPath
            }

            $mathmlPreview = Get-TextPreview $mathmlPath
        }
    }
    catch {
        $status = 'error'
        $errorMessage = $_.Exception.Message
    }

    $rows.Add([pscustomobject]@{
        name = $file.Name
        status = $status
        xml_exists = Test-Path $xmlPath
        mathml_exists = Test-Path $mathmlPath
        omml_exists = Test-Path $ommlPath
        tex_exists = Test-Path $texPath
        skipped_existing = $skippedExisting
        latex_preview = $latexPreview
        mathml_preview = $mathmlPreview
        error = $errorMessage
    })
}

$summaryPath = Join-Path $OutputDir 'summary.csv'
$rows | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding utf8
Write-Output $summaryPath
