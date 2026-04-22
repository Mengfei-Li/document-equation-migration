param(
    [string]$ThirdPartyDir = '',

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir

if ([string]::IsNullOrWhiteSpace($ThirdPartyDir)) {
    $ThirdPartyDir = Join-Path $repoRoot 'third_party'
}

New-Item -ItemType Directory -Force -Path $ThirdPartyDir | Out-Null
$thirdPartyRoot = (Resolve-Path -LiteralPath $ThirdPartyDir).Path

function Ensure-GitRepo {
    param(
        [string]$Url,
        [string]$Path
    )

    if (Test-Path -LiteralPath $Path) {
        if (-not (Test-Path -LiteralPath (Join-Path $Path '.git'))) {
            throw ('Path exists but is not a git repository: ' + $Path)
        }

        if ($Force) {
            & git -C $Path fetch --all --prune | Out-Null
        }

        return
    }

    & git clone $Url $Path
}

function Apply-PatchIfNeeded {
    param(
        [string]$RepoPath,
        [string]$PatchPath
    )

    & git -C $RepoPath apply --unidiff-zero --check $PatchPath 2>$null
    if ($LASTEXITCODE -eq 0) {
        & git -C $RepoPath apply --unidiff-zero $PatchPath
        Write-Output ('Applied patch: ' + $PatchPath)
        return
    }

    & git -C $RepoPath apply --unidiff-zero --reverse --check $PatchPath 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Output ('Patch already applied: ' + $PatchPath)
        return
    }

    throw ('Patch cannot be applied cleanly: ' + $PatchPath)
}

$mathtypeExtensionDir = Join-Path $thirdPartyRoot 'mathtype-extension'
$mathTypeToMathMlDir = Join-Path $thirdPartyRoot 'mathtype_to_mathml'

Ensure-GitRepo 'https://github.com/transpect/mathtype-extension.git' $mathtypeExtensionDir
Ensure-GitRepo 'https://github.com/jure/mathtype_to_mathml.git' $mathTypeToMathMlDir

$patch = Join-Path $repoRoot 'patches\mathtype_to_mathml-quality-fixes.patch'
Apply-PatchIfNeeded $mathTypeToMathMlDir $patch

Write-Output ('MATHTYPE_EXTENSION_DIR=' + $mathtypeExtensionDir)
Write-Output ('MATHTYPE_TO_MATHML_DIR=' + $mathTypeToMathMlDir)
