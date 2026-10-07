param(
    [string]$MovementPath,
    [string]$Month = '2026-09',
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
$runtimeRoot = Join-Path $env:USERPROFILE '.cache/codex-runtimes/codex-primary-runtime/dependencies'
$nodePath = Join-Path $runtimeRoot 'node/bin/node.exe'
$env:SZCH_PYTHON = Join-Path $runtimeRoot 'python/python.exe'
$env:SZCH_NODE_MODULES = Join-Path $runtimeRoot 'node/node_modules'

if (-not $MovementPath) {
    $MovementPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'РУХ_last.xlsx'
}
if (-not $OutputPath) {
    $OutputPath = Join-Path $taskRoot "Результат/СЗЧ_${Month}_с_итоговым_списком.xlsx"
}
foreach ($requiredPath in @($nodePath, $env:SZCH_PYTHON, $env:SZCH_NODE_MODULES, $MovementPath)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "Не найден путь: $requiredPath. Для другой версии среды используйте load_workspace_dependencies и запустите build.mjs с соответствующими SZCH_PYTHON и SZCH_NODE_MODULES."
    }
}

$builderOutput = & $nodePath (Join-Path $PSScriptRoot 'build.mjs') --sources-dir $taskRoot --movement $MovementPath --month $Month --output $OutputPath
if ($LASTEXITCODE -ne 0) { throw "Создание книги завершилось с кодом $LASTEXITCODE" }
$builderOutput | Write-Output
$buildSummary = $builderOutput[-1] | ConvertFrom-Json
$env:PYTHONIOENCODING = 'utf-8'
& $env:SZCH_PYTHON (Join-Path $PSScriptRoot 'verify.py') $buildSummary.output
if ($LASTEXITCODE -ne 0) { throw 'Результат не прошёл проверку структуры, дат или оформления' }