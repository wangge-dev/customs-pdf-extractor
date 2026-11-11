param(
    [string]$Python = 'python',
    [string]$MsvcDir = '',
    [string]$OutputRoot = (Join-Path $PSScriptRoot '..\portable_build')
)

$ErrorActionPreference = 'Stop'
$BuildDeps = Join-Path $PSScriptRoot 'builddeps'
$OcrDeps = Join-Path $PSScriptRoot 'ocr_deps'
$Model = Join-Path $PSScriptRoot 'ocr_models\PP-OCRv6_rec_medium.onnx'
if (!(Test-Path -LiteralPath $Model)) { $Model = Join-Path $PSScriptRoot '..\_internal\ocr_models\PP-OCRv6_rec_medium.onnx' }
if (!$MsvcDir) {
    $MsvcDir = Join-Path $PSScriptRoot 'runtime'
    if (!(Test-Path -LiteralPath (Join-Path $MsvcDir 'msvcp140.dll'))) { $MsvcDir = Join-Path $PSScriptRoot '..\_internal\runtime' }
}
$Entry = Join-Path $PSScriptRoot 'main.py'
$pythonCommand = Get-Command $Python -ErrorAction Stop
$Python = $pythonCommand.Source
foreach ($path in @($Model, $Entry, (Join-Path $MsvcDir 'msvcp140.dll'))) {
    if (!(Test-Path -LiteralPath $path)) { throw "Required build input missing: $path" }
}

# Each invocation writes to its own folder, so earlier builds are preserved.
$tag = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 6)
$buildRoot = Join-Path $OutputRoot $tag
New-Item -ItemType Directory -Path $buildRoot -Force | Out-Null
$oldPath, $oldPythonPath, $oldMsvc = $env:PATH, $env:PYTHONPATH, $env:CUSTOMS_BUILD_MSVC
try {
    $env:PATH = "$MsvcDir;$oldPath"
    # Prefer the selected Python environment. Local development copies are
    # optional fallbacks if the respective package is absent from that Python.
    $env:PYTHONPATH = ''
    & $Python -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('PyInstaller') else 1)"
    $hasPyInstaller = $LASTEXITCODE -eq 0
    & $Python -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('rapidocr') else 1)"
    $hasRapidOCR = $LASTEXITCODE -eq 0
    # PyInstaller analyses imports in helper interpreters. Load this app's MSVC
    # there as well so the host's obsolete system DLL cannot crash ONNX import.
    $startup = Join-Path $buildRoot 'startup'
    New-Item -ItemType Directory -Path $startup | Out-Null
    Set-Content -LiteralPath (Join-Path $startup 'sitecustomize.py') -Encoding UTF8 -Value "import ctypes, os`nctypes.WinDLL(os.path.join(os.environ['CUSTOMS_BUILD_MSVC'], 'msvcp140.dll'))"
    $env:CUSTOMS_BUILD_MSVC = (Resolve-Path $MsvcDir).Path
    $modulePaths = @($startup, $PSScriptRoot)
    if (!$hasPyInstaller -and (Test-Path -LiteralPath $BuildDeps)) { $modulePaths += $BuildDeps }
    if (!$hasRapidOCR -and (Test-Path -LiteralPath $OcrDeps)) { $modulePaths += $OcrDeps }
    $env:PYTHONPATH = $modulePaths -join ';'
    $argsList = @(
        '-m', 'PyInstaller', '--noconfirm', '--onedir', '--windowed',
        '--name', '本地PDF批处理',
        '--distpath', (Join-Path $buildRoot 'dist'),
        '--workpath', (Join-Path $buildRoot 'build'),
        '--specpath', $buildRoot,
        '--paths', $PSScriptRoot,
        '--collect-data', 'rapidocr',
        '--collect-binaries', 'onnxruntime',
        '--copy-metadata', 'rapidocr',
        '--copy-metadata', 'onnxruntime',
        '--collect-all', 'pypdfium2',
        '--add-data', "$Model;ocr_models"
    )
    if (!$hasRapidOCR -and (Test-Path -LiteralPath $OcrDeps)) { $argsList += @('--paths', $OcrDeps) }
    foreach ($dllName in @('msvcp140.dll','msvcp140_1.dll','msvcp140_2.dll','msvcp140_atomic_wait.dll','msvcp140_codecvt_ids.dll','vcruntime140.dll','vcruntime140_1.dll','vcruntime140_threads.dll','concrt140.dll')) {
        $dll = Join-Path $MsvcDir $dllName
        if (Test-Path -LiteralPath $dll) { $argsList += @('--add-binary', "$dll;runtime") }
    }
    $argsList += $Entry
    & $Python @argsList
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }
    $appDir = Join-Path (Join-Path $buildRoot 'dist') '本地PDF批处理'
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot '示例字段模板.json') -Destination $appDir
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot '使用说明.txt') -Destination $appDir
    Write-Host "Build complete: $appDir"
    Write-Host 'Copy the entire application folder when sharing. Verify it on another Windows machine with a real PDF.'
}
finally {
    $env:PATH, $env:PYTHONPATH, $env:CUSTOMS_BUILD_MSVC = $oldPath, $oldPythonPath, $oldMsvc
}
