$ErrorActionPreference = 'Stop'
$projectPath = $PSScriptRoot
$candidates = @()
if ($env:SUPERSLIDE_PYTHON) { $candidates += $env:SUPERSLIDE_PYTHON }
$candidates += (Join-Path $projectPath '.venv\Scripts\python.exe')
if ($env:USERPROFILE) {
    $candidates += (Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe')
}
$systemPython = Get-Command python.exe -ErrorAction SilentlyContinue
if ($systemPython -and $systemPython.Source -notmatch 'WindowsApps|PythonManager') {
    $candidates += $systemPython.Source
}
$runtimePath = $candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
if (-not $runtimePath) {
    throw 'Python не найден. Создайте .venv или задайте SUPERSLIDE_PYTHON полным путём к python.exe. См. README.md.'
}
$previousPythonPath = $env:PYTHONPATH
$previousEncoding = $env:PYTHONIOENCODING
$exitCode = 1
try {
    $env:PYTHONPATH = (Join-Path $projectPath 'src') + $(if ($previousPythonPath) { [IO.Path]::PathSeparator + $previousPythonPath } else { '' })
    $env:PYTHONIOENCODING = 'utf-8'
    & $runtimePath -m superslide @args
    $exitCode = $LASTEXITCODE
}
finally {
    $env:PYTHONPATH = $previousPythonPath
    $env:PYTHONIOENCODING = $previousEncoding
}
exit $exitCode
