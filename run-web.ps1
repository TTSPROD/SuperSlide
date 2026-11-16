$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'run.ps1') serve --open @args
exit $LASTEXITCODE
