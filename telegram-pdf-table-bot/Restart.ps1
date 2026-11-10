$ErrorActionPreference = 'Stop'
Write-Host 'Restarting the PDF Table to Excel bot...'
& (Join-Path $PSScriptRoot 'Stop.ps1')
Start-Sleep -Seconds 1
& (Join-Path $PSScriptRoot 'Start.ps1')
