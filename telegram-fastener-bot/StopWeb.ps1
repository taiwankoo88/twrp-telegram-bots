$webPath = Join-Path $PSScriptRoot 'web.mjs'
Get-CimInstance Win32_Process -Filter "Name = 'node.exe'" | Where-Object { $_.CommandLine -and $_.CommandLine.Contains($webPath) } | ForEach-Object { Stop-Process -Id $_.ProcessId }
# The tunnel is a child of the web server and is left running when that is stopped from outside.
$tunnel = Join-Path $PSScriptRoot 'cloudflared.exe'
Get-CimInstance Win32_Process -Filter "Name = 'cloudflared.exe'" | Where-Object { $_.ExecutablePath -eq $tunnel } | ForEach-Object { Stop-Process -Id $_.ProcessId }
Remove-Item -LiteralPath (Join-Path $PSScriptRoot 'web-url.json') -ErrorAction SilentlyContinue
Write-Host 'Stopped the web order form.'
