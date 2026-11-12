$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$web = Join-Path $PSScriptRoot 'web.mjs'
if (Get-CimInstance Win32_Process -Filter "Name = 'node.exe'" | Where-Object { $_.CommandLine -and $_.CommandLine.Contains($web) }) {
  Write-Host 'Web order form already running.'
  exit 0
}
# A tunnel left over from a web server that was closed some other way would keep an old address alive.
$tunnel = Join-Path $PSScriptRoot 'cloudflared.exe'
Get-CimInstance Win32_Process -Filter "Name = 'cloudflared.exe'" | Where-Object { $_.ExecutablePath -eq $tunnel } | ForEach-Object { Stop-Process -Id $_.ProcessId }
$node ='C:\Users\nicholas.koo\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'
$log = Join-Path $PSScriptRoot 'web.log'
$process = Start-Process -FilePath $node -ArgumentList ('"' + $web + '"') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $PSScriptRoot 'web-error.log') -PassThru
for ($attempt = 0; $attempt -lt 15; $attempt++) {
  Start-Sleep -Seconds 1
  $process.Refresh()
  if ($process.HasExited) { throw 'The web order form stopped while starting; web-error.log contains the details.' }
  $line = Get-Content -LiteralPath $log -ErrorAction SilentlyContinue | Where-Object { $_ -like 'Web order form online:*' } | Select-Object -Last 1
  if ($line) {
    Write-Host $line -ForegroundColor Green
    Write-Host 'On the office network colleagues open the address with this computer''s network IP; the outside address is in web-url.json.'
    exit 0
  }
}
Write-Host 'NOT ONLINE YET: see web-error.log.' -ForegroundColor Yellow
