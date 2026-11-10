$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$bot = Join-Path $PSScriptRoot 'bot.py'
if (Get-CimInstance Win32_Process -Filter "Name LIKE 'python%'" | Where-Object { $_.CommandLine -and $_.CommandLine.Contains($bot) }) {
  Write-Host 'Bot already running.'
  exit 0
}
$tokenPath = Join-Path $PSScriptRoot 'token.dpapi'
if (!(Test-Path -LiteralPath $tokenPath)) { throw 'Run Setup.cmd first.' }
$python = (Get-Command python -ErrorAction Stop).Source
$encryptedToken = (Get-Content -Raw -LiteralPath $tokenPath).Trim()
if ([string]::IsNullOrWhiteSpace($encryptedToken)) { throw 'Saved token is empty. Run Setup.cmd again.' }
$secret = ConvertTo-SecureString -String $encryptedToken
$ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret)
$statusPath = Join-Path $PSScriptRoot 'status.json'
$startedAt = Get-Date
try {
  $env:TELEGRAM_BOT_TOKEN = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
  $env:PYTHONIOENCODING = 'utf-8'
  Start-Process -FilePath $python -ArgumentList @('-u', ('"' + $bot + '"')) -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -RedirectStandardOutput (Join-Path $PSScriptRoot 'bot.log') -RedirectStandardError (Join-Path $PSScriptRoot 'error.log') | Out-Null
} finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr); Remove-Item Env:TELEGRAM_BOT_TOKEN -ErrorAction SilentlyContinue }
Write-Host 'Connecting to Telegram. Please wait...'
$online = $false
for ($attempt = 0; $attempt -lt 20; $attempt++) {
  Start-Sleep -Seconds 1
  if ((Test-Path -LiteralPath $statusPath) -and (Get-Item -LiteralPath $statusPath).LastWriteTime -ge $startedAt) {
    try { $botStatus = Get-Content -Raw -LiteralPath $statusPath | ConvertFrom-Json; $online = $true; break } catch { }
  }
  $running = Get-CimInstance Win32_Process -Filter "Name LIKE 'python%'" | Where-Object { $_.CommandLine -and $_.CommandLine.Contains($bot) }
  if (!$running -and $attempt -ge 3) { throw 'Bot stopped before connecting. error.log contains the details.' }
}
if ($online) {
  Write-Host ('ONLINE: @' + $botStatus.username) -ForegroundColor Green
  Write-Host 'Send a PDF to the bot in Telegram. You may close this window.'
} else {
  Write-Host 'NOT ONLINE YET: the process is running and retrying the Telegram connection.' -ForegroundColor Yellow
  Write-Host 'Do not repeat Setup or enter your token again. Check bot.log and error.log.'
}
