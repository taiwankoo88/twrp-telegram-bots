$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
Write-Host 'PDF Table to Excel Bot - local setup'
Write-Host 'Create a NEW bot with @BotFather /newbot (do not reuse the fastener bot). Paste its token below (hidden).'
$botSecret = Read-Host 'Bot token' -AsSecureString
$botPtr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($botSecret)
try {
  $botToken = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($botPtr)
  if ($botToken -notmatch '^\d+:[A-Za-z0-9_-]{20,}$') { throw 'Invalid token format.' }
  try {
    $botInfo = Invoke-RestMethod -Uri ('https://api.telegram.org/bot' + $botToken + '/getMe') -Method Post
    $hookInfo = Invoke-RestMethod -Uri ('https://api.telegram.org/bot' + $botToken + '/getWebhookInfo') -Method Post
  } catch { throw 'Telegram connection failed. Check token and internet access.' }
  if ($hookInfo.result.url) { throw 'This bot already has an integration. Create a new bot in BotFather.' }
  # Two programs polling one bot conflict, so refuse the fastener bot's token.
  $fastenerStatus = 'C:\Users\nicholas.koo\Documents\Codex\2026-08-14\i\work\telegram-fastener-bot\status.json'
  if (Test-Path -LiteralPath $fastenerStatus) {
    $fastener = Get-Content -Raw -LiteralPath $fastenerStatus | ConvertFrom-Json
    if ($fastener.username -eq $botInfo.result.username) { throw 'This is the fastener order bot. Create a separate bot in BotFather for PDF conversion.' }
  }
  $botSecret | ConvertFrom-SecureString | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'token.dpapi') -NoNewline
} finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($botPtr); $botToken = $null }
$codeBytes = New-Object byte[] 8
$rng = [Security.Cryptography.RandomNumberGenerator]::Create()
$rng.GetBytes($codeBytes)
$rng.Dispose()
$pairCode = -join ($codeBytes | ForEach-Object { $_.ToString('x2') })
@{pairCode=$pairCode} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'config.json') -Encoding UTF8
$statePath = Join-Path $PSScriptRoot 'state.json'
if (Test-Path -LiteralPath $statePath) {
  $state = Get-Content -Raw -LiteralPath $statePath | ConvertFrom-Json
  $state.pairFailures = 0
  $state | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
}
Write-Host ('Open Telegram: https://t.me/' + $botInfo.result.username)
Write-Host ('First administrator: send /pair ' + $pairCode)
Write-Host 'Existing administrator pairing is retained on reconfiguration.'
Write-Host 'Keep this PC awake and connected. Run Start.cmd after setup.'
Read-Host 'Press Enter to close'
