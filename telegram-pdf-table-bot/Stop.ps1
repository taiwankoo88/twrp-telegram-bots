$bot = Join-Path $PSScriptRoot 'bot.py'
Get-CimInstance Win32_Process -Filter "Name LIKE 'python%'" | Where-Object { $_.CommandLine -and $_.CommandLine.Contains($bot) } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
Write-Host 'Stopped the PDF Table to Excel bot.'
