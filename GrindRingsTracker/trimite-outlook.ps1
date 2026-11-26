# Trimite prin Outlook mailurile pregatite de "server.py --raport-mail" in folderul mail_out.
# -Afiseaza : doar deschide mailurile in Outlook (nu le trimite).
param([switch]$Afiseaza)
$ErrorActionPreference = "Stop"
$dir = Join-Path $PSScriptRoot "mail_out"
$index = Join-Path $dir "mailuri.json"
if (-not (Test-Path $index)) { Write-Host "Nu exista $index. Ruleaza intai: server.py --raport-mail"; exit 1 }
$mails = Get-Content $index -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $mails -or @($mails).Count -eq 0) { Write-Host "Nimic de trimis."; exit 0 }
try { $outlook = New-Object -ComObject Outlook.Application }
catch { Write-Host "Nu am putut porni Outlook pe acest calculator: $($_.Exception.Message)"; exit 1 }
$n = 0
foreach ($m in @($mails)) {
  $item = $outlook.CreateItem(0)
  $item.To = $m.to
  if ($m.cc) { $item.CC = $m.cc }
  $item.Subject = $m.subject
  $item.HTMLBody = Get-Content (Join-Path $dir $m.file) -Raw -Encoding UTF8
  if ($Afiseaza) { $item.Display() } else { $item.Send() }
  $n++
  Write-Host "  $($m.to) - $($m.subject)"
}
if ($Afiseaza) { Write-Host "$n mailuri deschise in Outlook." } else { Write-Host "$n mailuri trimise prin Outlook." }
