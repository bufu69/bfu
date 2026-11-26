# Trimite prin Outlook mailurile pregatite de "server.py --raport-mail" (folderul mail_out).
#   -Afiseaza  doar deschide mailurile in Outlook, fara sa le trimita
param([switch]$Afiseaza)
$ErrorActionPreference = "Stop"
$dir = Join-Path $PSScriptRoot "mail_out"
$index = Join-Path $dir "mailuri.json"
if (-not (Test-Path $index)) { Write-Host "Nu exista $index - ruleaza intai: server.py --raport-mail"; exit 1 }
$mailuri = @(Get-Content $index -Raw -Encoding UTF8 | ConvertFrom-Json)
if ($mailuri.Count -eq 0) { Write-Host "Nimic de trimis."; exit 0 }
try { $outlook = New-Object -ComObject Outlook.Application }
catch { Write-Host "Nu pot porni Outlook: $($_.Exception.Message)"; exit 1 }
$n = 0
foreach ($m in $mailuri) {
  $item = $outlook.CreateItem(0)
  $item.To = $m.catre
  if ($m.cc) { $item.CC = $m.cc }
  $item.Subject = $m.subiect
  $item.HTMLBody = Get-Content (Join-Path $dir $m.fisier) -Raw -Encoding UTF8
  if ($Afiseaza) { $item.Display() } else { $item.Send() }
  $n++
  Write-Host "  $($m.catre) - $($m.subiect)"
}
if ($Afiseaza) { Write-Host "$n mailuri deschise in Outlook." } else { Write-Host "$n mailuri trimise prin Outlook." }
