# Actualizare Grind Rings Action Tracker: copiaza fisierele programului din acest folder
# peste instalarea existenta si reporneste serverul.
# NU se ating: grind_rings.db, atasamente\, backups\, mail.ini, mail_out\, server.log
$ErrorActionPreference = "Stop"
$Src = $PSScriptRoot.TrimEnd('\')
$Task = "GrindRingsTracker"
function Pauza { Read-Host "`nApasa Enter pentru a inchide" | Out-Null }

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) { Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""; exit }
Write-Host "===== Actualizare Grind Rings Action Tracker =====" -ForegroundColor Cyan

$t = Get-ScheduledTask -TaskName $Task -ErrorAction SilentlyContinue
$Dst = if ($t) { $t.Actions[0].WorkingDirectory } else { "C:\GrindRingsTracker" }
$r = Read-Host "Folderul instalat [$Dst]"
if ($r) { $Dst = $r.Trim('"') }
if (-not (Test-Path (Join-Path $Dst "server.py"))) { Write-Host "Nu gasesc server.py in '$Dst'." -ForegroundColor Red; Pauza; exit 1 }
$Dst = (Resolve-Path $Dst).Path.TrimEnd('\')

Write-Host "Opresc serverul..."
if ($t) { Stop-ScheduledTask -TaskName $Task -ErrorAction SilentlyContinue }
Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^(python|pythonw|py)\.exe$' -and $_.CommandLine -like "*$Dst\server.py*" } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
Start-Sleep -Seconds 2

if ($Dst -ne $Src) {
  $prog = @(".py", ".html", ".bat", ".ps1", ".txt", ".ini", ".nsi")
  $nu = @("mail.ini", "server.log")
  $bk = Join-Path $Dst ("backups\program_" + (Get-Date -Format "yyyy-MM-dd_HHmmss"))
  New-Item -ItemType Directory -Force -Path $bk | Out-Null
  Get-ChildItem $Dst -File | Where-Object { $prog -contains $_.Extension -and $nu -notcontains $_.Name } | Copy-Item -Destination $bk
  Get-ChildItem $Src -File | Where-Object { $prog -contains $_.Extension -and $nu -notcontains $_.Name } | ForEach-Object {
    Copy-Item $_.FullName -Destination $Dst -Force
    Write-Host "  actualizat: $($_.Name)"
  }
  Write-Host "Versiunea veche a programului: $bk"
} else {
  Write-Host "Rulezi din folderul instalat - doar repornesc serverul."
}

Write-Host "Pornesc serverul..."
if ($t) { Start-ScheduledTask -TaskName $Task } else { Start-Process (Join-Path $Dst "porneste-tracker.bat") -WorkingDirectory $Dst }
Start-Sleep -Seconds 3
Write-Host "Gata." -ForegroundColor Green
Start-Process "http://localhost:8092/"
Pauza
