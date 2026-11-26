# Actualizare Grind Rings Action Tracker: copiaza fisierele programului din acest folder
# (versiunea noua) peste instalarea existenta si reporneste serverul.
# NU se ating: grind_rings.db, attachments\, backups\, mail.ini, mail_out\, server.log.
$ErrorActionPreference = "Stop"
$Src = $PSScriptRoot
$TaskName = "GrindRingsTracker"
function Pauza { Read-Host "`nApasa Enter pentru a inchide" | Out-Null }

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) {
  Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
  exit
}
Write-Host "=== Actualizare Grind Rings Action Tracker ===" -ForegroundColor Cyan

# folderul instalat: din sarcina programata, altfel il cerem
$Dst = $null
$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) { $Dst = $task.Actions[0].WorkingDirectory }
$r = Read-Host "Folderul instalat [$Dst]"
if ($r) { $Dst = $r.Trim('"') }
if (-not $Dst -or -not (Test-Path (Join-Path $Dst "server.py"))) {
  Write-Host "Nu am gasit server.py in '$Dst'." -ForegroundColor Red; Pauza; exit 1
}
$Dst = (Resolve-Path $Dst).Path.TrimEnd('\')
$same = ($Dst -eq $Src.TrimEnd('\'))

Write-Host "Opresc serverul..."
if ($task) { Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue }
Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^(python|pythonw|py)\.exe$' -and $_.CommandLine -like "*$Dst\server.py*" } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
Start-Sleep -Seconds 2

if (-not $same) {
  # copie de siguranta a versiunii vechi a programului (fara date)
  $bk = Join-Path $Dst ("backups\program_" + (Get-Date -Format "yyyy-MM-dd_HHmmss"))
  New-Item -ItemType Directory -Force -Path $bk | Out-Null
  Get-ChildItem $Dst -File | Where-Object { $_.Extension -in ".py",".html",".bat",".ps1",".txt",".ini",".nsi" -and $_.Name -ne "mail.ini" } |
    Copy-Item -Destination $bk
  # fisierele programului
  $keep = @("grind_rings.db","grind_rings.db-wal","grind_rings.db-shm","mail.ini","server.log")
  Get-ChildItem $Src -File | Where-Object { $keep -notcontains $_.Name -and $_.Extension -ne ".exe" } | ForEach-Object {
    Copy-Item $_.FullName -Destination $Dst -Force
    Write-Host "  actualizat: $($_.Name)"
  }
  if (Test-Path (Join-Path $Src "fonts")) {
    Copy-Item (Join-Path $Src "fonts") -Destination $Dst -Recurse -Force
    Write-Host "  actualizat: fonts\"
  }
  Write-Host "Versiunea veche a programului e salvata in $bk"
}

Write-Host "Pornesc serverul..."
if ($task) { Start-ScheduledTask -TaskName $TaskName }
else { Start-Process -FilePath (Join-Path $Dst "porneste-tracker.bat") -WorkingDirectory $Dst }
Start-Sleep -Seconds 3
Write-Host "Gata." -ForegroundColor Green
Start-Process "http://localhost:8092/"
Pauza
