# Actualizare Safety Dashboard: copiaza fisierele programului din acest folder
# peste instalarea existenta si reporneste serverul.
# NU se ating: safety.db, imagini\, backups\, server.log
$ErrorActionPreference = "Stop"
$Src = $PSScriptRoot.TrimEnd('\')
$Task = "SafetyDashboard"
function Pauza { Read-Host "`nApasa Enter pentru a inchide" | Out-Null }

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) { Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""; exit }
Write-Host "===== Actualizare Safety Dashboard =====" -ForegroundColor Cyan

$t = Get-ScheduledTask -TaskName $Task -ErrorAction SilentlyContinue
$Dst = if ($t) { $t.Actions[0].WorkingDirectory } else { "C:\SafetyDashboard" }
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
  $prog = @(".py", ".html", ".bat", ".ps1", ".txt", ".nsi")
  $nu = @("mail.ini", "server.log")
  $bk = Join-Path $Dst ("backups\program_" + (Get-Date -Format "yyyy-MM-dd_HHmmss"))
  New-Item -ItemType Directory -Force -Path $bk | Out-Null
  Get-ChildItem $Dst -File | Where-Object { $prog -contains $_.Extension -and $nu -notcontains $_.Name } | Copy-Item -Destination $bk
  Get-ChildItem $Src -File | Where-Object { $prog -contains $_.Extension -and $nu -notcontains $_.Name } | ForEach-Object {
    Copy-Item $_.FullName -Destination $Dst -Force
    Write-Host "  actualizat: $($_.Name)"
  }
  if (Test-Path (Join-Path $Src "fonts")) { Copy-Item (Join-Path $Src "fonts") -Destination $Dst -Recurse -Force; Write-Host "  actualizat: fonts\" }
  Write-Host "Versiunea veche a programului: $bk"
} else {
  Write-Host "Rulezi din folderul instalat - doar repornesc serverul."
}

Write-Host "Pornesc serverul..."
if ($t) { Start-ScheduledTask -TaskName $Task } else { Start-Process (Join-Path $Dst "porneste-dashboard.bat") -WorkingDirectory $Dst }
Start-Sleep -Seconds 3
Write-Host "Gata." -ForegroundColor Green
Start-Process "http://localhost:8093/"
Pauza
