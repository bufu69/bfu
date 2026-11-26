# Instalare Grind Rings Action Tracker (port 8092)
#  1. verifica Python 3
#  2. creeaza baza de date si primul editor
#  3. deschide portul 8092 in firewall
#  4. porneste serverul automat la pornirea calculatorului (fara utilizator logat)
#  5. optional: mail zilnic prin Outlook
$ErrorActionPreference = "Stop"
$Dir = $PSScriptRoot
$Task = "GrindRingsTracker"
$Port = 8092
Set-Location $Dir
function Pauza { Read-Host "`nApasa Enter pentru a inchide" | Out-Null }

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) { Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""; exit }

Write-Host "===== Instalare Grind Rings Action Tracker =====" -ForegroundColor Cyan
Write-Host "Folder: $Dir`n"

# 0. Folderul: OneDrive nu e potrivit pentru server (sincronizarea poate strica baza de date,
#    iar serverul pornit automat, ca SYSTEM, poate sa nu aiba acces la OneDrive)
if ($Dir -match 'OneDrive') {
  Write-Host "Folderul este in OneDrive. Pentru server e recomandat un folder local." -ForegroundColor Yellow
  $r = Read-Host "Il copiez in C:\GrindRingsTracker si instalez de acolo? D/N"
  if ($r -match '^[dDyY]') {
    $nou = "C:\GrindRingsTracker"
    New-Item -ItemType Directory -Force -Path $nou | Out-Null
    Copy-Item -Path (Join-Path $Dir "*") -Destination $nou -Recurse -Force
    Write-Host "Copiat in $nou - continui instalarea de acolo.`n"
    Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$nou\instalare-server.ps1`""
    exit
  }
}

# 1. Python: intai in folderul tracker-ului, apoi in PATH, apoi in locurile obisnuite
function Test-Python($exe) {
  if (-not $exe -or -not (Test-Path $exe) -or $exe -like '*\WindowsApps\*') { return $false }
  try { $v = & $exe -c "import sys, sqlite3; print(sys.version_info[0])" 2>$null; return ($v -eq "3") } catch { return $false }
}
$candidati = New-Object System.Collections.Generic.List[string]
# a) in folder (ex. python\python.exe, Python311\python.exe, sau langa server.py)
Get-ChildItem -Path $Dir -Filter python.exe -Recurse -Depth 3 -ErrorAction SilentlyContinue |
  Sort-Object { $_.FullName.Length } | ForEach-Object { $candidati.Add($_.FullName) }
# b) salvat la o instalare anterioara
$fisPy = Join-Path $Dir "python.txt"
if (Test-Path $fisPy) { $candidati.Add((Get-Content $fisPy -TotalCount 1).Trim()) }
# c) PATH (py launcher / python)
foreach ($c in @("py", "python")) {
  $cmd = Get-Command $c -ErrorAction SilentlyContinue
  if (-not $cmd) { continue }
  try {
    if ($c -eq "py") { $e = & $cmd.Source -3 -c "import sys; print(sys.executable)" 2>$null }
    else { $e = & $cmd.Source -c "import sys; print(sys.executable)" 2>$null }
    if ($e) { $candidati.Add($e.Trim()) }
  } catch {}
}
# d) locuri obisnuite de instalare (si pentru alti utilizatori)
$tipare = @("$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe", "C:\Users\*\AppData\Local\Programs\Python\Python3*\python.exe",
            "C:\Program Files\Python3*\python.exe", "C:\Python3*\python.exe", "C:\Program Files (x86)\Python3*\python.exe")
foreach ($t in $tipare) { Get-ChildItem $t -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | ForEach-Object { $candidati.Add($_.FullName) } }

$py = $null
foreach ($c in $candidati) { if (Test-Python $c) { $py = $c; break } }
while (-not $py) {
  Write-Host "Nu am gasit Python 3." -ForegroundColor Yellow
  Write-Host "Scrie calea completa catre python.exe (ex. C:\Python312\python.exe)," -ForegroundColor Yellow
  $r = Read-Host "sau apasa Enter ca sa renunti (si instalezi Python de pe https://www.python.org/downloads/)"
  if (-not $r) { Pauza; exit 1 }
  $r = $r.Trim('"')
  if ((Test-Path $r -PathType Container)) { $r = Join-Path $r "python.exe" }
  if (Test-Python $r) { $py = $r } else { Write-Host "'$r' nu este un Python 3 functional." -ForegroundColor Red }
}
$py = (Resolve-Path $py).Path
$pyw = Join-Path (Split-Path $py) "pythonw.exe"
if (-not (Test-Path $pyw)) { $pyw = $py }
# celelalte .bat-uri (editori, import, mail...) folosesc acelasi Python
Set-Content -Path $fisPy -Value $py -Encoding Default
Write-Host "Python: $py"

# 2. Baza de date + primul editor
$n = [int](& $py "$Dir\server.py" --numar-editori)
if ($n -eq 0) {
  Write-Host "`nNu exista niciun editor. Adauga cel putin unul (nume + cod personal)." -ForegroundColor Yellow
  & $py "$Dir\server.py" --editori
} else { Write-Host "Editori: $n (le schimbi oricand cu editori.bat)" }

# 3. Firewall
$regula = "Grind Rings Action Tracker $Port"
if (-not (Get-NetFirewallRule -DisplayName $regula -ErrorAction SilentlyContinue)) {
  New-NetFirewallRule -DisplayName $regula -Direction Inbound -Action Allow -Protocol TCP -LocalPort $Port | Out-Null
}
Write-Host "Firewall: portul $Port este deschis."

# 4. Pornire automata
Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^(python|pythonw|py)\.exe$' -and $_.CommandLine -like "*$Dir\server.py*" } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
Unregister-ScheduledTask -TaskName $Task -Confirm:$false -ErrorAction SilentlyContinue
$actiune = New-ScheduledTaskAction -Execute $pyw -Argument "`"$Dir\server.py`"" -WorkingDirectory $Dir
$declansator = New-ScheduledTaskTrigger -AtStartup
$setari = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$cont = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName $Task -Action $actiune -Trigger $declansator -Settings $setari -Principal $cont -Description "Grind Rings Action Tracker (port $Port)" | Out-Null
Start-ScheduledTask -TaskName $Task
Write-Host "Serverul porneste singur odata cu calculatorul si ruleaza acum."

# 5. Mail zilnic (optional)
$r = Read-Host "`nProgramezi mailul zilnic prin Outlook (luni-vineri)? D/N"
if ($r -match '^[dDyY]') {
  $ora = Read-Host "Ora (ex. 07:30)"
  if (-not $ora) { $ora = "07:30" }
  if (-not (Test-Path (Join-Path $Dir "mail.ini"))) {
    Copy-Item (Join-Path $Dir "mail-exemplu.ini") (Join-Path $Dir "mail.ini")
    Write-Host "Am creat mail.ini - completeaza adresele (se deschide acum)." -ForegroundColor Yellow
    Start-Process notepad (Join-Path $Dir "mail.ini")
  }
  $user = "$env:USERDOMAIN\$env:USERNAME"
  Unregister-ScheduledTask -TaskName "$Task-Mail" -Confirm:$false -ErrorAction SilentlyContinue
  $a2 = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"`"$Dir\outlook-mailuri.bat`" /auto`"" -WorkingDirectory $Dir
  $t2 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At $ora
  $p2 = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive
  Register-ScheduledTask -TaskName "$Task-Mail" -Action $a2 -Trigger $t2 -Principal $p2 -Description "Mail zilnic Grind Rings (Outlook)" | Out-Null
  Write-Host "Mailul pleaca luni-vineri la $ora prin Outlook, cand $user este logat."
}

Start-Sleep -Seconds 2
$ip = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notmatch '^(127|169\.254)\.' } | Select-Object -First 1).IPAddress
Write-Host "`nGata! Tracker: http://$($env:COMPUTERNAME):$Port/   sau   http://$($ip):$Port/" -ForegroundColor Green
Start-Process "http://localhost:$Port/"
Pauza
