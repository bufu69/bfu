# Instalare Grind Rings Action Tracker (port 8092) pe acest calculator.
#  - verifica Python
#  - creeaza baza de date si primul editor
#  - deschide portul 8092 in firewall
#  - porneste serverul automat la pornirea calculatorului (Task Scheduler, fara utilizator logat)
#  - optional: mail zilnic prin Outlook
$ErrorActionPreference = "Stop"
$Dir = $PSScriptRoot
$TaskName = "GrindRingsTracker"
$Port = 8092
Set-Location $Dir

function Pauza { Read-Host "`nApasa Enter pentru a inchide" | Out-Null }

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) {
  Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
  exit
}

Write-Host "=== Instalare Grind Rings Action Tracker ===" -ForegroundColor Cyan
Write-Host "Folder: $Dir`n"

# 1) Python
$py = $null; $pyw = $null
foreach ($c in @("py", "python")) {
  $cmd = Get-Command $c -ErrorAction SilentlyContinue
  if ($cmd) {
    try {
      $exe = & $cmd.Source $(if ($c -eq "py") { "-3" }) -c "import sys; print(sys.executable)" 2>$null
      if ($exe -and (Test-Path $exe)) { $py = $exe.Trim(); break }
    } catch {}
  }
}
if (-not $py) {
  Write-Host "Python 3 nu este instalat. Instaleaza-l de pe https://www.python.org/downloads/ (bifeaza 'Add python.exe to PATH') si ruleaza din nou instalarea." -ForegroundColor Red
  Pauza; exit 1
}
$pyw = Join-Path (Split-Path $py) "pythonw.exe"
if (-not (Test-Path $pyw)) { $pyw = $py }
Write-Host "Python: $py"

# 2) Baza de date + editori
& $py server.py --help | Out-Null
$ed = & $py -c "import sqlite3; c=sqlite3.connect('grind_rings.db'); print(c.execute('select count(*) from editori').fetchone()[0])"
if ([int]$ed -eq 0) {
  Write-Host "`nNu exista niciun editor. Adauga cel putin unul (nume + cod personal)." -ForegroundColor Yellow
  & $py server.py --editori
} else {
  Write-Host "Editori existenti: $ed (le poti modifica oricand cu editori.bat)"
}

# 3) Firewall
$rule = "Grind Rings Action Tracker $Port"
if (-not (Get-NetFirewallRule -DisplayName $rule -ErrorAction SilentlyContinue)) {
  New-NetFirewallRule -DisplayName $rule -Direction Inbound -Action Allow -Protocol TCP -LocalPort $Port | Out-Null
}
Write-Host "Firewall: portul $Port este deschis."

# 4) Pornire automata
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
$action = New-ScheduledTaskAction -Execute $pyw -Argument "`"$Dir\server.py`"" -WorkingDirectory $Dir
$trigger = New-ScheduledTaskTrigger -AtStartup
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Grind Rings Action Tracker (port $Port)" | Out-Null
Start-ScheduledTask -TaskName $TaskName
Write-Host "Serverul porneste automat la pornirea calculatorului (sarcina '$TaskName') si ruleaza acum."

# 5) Mail zilnic prin Outlook (optional)
$r = Read-Host "`nProgramezi mailul zilnic prin Outlook (luni-vineri)? D/N"
if ($r -match '^[dDyY]') {
  $ora = Read-Host "Ora (ex. 07:30)"
  if (-not $ora) { $ora = "07:30" }
  if (-not (Test-Path (Join-Path $Dir "mail.ini"))) {
    Copy-Item (Join-Path $Dir "mail-exemplu.ini") (Join-Path $Dir "mail.ini")
    Write-Host "Am creat mail.ini - completeaza adresele in el (se deschide acum)." -ForegroundColor Yellow
    Start-Process notepad (Join-Path $Dir "mail.ini")
  }
  $user = "$env:USERDOMAIN\$env:USERNAME"
  Unregister-ScheduledTask -TaskName "$TaskName-Mail" -Confirm:$false -ErrorAction SilentlyContinue
  $a2 = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"`"$Dir\outlook-mailuri.bat`" /auto`"" -WorkingDirectory $Dir
  $t2 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At $ora
  $p2 = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive
  Register-ScheduledTask -TaskName "$TaskName-Mail" -Action $a2 -Trigger $t2 -Principal $p2 -Description "Mail zilnic Grind Rings Action Tracker (Outlook)" | Out-Null
  Write-Host "Mailul se trimite luni-vineri la $ora, prin Outlook, cand $user este logat."
}

Start-Sleep -Seconds 2
$ip = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notmatch '^(127|169\.254)\.' } | Select-Object -First 1).IPAddress
Write-Host "`nGata! Dashboard-ul: http://$($env:COMPUTERNAME):$Port/   (sau http://$($ip):$Port/)" -ForegroundColor Green
Start-Process "http://localhost:$Port/"
Pauza
