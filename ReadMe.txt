[Type]$Parameter = "Beispiel"
----------------------------------
[string]$Job = "Patchupdates"
[string]$Hosts = "PC1 PC2 PC3"
[string]$Wann = "26-12-1997 16:20" #"1997-12-26 16:20" geht auch
[switch]$NoJobNeustart = $false,
[switch]$S = $false,
[switch]$Silent = $false,
[switch]$Darkmode = $false,
[switch]$NoTaskForce = $false,
[string]$LogPfad = "C:\Logs\Github\Scheduler\"
[string]$LogName = $(Get-Date -Format "yyyyMMdd") + "_Scheduler.log"
[string]$LogPfad = "C:\Logs\Github\Scheduler\"
[string]$ErrorLogName = $(Get-Date -Format "yyyyMMdd") + "_Error_Scheduler.log"
[string]$RemoteFQDN = $null
[string]$TaskPath = "\JobScheduler\",
[string]$TempLogFile = "C:\Temp\tmpLog_Scheduler.log"

Erklärung der Parameter:
$Job = Name des Jobs, Skripts, Programm, etc
$Hosts = Array mit Hostnamen. Trennung durch Leerzeichen oder Zeilenumbruch
$Wann = Zeitpunkt, an dem der Aufgabenplaner den Task triggert
$NoJobNeustart = Ob der Job im Zielsystem neu gestartet werden darf (Kann nach belieben umfunktioniert / entfernt werden... auch aus GUI)
$S = Kurzform vom $Silent
$Silent = Ermöglicht eine Ausführung des Skripts im Hintergrund
$Darkmode = Macht alles außer Eingabeelemente Schwarz mit grauen Text
$NoTaskForce = Verhindert das Überschreibung von bereits vorhandenen Tasks, falls Jobname, Hostname und Initiator übereinstimmt
$LogPfad = Pfad des Logs. Ein abschließendes "\" kann gesetzt werden, tut aber nicht Not
$LogName = Name des Logs
$ErrorLogPfad = Pfad des Error-Logs. Ein abschließendes "\" kann gesetzt werden, tut aber nicht Not
$ErrorLogName = Name des Error-Logs
$RemoteFQDN = Das ist der Verbindungsparameter für eine CIM-Session. Für eine localhost AUsführung bitte $null angeben
$TaskPath = Pfad im Aufgabenplaner, in den die Tasks geschrieben werden und der "Tasks"-Button seine ausliest
$TempLogFile = Dateipfad, in den bei einer non-Silent-Ausführung auf Wunsch eine Liste fehlgeschlagener Hosts exportiert werden

Beispiel Silent Taskanlage per Powershell oder CMD:
powershell.exe -ep bypass -windowstyle hidden -noprofile -file "C:\...\Scheduler.ps1" -Job "TestJob" -Hosts "PC Name 1; PC Name 2; ..." -Wann "25-05-2025 04:20" -Silent