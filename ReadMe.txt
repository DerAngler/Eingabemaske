Powershell-Scheduler

Ein einzelnes Powershell-Skript, mit welchem Tasks im Aufgabenplaner eines Zielservers erstellt und verwaltet werden können.

Simple Nutzung per GUI möglich. Einfach das Skript in einer Powershell ausführen / aufrufen
(Je nach Zielsystem ggf. Ausführung mit entsprechenden Berechtigungen oder als Admin nötig)

Für die Nutzung per Kommandozeile gibt es folgende Parameter:

[string]$Job = "Jobname" Name vom Job, Skript, Programm, etc
[array]$Hosts = "PC1, PC2, PC3" Array mit Hostnamen. Trennung standardmäßig mit Komma(,)
[string]$Spalter = "," Das Trennzeichen, mit welchem die Hosts aus einer Zeile in eigene Werte gesplited werden
[switch]$Switch = Tauscht Hosts und Jobs aus, sprich man weißt dann einem Host mehrere Jobs zu
[datime]$Wann = "26-12-1997 23:07" oder "2025-05-25 04:20" Zeitpunkt, an dem der Aufgabenplaner den Task triggert
[switch]$NoJobNeustart = Parameter, der im Task mitgegeben wird (Ist für einen Sonderfall, kann ignoriert/entfernt werden)
[switch]$Silent = Skript wird ohne GUI ausgeführt, für Ausführungen aus anderen Skripten heraus
[switch]$S = Kurzform von $Silent
[switch]$Darkmode = Macht alles außer Eingabeelemente Schwarz mit grauen Text
[switch]$NoTaskForce = Verhindert das Überschreibung von bereits vorhandenen Tasks, falls der Taskname schon vorhanden ist
[string]$LogPfad = Pfad des Logs. Ein abschließendes "\" kann gesetzt werden, tut aber nicht Not
[string]$LogName = Name des Logs
[string]$ErrorLogPfad = Pfad des Error-Logs. Ein abschließendes "\" kann gesetzt werden, tut aber nicht Not
[string]$ErrorLogName = Name des Error-Logs
[string]$RemoteFQDN = Host, auf dem die Tasks angelegt werden sollen. Für eine localhost Ausführung bitte $null angeben
[string]$TaskPath = Pfad im Aufgabenplaner, in dem die Tasks geschrieben werden und der "Tasks"-Button seine ausliest
[string]$TempLogFile = Dateipfad, in den in Sonderfällen auf Wunsch ein teporäres Log geschrieben wird (GUI Ausführung)

Beispiel Aufruf per Powershell oder CMD:
powershell.exe -ep bypass -windowstyle hidden -noprofile -file "C:\...\Scheduler.ps1"

Beispiele Parameter:

Silent:
-Job "Jobname" -Hosts "PC Name 1, PC-Name-2, 3.er PC, PCName4" -Wann "20-04-2025 07:06" -Silent

Remote:
-Job "Jobname" -Hosts "PC Name 1, PCName2, PC-Name-3" -Wann "20-04-2025 07:06" -RemoteFQDN "Server.domain.topleveldomain"

Jobs an Host zuweisen (Switched):
-Job "Hostname" -Hosts "Jobname1, Jobname(2)!, Job 3" -Switched -Wann "20-04-2025 07:06" -S

Soft-Ausführung:
-Job "Jobname" -Hosts "PC Name 1, PC-Name-2, PC-Name-3" -Wann "20-04-2025 07:06" -NoTaskForce -NoJobNeustart -S

Spalter (Host-Array durch '#!\' trennen):
-Job "Jobname" -Hosts "PC Name 1 # PC-Name-2 # 3.er PC # PCName4" -Spalter "#" -Wann "20-04-2025 07:06" -S

So viele wie gehen (Servername muss angepasst werden):
-Job "PC-Name" -Hosts "Job1; Job2" -Spalter ";" -Switch -Wann "20-04-2025 07:06" -NoTaskForce -NoJobNeustart -Silent -LogPfad "C:\Temp" -LogName "Test.log" -ErrorLogPfad "C:\Fehlerlogs\" -ErrorLogName "ErrorTest.log" -RemoteFQDN "servername" -TaskPath "\Test\" -TempLogFile "%appdata%"
