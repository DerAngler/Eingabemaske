#Region Variablen und Parameter
param(
    [string]$Job = "",
    [string]$Hosts = @(""),
    [string]$Wann = "",
    [switch]$NoJobNeustart = $false,
    [switch]$S = $false,
    [switch]$Silent = $false,
    [switch]$Darkmode = $false,
    [switch]$NoTaskForce = $false,
    [string]$LogPfad = "",
    [string]$LogName = "",
    [string]$ErrorLogPfad = "",
    [string]$ErrorLogName = "",
    [string]$RemoteFQDN = $null,
    [string]$TaskPath = "",
    [string]$TempLogFile = ""
)
###############################################################
##### Diese Variablen können nach Bedarf angepasst werden #####
###############################################################

# Log definieren
$LogPfad = "C:\Logs\Github\Scheduler\"
$LogName = $(Get-Date -Format "yyyyMMdd") + "Scheduler.log"
$ErrorLogPfad = $LogPfad
$ErrorLogName = $(Get-Date -Format "yyyyMMdd") + "Scheduler.log"

# Error Meldungen in powershell.exe verstecken
$ErrorActionPreference = "SilentlyContinue"

# Cim-Session Endpoint (Bei Planlosigkeit einfach so lassen)
$RemoteFQDN = $null #$null oder leer = localhost 

# Name vom Pfad/Ordner in der Aufgabenplanung des Remote-Servers
$TaskPath = "\JobScheduler\"

#Dateipfad für das persönliche / temporäre Logfile
$TempLogFile = "C:\Temp\" + $(Get-Date -Format "yyyyMMdd") + "_fehlgeschlagene_Clients.log"

###############################################################
################### Ab hier nix mehr ändern ###################
###############################################################
#Endregion Variablen

#Region Funktionen
function Datei_auswählen() {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = "shell:MyComputerFolder"
    $OpenFileDialog.Filter = "Textdateien (*.txt)| *.txt|Logs (*.log)| *.log|Alle Dateien (*.*)|*.*"
    if($OpenFileDialog.ShowDialog() -eq "OK"){
        return $OpenFileDialog.FileName
    }
}

function Dateiinhalt_per_OpenFileDialog_in_ne_CheckBox_schreiben($Box) {
    $DateiInhalt = Get-Content -LiteralPath "$(Datei_auswählen)"
    foreach($Zeile in $Dateiinhalt){
        if(!([string]::IsNullOrEmpty($Zeile))){
            $Hostnames.Text += "$($Zeile.Trim())`r`n"
        }
    }
}

function Darkmode {
    param (
        $DasWasDunkelGemachtWerdenSoll
    )
    if($Darkmode -eq $true){
        $DasWasDunkelGemachtWerdenSoll.ForeColor = "LightGray"
        $DasWasDunkelGemachtWerdenSoll.BackColor = "Black"
    }
}
#Endregion Funktionen

#Region GUI

 #Region GUI-Label
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles() #Wenn der Kalender Button nicht geht, dann diese Zeile auskommentieren und neue PS Session starten

# Main Form
$mainForm = New-Object System.Windows.Forms.Form
$Font = New-Object System.Drawing.Font("Courier New", 12)
$mainForm.Text = "Scheduler"
$mainForm.Font = $Font
$mainForm.Height = 420
$mainForm.Width = 282
$mainForm.MaximizeBox = $false
$mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$mainForm.StartPosition = "CenterScreen"
$mainForm.ShowIcon = $false
Darkmode($mainForm)

# Jobname Label
$JobnameLabel = New-Object System.Windows.Forms.Label
$JobnameLabel.Text = "Jobname"
$JobnameLabel.Location = "5, 10"
$JobnameLabel.Height = 22
$JobnameLabel.Width = 90
Darkmode($JobnameLabel)
$mainForm.Controls.Add($JobnameLabel)

# Hostnames Label
$HostnamesLabel = New-Object System.Windows.Forms.Label
$HostnamesLabel.Text = "Hostnames"
$HostnamesLabel.Location = "5, 47"
$HostnamesLabel.Height = 22
$HostnamesLabel.Width = 95
Darkmode($HostnamesLabel)
$mainForm.Controls.Add($HostnamesLabel)

# Datum Label
$DatumLabel = New-Object System.Windows.Forms.Label
$DatumLabel.Text = "Datum"
$DatumLabel.Location = "5, 180"
$DatumLabel.Height = 22
$DatumLabel.Width = 90
Darkmode($DatumLabel)
$mainForm.Controls.Add($DatumLabel)

# Uhrzeit Label
$UhrzeitLabel = New-Object System.Windows.Forms.Label
$UhrzeitLabel.Text = "Uhrzeit"
$UhrzeitLabel.Location = "5, 215"
$UhrzeitLabel.Height = 22
$UhrzeitLabel.Width = 90
Darkmode($UhrzeitLabel)
$mainForm.Controls.Add($UhrzeitLabel)

# Neustart Label
$NeustartLabel = New-Object System.Windows.Forms.Label
$NeustartLabel.Text = "Job-Neustart erlauben?"
$NeustartLabel.Location = "5, 250"
$NeustartLabel.Height = 22
$NeustartLabel.Width = 230
Darkmode($NeustartLabel)
$mainForm.Controls.Add($NeustartLabel)

# Rückinfos Label
$RueckinfosLabel = New-Object System.Windows.Forms.Label
$RueckinfosLabel.Text = "Rückinfo?"
$RueckinfosLabel.Location = "5, 283"
$RueckinfosLabel.Height = 22
$RueckinfosLabel.Width = 95
Darkmode($RueckinfosLabel)
$mainForm.Controls.Add($RueckinfosLabel)

 #Endregion GUI-Label

 #Region GUI-Elemente

# Jobname Auswahlfeld
$JobFeld = New-Object System.Windows.Forms.Textbox
$JobFeld.Location = "100, 7"
$JobFeld.Width = "155"
$mainForm.Controls.Add($JobFeld)

# Hostnames Auswahlfeld
$Hostnames = New-Object System.Windows.Forms.TextBox
$Hostnames.Location = "100, 42"
$Hostnames.Height = "125"
$Hostnames.Width = "155"
$Hostnames.Multiline = $TRUE
$Hostnames.ScrollBars = "Vertical"
$Hostnames.AcceptsReturn = $true
$Hostnames.WordWrap = $true
$mainForm.Controls.Add($Hostnames)

# Datum Auswahlfeld
$Datum = New-Object System.Windows.Forms.DateTimePicker
$Datum.Location = "100, 177"
$Datum.Width = "155"
$Datum.Format = [windows.forms.datetimepickerFormat]::custom
$Datum.CustomFormat = "dd/MM/yyyy"
$mainForm.Controls.Add($Datum)

# Uhrzeit Auswahlfeld
$Uhrzeit = New-Object System.Windows.Forms.DateTimePicker
$Uhrzeit.Location = "100, 212"
$Uhrzeit.Width = "155"
$Uhrzeit.Format = [windows.forms.datetimepickerFormat]::custom
$Uhrzeit.CustomFormat = "HH:mm"
$Uhrzeit.ShowUpDown = $TRUE
$Uhrzeit.BackColor = "Black"
$mainForm.Controls.Add($Uhrzeit)

# Neustart CheckBox
$Neustart = New-Object System.Windows.Forms.CheckBox
$Neustart.Location = "242, 248"
$Neustart.Checked = $TRUE
$mainForm.Controls.Add($Neustart)

# Rückinfo-Dropdown Menü (ComboBox)
$Rueckinfo = New-Object System.Windows.Forms.ComboBox
$Rueckinfo.Location = "100, 280"
$Rueckinfo.Width = "155"
$Rueckinfo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
@("Für jeden Host","Fehler + Gesamt","Nur am Ende","Nie(/Silent,/S)") | ForEach-Object {[void]$Rueckinfo.Items.Add($_)}
$Rueckinfo.SelectedIndex = 2
$mainForm.Controls.Add($Rueckinfo)

# Zuweisen Button
$Zuweisen = New-Object System.Windows.Forms.Button
$Zuweisen.Location = "5, 313"#282
$Zuweisen.Height = "32"
$Zuweisen.Width = "252"
$Zuweisen.Text = "Zuweisung einplanen"
$Zuweisen.BackColor = "White"
$Zuweisen.ForeColor = "Black"
$Zuweisen.add_Click({Zuweisung_einplanen})
$mainForm.Controls.Add($Zuweisen)

# Task-Button
$TasksButton = New-Object System.Windows.Forms.Button
$TasksButton.Location = "5, 350"
$TasksButton.Width = "124"
$TasksButton.Text = "Tasks"
$TasksButton.BackColor = "White"
$TasksButton.ForeColor = "Black"
$TasksButton.add_Click({Get-ScheduledTask -TaskPath $TaskPath | Get-ScheduledTaskInfo | Out-GridView -Title "JobScheduler - Markierte Tasks werden mit 'OK' gelöscht!!!" -PassThru | Unregister-ScheduledTask -Confirm:$false})
$mainForm.Controls.Add($TasksButton)

# Logs-Button
$LogsButton = New-Object System.Windows.Forms.Button
$LogsButton.Location = "133, 350"
$LogsButton.Width = "124"
$LogsButton.Text = "Logs"
$LogsButton.BackColor = "White"
$LogsButton.ForeColor = "Black"
$LogsButton.add_Click({(explorer.exe $LogPfad)})
$mainForm.Controls.Add($LogsButton)

# Import Hostnames from File - Button
$ImportHosts = New-Object System.Windows.Forms.Button
$ImportHosts.Location = "17, 70"
$ImportHosts.Height = "20"
$ImportHosts.Width = "65"
$ImportHosts.Font = ("Courier New, 10")
$ImportHosts.Text = "Import"
$ImportHosts.BackColor = "White"
$ImportHosts.ForeColor = "Black"
$ImportHosts.add_Click({Dateiinhalt_per_OpenFileDialog_in_ne_CheckBox_schreiben($Hostnames)})
$mainForm.Controls.Add($ImportHosts)

 #Endregion GUI-Elemente

#Endregion GUI

#Region Zuweisung einplanen
function Zuweisung_einplanen {
    # Skript nur Ausführen, wenn "Zuweisung einplanen" geklickt wurde
    if(($mainForm.ActiveControl.Text -ne "Zuweisung einplanen")){
        exit
    }

    # Variablen, welche Für jeden Host gleich bleiben in der Schleife
    $HostArray = $($Hostnames.Text) -split "`r`n"
    $Jobname = $($JobFeld.Text)
    $Neustart = $($Neustart.Checked)
    $Zeitpunkt = [datetime]::ParseExact($Wann,'ddMMyyyy_HHmmss',$null)
    $Zeitpunkt = Get-Date -Date $Datum.Text -Hour $Uhrzeit.Value.Hour -Minute $Uhrzeit.Value.Minute -Second 0
    $Trigger = New-ScheduledTaskTrigger -Once -At $Zeitpunkt
    $principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $Initiator = whoami
    $InitiatorShort = $Initiator.split("\")[1]
    $Counter = 0
    $ErrorCounter = 0
    $ErrorHosts = @()

    # CimSession aufbauen
    if([string]::IsNullOrEmpty($RemoteFQDN)){
        $CimSession = New-CimSession
    }else{
        $CimSession = New-CimSession -ComputerName $RemoteFQDN -Credential Get-Credential
    }

    # Log-Dateien erstellen und wenns Not tut auch Pfad
    $LogDatei = ($LogPfad+"\"+$LogName).Replace("\\","\")
    $ErrorLogDatei = ($ErrorLogPfad+"\"+$ErrorLogName).Replace("\\","\")
    if(!(Test-Path $LogPfad)){
        mkdir $LogPfad
    }
    if(!(Test-Path $LogDatei)){
        New-Item $LogDatei
    }
    if(!(Test-Path $ErrorLogDatei)){
        New-Item $ErrorLogDatei
    }
    Write-Output "########## NEUE AUSFÜHRUNG ##########" >> $LogDatei
    Write-Output "########## NEUE AUSFÜHRUNG ##########" >> $ErrorLogDatei

    # Variablen, welche sich ändern pro Schleifendurchlauf
    foreach($Hostname in $HostArray){
        $Hostname = $Hostname.Trim()
        if($($Hostname.Length) -ne 0){
            $Aktion = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-windowstyle hidden -ep bypass -noprofile -file 'PFAD_ZU_SKRIPT' -Parameter1 'Wert1' -Switch" 
            $TaskName = ($InitiatorShort+" - "+$Hostname+" - "+$Jobname)

            # Check, ob Task Name bereits vergeben ist und ob der überschrieben werden darf
            if(Get-ScheduledTask -CimSession $CimSession -TaskName $TaskName -ErrorAction SilentlyContinue){
                if($Rueckinfo.Text -eq "Für jeden Host" -OR $Rueckinfo.Text -eq "Fehler + Gesamt"){
                    if(([System.Windows.Forms.MessageBox]::Show("Der Task '$Taskname' ist bereits vorhanden, soll der Eintrag überschrieben werden?","Benutzerabfrage",4)) -eq "Yes"){
                    Unregister-ScheduledTask -CimSession $CimSession -TaskName $TaskName -Confirm:$false
                    }else{
                        $NoTaskForce = $true
                    }
                }else{
                    Unregister-ScheduledTask -CimSession $CimSession -TaskName $TaskName -Confirm:$false
                }
            }

            # Task anlegen
            if(($Hostname.Count) -eq 1 -and ($Jobname.Count) -eq 1){
                Register-ScheduledTask -CimSession $CimSession -Action $Aktion -Trigger $Trigger -Taskpath $Taskpath -TaskName $TaskName -Principal $Principal #-ErrorAction SilentlyContinue
            }

            # Log erstellen und bisher bekannte Infos reinschreiben
            $Log = @{
                "Erstellt am:" = $(Get-Date).ToString()
                "Hostname:" = $Hostname 
                "Jobname:" = $Jobname
                "Neustart:" = $Neustart
                "Zeitpunkt:" = $($Zeitpunkt.ToString())
                "Initiator:" = $Initiator
            }

            # Check ob Task wirklich erstellt wurde und Rückinfo generieren
            $Check = Get-ScheduledTask -CimSession $CimSession -TaskPath $TaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
            if(($Check) -and ($NoTaskForce -eq $false)){
                $Log += @{"Task erstellt" = "Erfolgreich"}
                $Counter++
                if($Rueckinfo.Text -eq "Für jeden Host"){
                    [System.Windows.Forms.MessageBox]::Show("Der Task '$Taskname' wurde erfolgreich erstellt und wird am $($Zeitpunkt.DateTime) ausgeführt.","Erfolg!",0)
                }
            }elseif(($Check) -and ($NoTaskForce -eq $true)){
                $Log += @{"Task erstellt" =  "Bereits vorhanden und sollte nicht überschrieben werden"}
                if($Rueckinfo.Text -eq "Für jeden Host"){
                    [System.Windows.Forms.MessageBox]::Show("Der Task wurde auf Ihren Wunsch nicht überschrieben.","Alles beim Alten!",0)
                }
            }else{
                $Log += @{"Task erstellt" = "Fehlgeschlagen"}
                $ErrorLog = $Log
                $ErrorCounter++
                $ErrorHosts += @($Hostname)
                if($Rueckinfo.Text -eq "Für jeden Host" -OR $Rueckinfo.Text -eq "Fehler + Gesamt"){
                    if(([System.Windows.Forms.MessageBox]::Show("Es ist ein Fehler aufgetreten!`r`n`r`nSoll der entsprechende Log-Eintrag geöffnet werden?`r`n`r`n`r`n`r`nWenn im Log alle Variablen gefüllt sind, dann wird das Tool wahrscheinlich nicht als Admin oder mit zu wenigen Rechten ausgeführt","Fehler!",4)) -eq "Yes"){
                        $Log | Out-GridView -Title "Fehlgeschlagener Hostname"
                    }
                }
            }

            # Log schreiben
            Write-Output $Log >> $LogDatei
            Write-Output $ErrorLog >> $ErrorLogDatei

            #Relevante Variablen nullen in jedem Schleifen durchlaufen
            $HostID.Clear()
            $Aktion.Clear()
            $TaskName.Clear()
            $Check.Clear()
            $Log.Clear()
            $ErrorLog.Clear()
            $NoTaskForce = $false            
        }
    }
    #"Am Ende"-Fenster
    if($Rueckinfo.Text -eq "Für jeden Host" -OR $Rueckinfo.Text -eq "Nur am Ende" -OR $Rueckinfo.Text -eq "Fehler + Gesamt"){
        if($ErrorCounter -gt 0){
            [System.Windows.Forms.MessageBox]::Show("$ErrorCounter Fehler aufgetreten!`r`n$Counter Ausführungen war(en) erfolgreich.`r`n`r`n`r`n`r`nWenn im Log alle Variablen gefüllt sind, dann wird das Tool wahrscheinlich nicht als Admin oder mit zu wenigen Rechten ausgeführt?","$ErrorCounter Fehler sind aufgetreten",0)
        }else{
            [System.Windows.Forms.MessageBox]::Show("Alle ($Counter) Tasks wurden erfolgreich angelegt","$Counter Task(s) erstellt",0)
        }
    }
    #Eingegebene Werte nach Zuweisungen nullen
    $Hostnames.Clear()
    $JobFeld.Clear()
    $Datum.ResetText()
    $Uhrzeit.ResetText()
    $ErrorCounter.Clear()

}
#Endregion Zuweisung einplanen

#Region Ausführung starten

if($Silent -OR $S){
    Zuweisung_einplanen
}else{
    [void]$mainForm.ShowDialog()
}
