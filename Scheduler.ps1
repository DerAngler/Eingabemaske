#Region Variablen und Parameter
###############################################################
######## Diese Parameter-Werte können angepasst werden ########
###############################################################
param(
    [string]$Job = "",
    [array]$Hosts = @(""),
    [string]$Wann = "",
    [string]$Execute = "powershell.exe",
    [string]$Argument = "-windowstyle hidden -ep bypass -noprofile -file 'PFAD_ZU_SKRIPT' -Parameter1 'Wert1' -Switch",
    [switch]$NoJobNeustart = $false,
    [switch]$S = $false,
    [switch]$Silent = $false,
    [switch]$Darkmode = $false,
    [switch]$NoTaskForce = $false,
    [string]$LogPfad = "C:\Logs\Github\Scheduler\",
    [string]$LogName = $(Get-Date -Format "yyyyMMdd") + "_Scheduler.log",
    [switch]$Admin = $false,
    [string]$ErrorLogPfad = $LogPfad,
    [string]$ErrorLogName = $(Get-Date -Format "yyyyMMdd") + "_ERROR_Scheduler.log",
    [string]$RemoteFQDN = $null,
    [string]$TaskPath = "\EigenerScheduler\",
    [string]$TempLogFile = "C:\Temp\" + $(Get-Date -Format "yyyyMMdd") + "_fehlgeschlagene_Clients.log",
    [switch]$Switch = $false,
    [string]$Spalter = ","
)
$ErrorActionPreference = "SilentlyContinue"
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

function Tasks_anzeigen {
    if(Get-ScheduledTask -TaskPath $TaskPath){
        if($Admin -ne $TRUE){
            Get-ScheduledTask -TaskPath $TaskPath | Get-ScheduledTaskInfo | Out-GridView -Title "Scheduler"
        }else{
            Get-ScheduledTask -TaskPath $TaskPath | Get-ScheduledTaskInfo | Out-GridView -Title "Scheduler - Markierte Tasks werden mit 'OK' gelöscht!!!" -PassThru | Unregister-ScheduledTask -Confirm:$false
        }
    }else{
        [System.Windows.Forms.MessageBox]::Show("Es konnten keine Tasks unter '$TaskPath' gefunden werden`r`n`r`nBitte das Skript mit einem anderen User oder als Admin ausführen`r`nWenn diese Meldung dennoch erscheint, gibt es wohl keine Tasks","Keine Tasks gefunden",0)
    }
}

# Switch-Button soll $Jobname und $Hosts tauschen, sodass $Hosts = JobNabemArray ist und $Jobname = ein Host
function Switch_Variablen {
    param (
        $Jobname,
        $HostArray
    )
    $Temp_Switch = $HostArray
    $HostArray = $Jobname
    $Jobname = $Temp_Switch
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
$HostnamesLabel.Location = "5, 46"
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
$Neustart.Location = "241, 248"
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
$Zuweisen.Location = "5, 313"
$Zuweisen.Height = "32"
$Zuweisen.Width = "252"
$Zuweisen.Text = "Zuweisung einplanen"
$Zuweisen.BackColor = "White"
$Zuweisen.ForeColor = "Black"
$Zuweisen.add_Click({Zuweisung_einplanen})
$mainForm.Controls.Add($Zuweisen)

# Task-Button
$TasksButton = New-Object System.Windows.Forms.Button
$TasksButton.Location = "5, 349"
$TasksButton.Width = "124"
$TasksButton.Text = "Tasks"
$TasksButton.BackColor = "White"
$TasksButton.ForeColor = "Black"
$TasksButton.add_Click({Tasks_anzeigen})
$mainForm.Controls.Add($TasksButton)

# Logs-Button
$LogsButton = New-Object System.Windows.Forms.Button
$LogsButton.Location = "133, 349"
$LogsButton.Width = "124"
$LogsButton.Text = "Logs"
$LogsButton.BackColor = "White"
$LogsButton.ForeColor = "Black"
$LogsButton.add_Click({(explorer.exe $LogPfad)})
$mainForm.Controls.Add($LogsButton)

# Import Hostnames from File - Button
$ImportHosts = New-Object System.Windows.Forms.Button
$ImportHosts.Location = "17, 80"
$ImportHosts.Height = "20"
$ImportHosts.Width = "65"
$ImportHosts.Font = ("Courier New, 10")
$ImportHosts.Text = "Import"
$ImportHosts.BackColor = "White"
$ImportHosts.ForeColor = "Black"
$ImportHosts.add_Click({Dateiinhalt_per_OpenFileDialog_in_ne_CheckBox_schreiben($Hostnames)})
$mainForm.Controls.Add($ImportHosts)

# Switch CheckBox
$SwitchCB = New-Object System.Windows.Forms.CheckBox
$SwitchCB.Location = "8, 27"
$SwitchCB.Font = ("Courier New, 10")
$SwitchCB.Text = "Switch"
$SwitchCB.Add_CheckStateChanged({
    if($SwitchCB.checked){
        $HostnamesLabel.Text = "Jobnames"
        $JobnameLabel.Text = "Hostname"
        $Hostnames.Width = "1900"
        $mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $Switch = $true
    } else {
        $HostnamesLabel.Text = "Hostnames"
        $JobnameLabel.Text = "Jobname"
        $Hostnames.Width = "155"
        $mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
        $mainForm.Height = 420
        $mainForm.Width = 282
        $Switch = $false
   }  
})
$mainForm.Controls.Add($SwitchCB)

 #Endregion GUI-Elemente

#Endregion GUI

#Region Zuweisung einplanen
function Zuweisung_einplanen {
    #Switch-Button
    if($Switch -eq $true){
        Switch_Variablen($Jobname, $Hosts)
    }
    # Skript nur Ausführen, wenn "Zuweisung einplanen" geklickt oder einer der Silent-Parameter genutzt wurde
    # Zudem müssem manche Variablen je nach Silent oder GUI Ausführung anders aufgebaut werden
    if(!($Silent -OR $S)){
        if(($mainForm.ActiveControl.Text -ne "Zuweisung einplanen")){#Cancle-Buttob tatsächlich canclen lassen
            exit
        }else{#Wenn GUI
            $HostArray = $($Hostnames.Text) -split "`r`n"
            $Jobname = $($JobFeld.Text).Trim()
            $Zeitpunkt = $(Get-Date -Date $Datum.Text -Hour $Uhrzeit.Value.Hour -Minute $Uhrzeit.Value.Minute -Second 0)
        }
    }else{#Wenn Silent
        $HostArray = ($Hosts -split "$Spalter")
        $Jobname = $Job
        $Zeitpunkt = $(Get-Date -Date $Wann)
        $Rueckinfo.SelectedIndex = 4
    }

    # Variablen, die für jeden Host gleich bleiben in der Schleife
    $Neustart = $($Neustart.Checked)
    $Trigger = New-ScheduledTaskTrigger -Once -At $Zeitpunkt
    $principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $Initiator = whoami
    $InitiatorShort = $Initiator.split("\")[1]
    $Counter = 0
    $ErrorCounter = 0
    $SkippedCounter = 0
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

    # Check, ob Eingaben / Parameter valide sind
    if(!($Silent -eq $true -or $S -eq $true -or $Rueckinfo.Task -eq "Nur am Ende")){
        if([string]::IsNullOrEmpty($HostArray)){
            [System.Windows.Forms.MessageBox]::Show("Es wurden keine Hostnamen angegeben","Hostname fehlt",0)
            $InvalideWerte = $true
        }
        if([string]::IsNullOrEmpty($Jobname)){
            [System.Windows.Forms.MessageBox]::Show("Es wurde kein Jobname angegeben","Jobname fehlt",0)
            $InvalideWerte = $true
            $HostArray = ""
        }
        if((Get-Date) -gt $Zeitpunkt){
            [System.Windows.Forms.MessageBox]::Show("Der angegebene Zeitpunkt liegt in der Vergangenheit.`r`nBitte wählen Sie einen Zeitpunkt aus der Zukunft aus.","Inakzeptabler Zeitpunkt",0)
            $InvalideWerte = $true
            $HostArray = ""
        }
    }

    Write-Output "########## $(Get-Date -Format "hh:mm:ss") - NEUE AUSFÜHRUNG ##########" >> $LogDatei
    
    # Variablen, die sich ändern pro Schleifendurchlauf
    foreach($Hostname in $HostArray){
        $Hostname = $Hostname.Trim()
        if($($Hostname.Length) -ne 0){
            $Aktion = New-ScheduledTaskAction -Execute $Execute -Argument $Argument 
            $TaskName_Switched = ($InitiatorShort+" - "+$Jobname+" - "+$Hostname)
            $TaskName_Unswitched = ($InitiatorShort+" - "+$Hostname+" - "+$Jobname)
            if($Switch -eq $true){
                $TaskName = $TaskName_Switched
            }else{
                $TaskName = $TaskName_Unswitched
            }

            # Check, ob Task Name bereits vorhanden ist und ob der überschrieben werden darf unter Berücksichtigung vom Switch-Button
            if(Get-ScheduledTask -CimSession $CimSession -TaskName $TaskName_Unswitched -ErrorAction SilentlyContinue){
                if($($Rueckinfo.Text) -eq "Für jeden Host" -OR $($Rueckinfo.Text) -eq "Fehler + Gesamt"){
                    if(([System.Windows.Forms.MessageBox]::Show("Der Task '$Taskname_Unswitched' ist bereits vorhanden, soll der Eintrag überschrieben werden?","Benutzerabfrage",4)) -eq "Yes"){
                        Unregister-ScheduledTask -CimSession $CimSession -TaskName $Taskname_Unswitched -Confirm:$false
                    }
                }
                if($NoTaskForce -eq $false){
                    Unregister-ScheduledTask -CimSession $CimSession -TaskName $TaskName_Unswitched -Confirm:$false
                }
            }elseif(Get-ScheduledTask -CimSession $CimSession -TaskName $TaskName_Switched -ErrorAction SilentlyContinue){
                if($($Rueckinfo.Text) -eq "Für jeden Host" -OR $($Rueckinfo.Text) -eq "Fehler + Gesamt"){
                    if(([System.Windows.Forms.MessageBox]::Show("Der Task '$Taskname_Switched' ist bereits vorhanden, soll der Eintrag überschrieben werden?","Benutzerabfrage",4)) -eq "Yes"){
                        Unregister-ScheduledTask -CimSession $CimSession -TaskName $Taskname_Switched -Confirm:$false
                    }
                }
                if($NoTaskForce -eq $false){
                    Unregister-ScheduledTask -CimSession $CimSession -TaskName $TaskName_Switched -Confirm:$false
                }
            }

            # Task anlegen
            if(($Hostname.Count) -eq 1 -AND ($Jobname.Count) -eq 1 -AND ($InvalideWerte -ne $true)){
                Register-ScheduledTask -CimSession $CimSession -Action $Aktion -Trigger $Trigger -Taskpath $Taskpath -TaskName $TaskName -Principal $Principal #-ErrorAction SilentlyContinue
            }

            # Log erstellen und bisher bekannte Infos reinschreiben
            $Log = @{
                "Erstellt am:" = $(Get-Date).ToString()
                "Hostname:" = $Hostname 
                "Jobname:" = $Jobname
                "JobNeustart:" = $Neustart
                "Zuweisungs-Zeitpunkt:" = $($Zeitpunkt.ToString())
                "Valide Werte erhalten" = $((!$InvalideWerte))
                "Initiator:" = $Initiator
                "Job und Hosts getauscht:" = $Switch
                "TaskForce" = $(!($NoTaskForce))
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
                    $SkippedCounter++
                }
            }else{
                $Log += @{"Task erstellt" = "Fehlgeschlagen"}
                Write-Output "########## $(Get-Date -Format "hh:mm:ss") - NEUE AUSFÜHRUNG ##########" >> $ErrorLogDatei
                $ErrorLog = $Log
                $ErrorCounter++
                $ErrorHosts += @($Hostname)
                if($Rueckinfo.Text -eq "Für jeden Host" -OR $Rueckinfo.Text -eq "Fehler + Gesamt"){
                    if(([System.Windows.Forms.MessageBox]::Show("Es ist ein Fehler aufgetreten!`r`n`r`nSoll der entsprechende Log-Eintrag geöffnet werden?`r`n`r`nWenn im Log alle Variablen gefüllt sind, dann wird das Tool wahrscheinlich nicht als Admin oder mit zu wenigen Rechten ausgeführt. Soll das Log des zugehörigen Task geöffnet werden?","Fehler!",4)) -eq "Yes"){
                        $Log | Out-GridView -Title "Fehlgeschlagener Host"
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
        }
    }
    #"Am Ende"-Fenster
    if($Silent -ne $true -AND $S -ne $true){
        if($Rueckinfo.Text -eq "Für jeden Host" -OR $Rueckinfo.Text -eq "Nur am Ende" -OR $Rueckinfo.Text -eq "Fehler + Gesamt"){
            if($ErrorCounter -gt 0){
                if(([System.Windows.Forms.MessageBox]::Show("$ErrorCounter Ausführung ist auf Fehler gelaufen!`r`n$Counter Ausführungen war(en) erfolgreich.`r`n`r`nWenn im Log alle Variablen gefüllt sind, dann wird das Tool wahrscheinlich nicht als Admin oder mit zu wenigen Rechten ausgeführt. Soll das eben gemeinte Log direkt im Notepad geöffnet werden?","$ErrorCounter Fehler sind aufgetreten",4)) -eq "Yes"){
                    notepad.exe $ErrorLogDatei
                }
            }elseif($SkippedCounter -gt 0){
                if(([System.Windows.Forms.MessageBox]::Show("Alle $Counter Task(s) wurden erfolgreich angelegt`r`n$SkippedCounter Task(s) sollten nicht überschrieben werden (/NoTaskForce)`r`n`r`nSoll das Log geöffnet werden?","$Counter Task(s) erstellt",4)) -eq "Yes"){
                    notepad.exe $LogDatei
                }
            }else{
                [System.Windows.Forms.MessageBox]::Show("Alle $Counter Task(s) wurden erfolgreich angelegt","$Counter Task(s) erstellt",0)
            }
        }
    }
    #Eingegebene Werte nach Zuweisungen nullen
    $Counter.Clear()
    $ErrorCounter.Clear()
    $Skipped.Clear()
    if($InvalideWerte -ne $true){
        $Hostnames.Clear()
        $JobFeld.Clear()
        $Datum.ResetText()
        $Uhrzeit.ResetText()
    }
    $InvalideWerte.Clear()
}
#Endregion Zuweisung einplanen

#Region Ausführung starten
if($Silent -eq $true -or $S -eq $true){
    Zuweisung_einplanen
}else{
    [void]$mainForm.ShowDialog()
}
#Endregion Ausführung starten