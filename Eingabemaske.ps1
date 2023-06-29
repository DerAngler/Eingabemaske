###############################################################
##### Diese Variablen können nach Bedarf angepasst werden #####
###############################################################

#Log definieren
$LogPfad = "C:\Logs\Github\Eingabemaske\"
$LogName = $(Get-Date -Format "yyyyMMdd") + "_Eingabemaske.log"

#Cim-Session Endpoint (Bei Planlosigkeit einfach so lassen)
$RemoteFQDN = $null #$null oder leer = localhost 

###############################################################
################### Ab hier nix mehr ändern ###################
###############################################################

#Region Eingabemaske
Add-Type -AssemblyName System.Windows.Forms

#Main Form
$mainForm = New-Object System.Windows.Forms.Form
$Font = New-Object System.Drawing.Font("Courier New", 12)
$mainForm.Text = "Eingabemaske"
$mainForm.Font = $Font
$mainForm.ForeColor = "White"
$mainForm.BackColor = "Black"
$mainForm.Height = 382
$mainForm.Width = 280
$mainForm.MaximizeBox = $false
$mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$mainForm.StartPosition = "CenterScreen"
$mainForm.ShowIcon = $false

#Jobname Label
$JobnameLabel = New-Object System.Windows.Forms.Label
$JobnameLabel.Text = "Jobname"
$JobnameLabel.Location = "5, 10"
$JobnameLabel.Height = 22
$JobnameLabel.Width = 90
$mainForm.Controls.Add($JobnameLabel)

#Hostnames Label
$HostnamesLabel = New-Object System.Windows.Forms.Label
$HostnamesLabel.Text = "Hostnames"
$HostnamesLabel.Location = "5, 45"
$HostnamesLabel.Height = 22
$HostnamesLabel.Width = 95
$mainForm.Controls.Add($HostnamesLabel)

#Datum Label
$DatumLabel = New-Object System.Windows.Forms.Label
$DatumLabel.Text = "Datum"
$DatumLabel.Location = "5, 180"
$DatumLabel.Height = 22
$DatumLabel.Width = 90
$mainForm.Controls.Add($DatumLabel)

#Uhrzeit Label
$UhrzeitLabel = New-Object System.Windows.Forms.Label
$UhrzeitLabel.Text = "Uhrzeit"
$UhrzeitLabel.Location = "5, 215"
$UhrzeitLabel.Height = 22
$UhrzeitLabel.Width = 90
$mainForm.Controls.Add($UhrzeitLabel)

#Neustart Label
$NeustartLabel = New-Object System.Windows.Forms.Label
$NeustartLabel.Text = "Job-Neustart erlauben?"
$NeustartLabel.Location = "5, 250"
$NeustartLabel.Height = 22
$NeustartLabel.Width = 230
$mainForm.Controls.Add($NeustartLabel)

#Jobname Auswahlfeld
$Job = New-Object System.Windows.Forms.Textbox
$Job.Location = "100, 7"
$Job.Width = "155"
$mainForm.Controls.Add($Job)

# Hostname Auswahlfeld
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
$mainForm.Controls.Add($Uhrzeit)

# Neustart CheckBox
$Neustart = New-Object System.Windows.Forms.CheckBox
$Neustart.Location = "242, 248"
$Neustart.Checked = $TRUE
$mainForm.Controls.Add($Neustart)

# Zuweisen Button
$Zuweisen = New-Object System.Windows.Forms.Button
$Zuweisen.Location = "5, 282"
$Zuweisen.Width = "252"
$Zuweisen.ForeColor = "Black"
$Zuweisen.BackColor = "White"
$Zuweisen.Text = "Zuweisung einplanen"
$Zuweisen.add_Click({Zuweisung_einplanen})
$mainForm.Controls.Add($Zuweisen)

#Task-Button
$TasksButton = New-Object System.Windows.Forms.Button
$TasksButton.Location = "5, 310"
$TasksButton.Width = "124"
$TasksButton.ForeColor = "Black"
$TasksButton.BackColor = "White"
$TasksButton.Text = "Tasks"
$TasksButton.add_Click({Get-ScheduledTask -TaskPath "\Microsoft\Office\" | Out-GridView -OutputMode Multiple -Title Bara-Job-Scheduler | Get-ScheduledTaskInfo | Out-GridView -Title Bara-Job-Scheduler -PassThru})
$mainForm.Controls.Add($TasksButton)

#Log-Button
$LogsButton = New-Object System.Windows.Forms.Button
$LogsButton.Location = "133, 310"
$LogsButton.Width = "124"
$LogsButton.ForeColor = "Black"
$LogsButton.BackColor = "White"
$LogsButton.Text = "Logs"
$LogsButton.add_Click({(explorer.exe $LogPfad)})
$mainForm.Controls.Add($LogsButton)
#Endregion Eingabemaske

function Zuweisung_einplanen{
    #Skript nur Ausführen, wenn "Zuweisung einplanen" geklickt wurde
    if(($mainForm.ActiveControl.Text -ne "Zuweisung einplanen")){
        exit
    }

    #Variablen, welche immer gleich bleiben in der Schleife
    $HostArray = $($Hostnames.Text) -split "`r`n"
    $Jobname = $($Job.Text)
    $Neustart = $($Neustart.Checked)
    $Zeitpunkt = Get-Date -Date $Datum.Text -Hour $Uhrzeit.Value.Hour -Minute $Uhrzeit.Value.Minute -Second 0
    $Trigger = New-ScheduledTaskTrigger -Once -At $Zeitpunkt
    $Principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $Initiator = whoami

    #CimSession aufbauen
    if([string]::IsNullOrEmpty($RemoteFQDN)){
        $CimSession = New-CimSession
    }else{
        $CimSession = New-CimSession -ComputerName $RemoteFQDN -Credential Get-Credential
    }

    #Log-Datei erstellen und wenns Not tut auch Pfad
    $LogDatei = ($LogPfad+"\"+$LogName).Replace("\\","\")
    if(!(Test-Path $LogPfad)){
        New-Item -Path $LogPfad -ItemType Directory
    }
    if(!(Test-Path $LogDatei)){
        New-Item -Path $LogDatei -ItemType File
    }

    #Variablen, welche sich ändern pro Schleifendurchlauf
    foreach($Hostname in $HostArray){
        $Hostname = $Hostname.Trim()
        if($($Hostname.Length) -ne 0){
            $Aktion = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ep bypass -noprofile -file ... -Parameter Wert ..." 
            $TaskName = ($Hostname+" - "+$Jobname+" - "+$Initiator.Replace("\","_")) #Das "\" muss ersetzt werden, da der String sonst als Pfad erkannt wird und ungewollte Ordner erstellt werden in der Aufgabenplanung

            #Check, ob Task Name bereits vergeben ist und ob der überschrieben werden darf
            if(Get-ScheduledTask -CimSession $CimSession -TaskName $TaskName -ErrorAction SilentlyContinue){
                if(([System.Windows.Forms.MessageBox]::Show("Der Task '$Taskname' ist bereits vorhanden, soll der Eintrag überschrieben werden?","Benutzerabfrage",4)) -eq "Yes"){
                Unregister-ScheduledTask -CimSession $CimSession -TaskName $TaskName -Confirm:$false
                $TaskForce = $true
                }else{
                    $TaskForce = $false
                }
            }else{
                $TaskForce = $true
            }

            #Task anlegen
            if(($HostID.Count) -eq 1 -and ($JobID.Count) -eq 1){
                Register-ScheduledTask -CimSession $CimSession -Action $Aktion -Trigger $Trigger -Taskpath "Barajob-Scheduler" -TaskName $TaskName -Principal $Principal -ErrorAction SilentlyContinue
            }

            #Log-Infos abfragen
            $Log = @{
                "Erstellt am:" = $(Get-Date).ToString()
                "Hostname:" = $Hostname 
                "Jobname:" = $Jobname
                "Neustart:" = $Neustart
                "Zeitpunkt:" = $($Zeitpunkt.ToString())
                "Initiator:" = $Initiator
            }

            #Check ob Task wirklich erstellt wurde
            $Check = Get-ScheduledTask -CimSession $CimSession -TaskName $TaskName -ErrorAction SilentlyContinue
            if(($Check) -and ($TaskForce -eq $true)){
                [System.Windows.Forms.MessageBox]::Show("Der Task '$Taskname' wurde erfolgreich erstellt und wird am $($Zeitpunkt.DateTime) ausgeführt.","Erfolg!",0)
                $Log += @{"Task erstellt" = "Erfolgreich"}
            }
            elseif(($Check) -and ($TaskForce -eq $false)){
                [System.Windows.Forms.MessageBox]::Show("Der Task wurde auf Ihren Wunsch nicht überschrieben.","Alles beim Alten!",0)
                $Log += @{"Task erstellt" =  "Bereits vorhanden und sollte nicht überschrieben werden"}
            }else{
                if(([System.Windows.Forms.MessageBox]::Show("Es ist ein Fehler aufgetreten!`r`nBitte Eingaben prüfen und erneut versuchen.`r`nSoll das aktuelle Log angezeigt werden?","Fehler!",4)) -eq "Yes"){
                    $Log | Out-GridView
                }
                $Log += @{"Task erstellt" = "Fehlgeschlagen"}
            }

            #Log schreiben und relevante Variablen nullen
            Write-Output $Log >> $LogDatei
            $HostID = ""
            $Aktion = ""
            $TaskName = ""
            $Check = ""
        }
    }
}

# Eingabemaske darstellen
[void]$mainForm.ShowDialog()