<#
    .SYNOPSIS
    Veeam Tape Report is a script for Veeam Backup and Replication
    to assist operators to export and import tapes.

    .DESCRIPTION
    Veeam Tape Report is a script for Veeam Backup and Replication
    to assist operators to export and import tapes. This script 
    generate a report with a state of the tape infrastructure
    and the list of tapes to export and the one to import
    into the library.

    .EXAMPLE
    .\VeeamTapeReport.ps1
    Run script from (an elevated) PowerShell console  
  
    .NOTES
    Author: Pierre-Benoit Gerber / Athéo Ingénierie
    Last Updated: January 2025
    Version: 2.0
  
    Requires :
    Veeam Backup and Replication and Tape Library
    Disable IO Module(s) in tape library configuration -> use magazines to export tapes.
    Activate in all tape pools the option "Take tape from Free pool"
#> 


## Autheur : Athéo Ingénierie - Pierre-Benoit Gerber
## Fonctions du script : Génération de la liste des bandes Veeam à externaliser et à intégrer. 
## Date - Version - Commentaire : 15/06/2020 - v1.0 - Version initiale
## Date - Version - Commentaire : 16/06/2020 - v1.1 - Ajout connexion serveur VBR / Optimisations visuelles.
## Date - Version - Commentaire : 18/06/2020 - v1.2 - Affichage en rouge si Job en attente de bandes Free et modification du sujet du mail.
## Date - Version - Commentaire : 29/06/2020 - v1.3 - Demande ajout de bandes vièrges si aucune bande expirée.
## Date - Version - Commentaire : 29/06/2020 - v1.4 - Compte uniquement les bandes qui sont dans la robotique pour le Free
## Date - Version - Commentaire : 26/11/2020 - v1.5 - N'indique plus de réintégrer les bandes protégées par Veeam
## Date - Version - Commentaire : 03/12/2020 - v1.6 - Gestion des multi-robotiques pour liste des drives et liste de bandes à externaliser
## Date - Version - Commentaire : 27/06/2023 - v1.7 - Variables pour définir les seuils warning et critical sur le nombre de bandes dispo
## Date - Version - Commentaire : 12/03/2024 - v1.8 - Meilleurs gestion des robotiques mulitples (pour l'etat des drives et nbre de bande Free par robotique) / Inclusion des mediaset a externaliser plutot que exlusion
## Date - Version - Commentaire : 10/04/2024 - v1.9.2 - Correction de la gestion des bandes à externaliser avec plusieurs robotiques / Mise en forme des bandes à externaliser, à intégrées et des jobs en cours dans des tableaux / Ajout du n° de séquence des bandes à externaliser
## Date - Version - Commentaire : 01/01/2025 - v1.10 - Indique dans quel magasin de la robotique se trouvent les bandes à externaliser. (pour l'instant seuls 2 magasins sont gérés) / Ajout du n° de version du script
## Date - Version - Commentaire : 01/01/2025 - v1.10.1 - Permet d'exclure la gestion des magasins par exemple pour les robotiques qui n'en ont qu'un seul.

## Evolutions à réaliser
# Revoir les seuils d'alerte des bandes dispo (faire simple en comptant le nombre de bandes dans le pool free et le nombre de bandes expirées dans la robotique). Revoir pour fonctionner avec plusieurs robotiques.
# Ne pas ramener les bandes des backup Full expires du weekend precedent si Daily externa lisées. (sinon jouer avec la protection des bandes en imposant 1 semaine de plus que la rétention du backup)
# Expliquer le script, comment créer les jobs Veeam et comment créer les taches planifiées
# Ajouter la ligne de commande Powersell qui créé automatiquement la tache planifiées


## Region User-Variables

# VBR Server (Server Name, FQDN, IP or localhost)
$vbrServer = $env:computername
# Report Title
$rptTitle = "Veeam Tape Report"
# Show VBR Server name in report header
$showVBR = $true
# HTML Report Width (Percent)
$rptWidth = 97
# HTML Table Odd Row color
$oddColor = "#f0f0f0"

# Save HTML output to a file
$ExportFile = $true
# HTML File output path and filename
$ExportFolder = "C:\Scripts\VeeamTapeExport"
$ExportFilename = "VeeamTapeExport.html"

# Library list separated by semi-collon
$Libs = "TS4300"

# Determine tape location in left or right magazine
#Give the library first right magazine slot ID to determine in wich magazine are located tapes (right or left)
#To determine this slot ID: put one tape in the top left slot in the right magazine and execute the powershell command "(Get-VBRTapeMedium -name TAPE_NAME).location.SlotAddress"
#Set to $Null to not determine in wich magazine are located tapes
$LibFirstRightSlotID = "16" #IBM TS4300 3573-TL Base Controller Revision B000
#$LibFirstRightSlotID = $null

# MediaSets to be exported 
#Set strings to be mached in the mediaset name separated by pipe
#Set to $Null to export all mediasets
$MediaSetToExport = "Weekly|Monthly|Quarterly|Yearly"
#$MediaSetToExport = $null

# Free tapes threshold
#Set Warning and Critical threshold for available free tapes in Free Media Pool
#Basicaly you must have at least necessary free tapes count for full backup
$WarningFreeTapeCount = "9"
$CriticalFreeTapeCount = "7"

# Email configuration
#Send Email - $true or $false
$sendEmail = $false
$emailHost = "smtprelay.yourdomain.local"
$emailUser = ""
$emailPass = ""
$emailFrom = "VeeamTapeExport@yourdomain.local"
$emailTo = "you@yourdomain.local"
#Send report as attachment - $true or $false
$emailAttach = $false


## Script begining

# Script version
$Version = "v2.0"

# VBR Server Connexion
$OpenConnection = (Get-VBRServerSession).Server
If ($OpenConnection -ne $vbrServer){
  Disconnect-VBRServer
  Try {
    Connect-VBRServer -server $vbrServer -ErrorAction Stop
  } Catch {
    Write-Host "Unable to connect to VBR server - $vbrServer" -ForegroundColor Red
    exit
  }
}

# Toggle VBR Server name in report header
If ($showVBR) {
  $vbrName = "VBR Server - $vbrServer"
} Else {
  $vbrName = $null
}

# Tape drives state
$TABDriveState = @()
Foreach ($Lib in $Libs){
    $LibId = (Get-VBRTapeLibrary -Name $Lib).id
    Foreach ($Drive in Get-VBRTapeDrive -Library $LibId ){
        $DriveState = Get-VBRTapeDrive -Name $Drive | Where-Object {$_.LibraryId -Match $LibId} | Select Name,SerialNumber,Enabled,State,Medium
        $OBJDrive = New-Object System.Object
        $OBJDrive | Add-Member -type NoteProperty -Name "Library" -Value $Lib
        $OBJDrive | Add-Member -type NoteProperty -Name "Drive" -Value $DriveState.Name
        $OBJDrive | Add-Member -type NoteProperty -Name "Serial Number" -Value $DriveState.SerialNumber
        $OBJDrive | Add-Member -type NoteProperty -Name "Enabled" -Value $DriveState.Enabled
        $OBJDrive | Add-Member -type NoteProperty -Name "State" -Value $DriveState.State
        $OBJDrive | Add-Member -type NoteProperty -Name "Loaded tape" -Value $DriveState.Medium
        $TABDriveState += $OBJDrive
        If ($DriveState.Enabled -ne $true) {$DriveDisabled = $true}
    }  
}

# Online free tape count
$FreePoolId = (Get-VBRTapeMediaPool -Name Free).id
$TABFreeTapeCount = @()
Foreach ($Lib in $Libs){
    $LibId = (Get-VBRTapeLibrary -Name $Lib).id
    $FreeTapeCount = (Get-VBRTapeMedium | Where-Object {$_.MediaPoolId -match $FreePoolId -and ($_.LibraryId -match $LibId) -and ($_.Location -Match "Slot" -or $_.Location -Match "Drive")}).count
    $OBJFreeTape = New-Object System.Object
    $OBJFreeTape | Add-Member -type NoteProperty -Name "Library" -Value $Lib
    $OBJFreeTape | Add-Member -type NoteProperty -Name "Free Tapes" -Value $FreeTapeCount
    $TABFreeTapeCount += $OBJFreeTape
    If ($FreeTapeCount -le $CriticalFreeTapeCount) {
        $FreeTapeCritical = $true
    } Else { 
        If ($FreeTapeCount -le $WarningFreeTapeCount) {
        $FreeTapeWarning = $true
           }
    }
}

# Waiting Tape Job list
$TABWaitingTapeJobs = @()
$WaitingTapeJobs = Get-VBRTapeJob | where {$_.LastState -eq "WaitingTape"}
Foreach ($WaitingTapeJob in $WaitingTapeJobs){
        $OBJWaitingTapeJob = New-Object System.Object
        $OBJWaitingTapeJob | Add-Member -type NoteProperty -Name "Tape Job " -Value $WaitingTapeJob.Name
        #$OBJWaitingTapeJob | Add-Member -type NoteProperty -Name "Destination tape pool" -Value $WaitingTapeJob.FullBackupMediaPool
        $OBJWaitingTapeJob | Add-Member -type NoteProperty -Name "Destination tape pool" -Value $WaitingTapeJob.Target
        $OBJWaitingTapeJob | Add-Member -type NoteProperty -Name "State" -Value $WaitingTapeJob.LastState
        $TABWaitingTapeJobs += $OBJWaitingTapeJob
}

# Running Tape Job List
$TABWorkgingTapeJobs = @()
$WorkgingTapeJobs = Get-VBRTapeJob | where {$_.LastState -eq "Working"}
Foreach ($WorkgingTapeJob in $WorkgingTapeJobs){
        $OBJWorkgingTapeJob = New-Object System.Object
        $OBJWorkgingTapeJob | Add-Member -type NoteProperty -Name "Tape Job " -Value $WorkgingTapeJob.Name
        $OBJWorkgingTapeJob | Add-Member -type NoteProperty -Name "Destination tape pool" -Value $WorkgingTapeJob.Target
        $OBJWorkgingTapeJob | Add-Member -type NoteProperty -Name "State" -Value $WorkgingTapeJob.LastState
        $TABWorkgingTapeJobs += $OBJWorkgingTapeJob
}


# Tape list to be exported
$TABTapeExport = @()
Foreach ($Lib in $Libs){
    $LibId = (Get-VBRTapeLibrary -Name $Lib).id
    $TapeExport = Get-VBRTapeMedium | Where-Object {$_.Location -Match "Slot" -and $_.LibraryId -Match $LibID -and $_.MediaSet -match $MediaSetToExport -and $_.MediaSet -ne $Null}
    Foreach ($Tape in $TapeExport){    
        $OBJTapeExport = New-Object System.Object
        $OBJTapeExport | Add-Member -type NoteProperty -Name "Library" -Value $Lib
        $OBJTapeExport | Add-Member -type NoteProperty -Name "Tape Name" -Value $Tape.Name
        $OBJTapeExport | Add-Member -type NoteProperty -Name "Media Set" -Value $Tape.MediaSet
        $OBJTapeExport | Add-Member -type NoteProperty -Name "Sequence Number" -Value $Tape.SequenceNumber
        $OBJTapeExport | Add-Member -type NoteProperty -Name "Expiration Date" -Value $Tape.ExpirationDate
        If ($LibFirstRightSlotID ){
            If ((Get-VBRTapeMedium -name $Tape).location.SlotAddress -lt $LibFirstRightSlotID) {$OBJTapeEXport | Add-Member -type NoteProperty -Name "Magazine" -Value "Left"}
            If ((Get-VBRTapeMedium -name $Tape).location.SlotAddress -ge $LibFirstRightSlotID) {$OBJTapeEXport | Add-Member -type NoteProperty -Name "Magazine" -Value "Right"}
        }
        $TABTapeExport += $OBJTapeExport    }
}

# Tape list that could be imported into the library
$TABTapeImport = @()
$TapeImport = Get-VBRTapeMedium | Where-Object {$_.Location -Match "Vault" -and $_.IsExpired -Match "True" -and $_.ProtectedBySoftware -Match "False"}
Foreach ($Tape in $TapeImport){
        $OBJTapeImport = New-Object System.Object
        $OBJTapeImport | Add-Member -type NoteProperty -Name "Tape Name" -Value $Tape.Name
        $OBJTapeImport | Add-Member -type NoteProperty -Name "Media Set" -Value $Tape.MediaSet
        $OBJTapeImport | Add-Member -type NoteProperty -Name "Expirated" -Value $Tape.IsExpired
        $OBJTapeImport | Add-Member -type NoteProperty -Name "Expiration Date" -Value $Tape.ExpirationDate
        $TABTapeImport += $OBJTapeImport
}

## HTML formatting

$HtmlHeaderObj = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>$rptTitle</title>
            <style type="text/css">
              body {font-family: Tahoma; background-color: #ffffff;}
              table {font-family: Tahoma; width: $($rptWidth)%; font-size: 12px; border-collapse: collapse; margin-left: auto; margin-right: auto;}
              table tr:nth-child(odd) td {background: $oddColor;}
              th {background-color: #e2e2e2; border: 1px solid #a7a9ac;border-bottom: none;}
              td {background-color: #ffffff; border: 1px solid #a7a9ac;padding: 2px 3px 2px 3px;}
            </style>
    </head>
"@

$HtmlBodyTop = @"
    <body>
          <table>
              <tr>
                  <td style="width: 50%;height: 14px;border: none;background-color: #00b050;color: White;font-size: 10px;vertical-align: bottom;text-align: left;padding: 2px 0px 0px 5px;"></td>
                  <td style="width: 50%;height: 14px;border: none;background-color: #00b050;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 2px 5px 0px 0px;">Report generated on $(Get-Date -format g)</td>
              </tr>
              <tr>
                  <td style="width: 50%;height: 24px;border: none;background-color: #00b050;color: White;font-size: 24px;vertical-align: bottom;text-align: left;padding: 0px 0px 0px 15px;">$rptTitle</td>
                  <td style="width: 50%;height: 24px;border: none;background-color: #00b050;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 2px 0px;">$vbrName</td>
              </tr>
          </table>
"@

$HtmlSubHead01 = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #f3f4f4;color: #626365;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$HtmlSubHead01suc = @"
<table>
                 <tr>
                    <td style="height: 35px;background-color: #00b050;color: #626365;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$HtmlSubHead01war = @"
<table>
                 <tr>
                    <td style="height: 35px;background-color: #ffd96c;color: #626365;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$HtmlSubHead01err = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #FB9895;color: #626365;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$HtmlSubHead01inf = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #3399FF;color: #626365;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$HtmlSubHead02 = @"
</td>
                </tr>
             </table>
"@

$HTMLbreak = @"
<table>
                <tr>
                    <td style="height: 10px;background-color: #626365;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;"></td>
						    </tr>
            </table>
"@

$HtmlFooterObj = @"
            <table>
                <tr>
                    <td style="height: 15px;background-color: #ffffff;border: none;color: #626365;font-size: 10px;text-align:center;">My Veeam Report developed by <a href="http://blog.smasterson.com" target="_blank">http://blog.smasterson.com</a> and modified for V11 by <a href="http://horstmann.in" target="_blank">http://horstmann.in</a></td>
                </tr>
            </table>
    </body>
</html>
"@

$HtmlFooterObj = @"
            <table>
                <tr>
                    <td style="height: 15px;background-color: #ffffff;border: none;color: #626365;font-size: 10px;text-align:center;">Veeam Tape Report developed by <a href="http://blog.smasterson.com" target="_blank">http://blog.smasterson.com</a> and modified for V11 by <a href="http://horstmann.in" target="_blank">http://horstmann.in</a></td>
                </tr>
            </table>
    </body>
</html>
"@


If ($FreeTapeCritical) {
    $HtmlFreeTapeCount = $HtmlSubHead01err + "Available Free tapes in each Library:" + ($TABFreeTapeCount | ConvertTo-Html -Fragment) + $HtmlSubHead02
}
Else{
    If ($FreeTapeWarning) {
    $HtmlFreeTapeCount = $HtmlSubHead01war + "Available Free tapes in each Library:" + ($TABFreeTapeCount | ConvertTo-Html -Fragment) + $HtmlSubHead02
    }
    Else {$HtmlFreeTapeCount = $HtmlSubHead01 + "Available Free tapes in each Library:" + ($TABFreeTapeCount | ConvertTo-Html -Fragment) + $HtmlSubHead02}
}

If ($DriveDisabled){
    $HtmlDriveState = $HtmlSubHead01err + "Tape drive state error" + $HtmlDriveState + ($TABDriveState | ConvertTo-Html -Fragment) + $HtmlSubHead02
}
Else {
    $HtmlDriveState = $HtmlSubHead01 + "Tape drive state" + ($TABDriveState | ConvertTo-Html -Fragment) + $HtmlSubHead02
}

If ($TABWaitingTapeJobs){ 
    $HtmlWaitingTapeJob = $HtmlSubHead01err + "Jobs waiting for free tapes" + ($TABWaitingTapeJobs| ConvertTo-Html -Fragment) + $HtmlSubHead02
}
Else {
    If ($TABWorkingTapeJobs){
        $HtmlWorkingTapeJob = $HtmlSubHead01warn + "Running Tape Jobs" + ($TABWorkingTapeJobs| ConvertTo-Html -Fragment) + $HtmlSubHead02
    }
    Else {$HtmlWorkgingTapeJob = $HtmlSubHead01 + "No running Tape Job" + $HtmlSubHead02}
}


If ($TABTapeExport){ 
        $HtmlTapeExport = $HtmlSubHead01 + "Tape(s) to export:" + ($TABTapeExport | ConvertTo-Html -Fragment) + $HtmlSubHead02
}
Else {$HtmlTapeExport = $HtmlSubHead01 + "No tape to export" + $HtmlSubHead02}

If ($TABTapeImport){ 
        $HtmlTapeImport = $HtmlSubHead01 + "Tape(s) that can be imported" + ($TABTapeImport | ConvertTo-Html -Fragment) + $HtmlSubHead02
}
Else {$HtmlTapeImport = $HtmlSubHead01 + "No tape to import - Consider importing new Free tapes" + $HtmlSubHead02}

$HtmlOutput = $HtmlHeaderObj + $HtmlBodyTop + $HtmlFreeTapeCount + $HtmlDriveState + $HtmlWaitingTapeJob + $HtmlWorkingTapeJob + $HtmlTapeExport + $HtmlTapeImport + $HtmlFooterObj

#HTML formatting optimisations
#$HtmlOutput = $HtmlOutput.Replace("<table>","<table border=1 cellspacing=0 cellpadding=10>")
#$HtmlOutput = $HtmlOutput.Replace("<tr><th>","<tr bgcolor='Silver'><th>")

#HTML file export
if ($ExportFile){
    $HtmlOutput | Out-File -FilePath $ExportFolder\$(get-date -f yyyy-MM-dd.HH-mm-s)_$ExportFilename
}

#Envoi du mail
If ($sendEmail) {
  $smtp = New-Object System.Net.Mail.SmtpClient($emailHost, $emailPort)
  $smtp.Credentials = New-Object System.Net.NetworkCredential($emailUser, $emailPass)
  $smtp.EnableSsl = $emailEnableSSL
  $msg = New-Object System.Net.Mail.MailMessage($emailFrom, $emailTo)
  $msg.Subject = $emailSubject
  #$body = $HtmlOutput
  $msg.Body = $HtmlOutput
  $msg.isBodyhtml = $true
  $smtp.send($msg)
} 
