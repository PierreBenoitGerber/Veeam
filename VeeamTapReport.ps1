<#
    .SYNOPSIS
    Veeam Tape Report is a script for Veeam Backup and Replication
    to assist operators to export and import tapes.

    .DESCRIPTION
    Veeam Tape Report is a script for Veeam Backup and Replication
    to assist operators to export and import tapes. This script 
    generate a report with a state of the tape infrastructure
    and the list of the tapes to exports and the one to import
    into the library.
    Disable IO Module(s) -> use magasines to checkout tapes.
    Activate the tape pool option "Take tape from Free pool"

    .EXAMPLE
    .\MyVeeamReport.ps1
    Run script from (an elevated) PowerShell console  
  
    .NOTES
    Author: Pierre-Benoit Gerber
    Last Updated: January 2025
    Version: 1.10.1
  
    Requires :
    Veeam Backup and Replication
    Tape Library
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

# VBR Server (Server Name, FQDN or IP)
$vbrServer = "localhost"

# Save HTML output to a file
$ExportFile = $true
# HTML File output path and filename
$ExportFolder = "C:\Scripts\VeeamTapeReport"
$ExportFilename = "VeeamTapeReport.htm_$(Get-Date -format yyyyMMdd_hhmmss)"

# Library list seprate by semi-collon
$Libs = "MyLibrary1";"MyLibrary2"

# MediaSets to be exported 
#Set strings to be mached in the mediaset name separated by pipe |
#Set to $Null to export all mediasets
$MediaSetToCheckout = "Weekly|Monthly|Quarterly|Yearly"
#$MediaSetToCheckout = $Null

# Free tapes threshold
$WarningFreeTapeCount = "9"
$CriticalFreeTapeCount = "7"


# Afin de déterminer dans quel magasin se trouve chaque bande il faut renseinger le slot ID retourné par veeam (pas le même que la robotique) qui concerne la première bande du magasin de droite.
# Commenter la variable ou la mettre à $null pour ne pas gérer les magasins.
# 1. Repérer le numéro de la bande qui se trouve dans le slot en bas à gauche du magasin de droite à partir de la console de gestion de la robotique.
# 2. Retrouver l'ID de Slot Veeam avec la commande (Get-VBRTapeMedium -name NOMBANDE).location.SlotAddress
# Ici la valeur 16 pour une robotique IBM TS4300 de première génération dont tous les slots du bas sont inhibés
#$VeeamFirstRightSlotID = $null
$VeeamFirstRightSlotID = "16"

# Configuration SMTP
$sendEmail = $true
$emailHost = "smtprelay.camacte.local"
$emailPort = "25"
$emailEnableSSL = $false
$emailUser = ""
$emailPass = ""
$emailFrom = "veeam@groupe-cam.com"
$emailTo = "serviceit@groupe-cam.com,pierre-benoit.gerber@atheo.net"
$emailSubject = "Gestion des bandes Veeam"

# Script version
$Version = "v1.10.1"

# Connection serveur VBR
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

## Nombre de bandes Online dans le pool Free
$FreePoolId = (Get-VBRTapeMediaPool -Name Free).id
$TABFreeTapeCount = @()
Foreach ($Lib in $Libs){
    $LibId = (Get-VBRTapeLibrary -Name $Lib).id
    $FreeTapeCount = (Get-VBRTapeMedium | Where-Object {$_.MediaPoolId -match $FreePoolId -and ($_.LibraryId -match $LibId) -and ($_.Location -Match "Slot" -or $_.Location -Match "Drive")}).count
    $OBJFreeTape = New-Object System.Object
    $OBJFreeTape | Add-Member -type NoteProperty -Name "Robotique" -Value $Lib
    $OBJFreeTape | Add-Member -type NoteProperty -Name "Bandes Libres" -Value $FreeTapeCount
    $TABFreeTapeCount += $OBJFreeTape
}

## Liste des Jobs Tape en attente de bandes
If (Get-VBRTapeJob | where {$_.LastState -eq "WaitingTape"}){
    $WaitingTapeJob = Get-VBRTapeJob | where {$_.LastState -eq "WaitingTape"} | Select Name,FullBackupMediaPool,LastState
    $emailSubject = "Gestion des bandes Veeam !!! Ajouter des bandes vierges !!!"
}

##Liste des Jobs Tape en cours
$TABWorkgingTapeJobs = @()
$WorkgingTapeJobs = Get-VBRTapeJob | where {$_.LastState -eq "Working"}
Foreach ($WorkgingTapeJob in $WorkgingTapeJobs){
        $OBJWorkgingTapeJob = New-Object System.Object
        $OBJWorkgingTapeJob | Add-Member -type NoteProperty -Name "Nom" -Value $WorkgingTapeJob.Name
        $OBJWorkgingTapeJob | Add-Member -type NoteProperty -Name "Pool de destination" -Value $WorkgingTapeJob.Target
        $OBJWorkgingTapeJob | Add-Member -type NoteProperty -Name "Etat" -Value $WorkgingTapeJob.LastState
        $TABWorkgingTapeJobs += $OBJWorkgingTapeJob
}

## Etat des lecteurs
$TABDriveState = @()
Foreach ($Lib in $Libs){
    $LibId = (Get-VBRTapeLibrary -Name $Lib).id
    Foreach ($Drive in Get-VBRTapeDrive -Library $LibId ){
        $DriveState = Get-VBRTapeDrive -Name $Drive | Where-Object {$_.LibraryId -Match $LibId} | Select Name,SerialNumber,Enabled,State,Medium
        $OBJDrive = New-Object System.Object
        $OBJDrive | Add-Member -type NoteProperty -Name "Robotique" -Value $Lib
        $OBJDrive | Add-Member -type NoteProperty -Name "Lecteur" -Value $DriveState.Name
        $OBJDrive | Add-Member -type NoteProperty -Name "Numéro de série" -Value $DriveState.SerialNumber
        $OBJDrive | Add-Member -type NoteProperty -Name "Activé" -Value $DriveState.Enabled
        $OBJDrive | Add-Member -type NoteProperty -Name "Etat" -Value $DriveState.State
        $OBJDrive | Add-Member -type NoteProperty -Name "Média" -Value $DriveState.Medium
        $TABDriveState += $OBJDrive
    }  
}

## Liste des bandes à externaliser
$TABTapeOut = @()
Foreach ($Lib in $Libs){
    $LibId = (Get-VBRTapeLibrary -Name $Lib).id
    $TapeOut = Get-VBRTapeMedium | Where-Object {$_.Location -Match "Slot" -and $_.LibraryId -Match $LibID -and $_.MediaSet -match $MediaSetToCheckout -and $_.MediaSet -ne $Null}
    Foreach ($Tape in $TapeOut){    
        $OBJTapeOut = New-Object System.Object
        $OBJTapeOut | Add-Member -type NoteProperty -Name "Robotique" -Value $Lib
        $OBJTapeOut | Add-Member -type NoteProperty -Name "Bande" -Value $Tape.Name
        $OBJTapeOut | Add-Member -type NoteProperty -Name "Jeux" -Value $Tape.MediaSet
        $OBJTapeOut | Add-Member -type NoteProperty -Name "N° de séquence" -Value $Tape.SequenceNumber
        $OBJTapeOut | Add-Member -type NoteProperty -Name "Date d'Expiration" -Value $Tape.ExpirationDate
        If ($VeeamFirstRightSlotID ){
            If ((Get-VBRTapeMedium -name $Tape).location.SlotAddress -lt $VeeamFirstRightSlotID) {$OBJTapeOut | Add-Member -type NoteProperty -Name "Magasin" -Value "Left"}
            If ((Get-VBRTapeMedium -name $Tape).location.SlotAddress -ge $VeeamFirstRightSlotID) {$OBJTapeOut | Add-Member -type NoteProperty -Name "Magasin" -Value "Right"}
        }
        $TABTapeOut += $OBJTapeOut
    }
}

#Liste des bandes pouvant être intégrées dans la ou les robotiques
$TABTapeIn = @()
$TapeIn = Get-VBRTapeMedium | Where-Object {$_.Location -Match "Vault" -and $_.IsExpired -Match "True" -and $_.ProtectedBySoftware -Match "False"}
Foreach ($Tape in $TapeIn){
        $OBJTapeIn = New-Object System.Object
        $OBJTapeIn | Add-Member -type NoteProperty -Name "Bande" -Value $Tape.Name
        $OBJTapeIn | Add-Member -type NoteProperty -Name "Jeux" -Value $Tape.MediaSet
        $OBJTapeIn | Add-Member -type NoteProperty -Name "Expirée" -Value $Tape.IsExpired
        $OBJTapeIn | Add-Member -type NoteProperty -Name "Date d'Expiration" -Value $Tape.ExpirationDate
        $TABTapeIn += $OBJTapeIn
}

## Compilation des données et mise en forme HTML

#If ($FreeTapeCount -le $CriticalFreeTapeCount){
#    $HtmlFreeTapeCount = "<p><span style='color:RED'>" + "Nombre de bandes intégrées dans la robotique et dans le pool Free: " + $FreeTapeCount +  "<br />Ajouter des bandes dès que possible!!!</span></p>"
#}
#ElseIf ($FreeTapeCount -gt $WarningFreeTapeCount){
#        $HtmlFreeTapeCount = "<p><span style='color:GREEN'>" + "Nombre de bandes intégrées dans la robotique et dans le pool Free: " + $FreeTapeCount +  "</span></p>"
#}
#Else {$HtmlFreeTapeCount = "<p><span style='color:ORANGE'>" + "Nombre de bandes intégrées dans la robotique et dans le pool Free: " + $FreeTapeCount +  "<br />Ajouter des bandes dès que possible.</span></p>"}
$HtmlFreeTapeCount = $TABFreeTapeCount | ConvertTo-Html -Fragment
$HtmlFreeTapeCount = "<p><span style='color:BLACK'>" + "<b>Nombre de bandes dans le pool Free intégrées dans chaque robotique:</b>" + "</span></p>" + $HtmlFreeTapeCount

If ($WaitingTapeJob){ 
    $HtmlWaitingTapeJob = $WaitingTapeJob| ConvertTo-Html -Fragment
    $HtmlWaitingTapeJob = $HtmlWaitingTapeJob.Replace("<td>WaitingTape","<td style='color:RED'><b>WaitingTape</b>")
    $HtmlWaitingTapeJob = "<p><span style='color:RED'>" + "<b>Liste des Jobs Tape en attente de bandes:</b>"  + "</span></p>" + $HtmlWaitingTapeJob
}
Else {
    If ($TABWorkgingTapeJobs){
        $HtmlWorkgingTapeJob = $TABWorkgingTapeJobs | ConvertTo-Html -Fragment
        $HtmlWorkgingTapeJob = $HtmlWorkgingTapeJob.Replace("<td>Working","<td style='color:ORANGE'>Working")
        $HtmlWorkgingTapeJob = "<p><span style='color:BLACK'>" + "<b>Liste des Jobs Tape en cours:</b>" + "</span></p>" + $HtmlWorkgingTapeJob
    }
    Else {$HtmlWorkgingTapeJob = "<p><span style='color:BLACK'>" + "<b>Aucun Job Tape en cours.</b>" + "</span></p>"}
}

$HtmlDriveState = $TABDriveState | ConvertTo-Html -Fragment
$HtmlDriveState = $HtmlDriveState.Replace("<td>False","<td style='color:RED'>False")
$HtmlDriveState = "<p><span style='color:BLACK'>" + "<b>Etat des lecteurs:</b>" + "</span></p>" + $HtmlDriveState

If ($TABTapeOut){ 
    $HtmlTapeOut = $TABTapeOut | ConvertTo-Html -Fragment
    $HtmlTapeOut = "<p><span style='color:BLACK'>" + "<b>Liste des bandes à externaliser:</b>" + "</span></p>" + $HtmlTapeOut
}
Else {
    $HtmlTapeOut = "<p><span style='color:BLACJ'>" + "<b>Aucune bande à externaliser.</b>" + "</span></p>"
}

If ($TABTapeIn){ 
    $HtmlTapeIn = $TABTapeIn | ConvertTo-Html -Fragment
    $HtmlTapeIn = "<p><span style='color:BLACK'>" + "<b>Liste des bandes pouvant être intégrées dans la robotique:</b>" + "</span></p>" + $HtmlTapeIn
    if (($TapeIn).count -le $CriticalFreeTapeCount -and $FreeTapeCount -le $CriticalFreeTapeCount) {$HtmlTapeIn = "$HtmlTapeIn" + "<p><span style='color:ORANGE'>" + "<b>Le nombre de bandes expirées à intégrer ou disponibles dans le pool Free est faible, si nécessaire intégrer des bandes vierges</b>" + "</span></p>"}
}
Else {
    if ($FreeTapeCount -le $CriticalFreeTapeCount){
        $HtmlTapeIn = "<p><span style='color:RED'>" + "<b>Aucune bande expirée dans les pools et nombre faible de bandes disponibles dans le pool Free, insérer des bandes vierges si nécessaire.</b>" + "</span></p>"
    }
}

$HtmlHeader = '<html lang="fr"><head><meta charset="utf-8" />'
$HtmlHeader = $HtmlHeader  + '<title>Gestion des bandes Veeam</title>'
$HtmlHeader = $HtmlHeader  + '<style type="text/css">body{font-family: Tahoma}</style>'
$HtmlHeader = $HtmlHeader  + '<link rel="stylesheet" href="style.css"><script src="script.js"></script></head>'
$HtmlBody = "<body>" + $HtmlFreeTapeCount + $HtmlWorkgingTapeJob + $HtmlWaitingTapeJob + $HtmlDriveState + $HtmlTapeOut + $HtmlTapeIn + "</body>"
$HtmlFooter = '<footer><p>Athéo Ingénierie - PBG - ' + $Version + '<br /><a href="#">https://www.atheo.net</a></p></footer></html>'
$HtmlOutput = $HtmlHeader + $HTMLBody + $HtmlFooter

#Optimisations du format HTML
$HtmlOutput = $HtmlOutput.Replace("<table>","<table border=1 cellspacing=0 cellpadding=10>")
$HtmlOutput = $HtmlOutput.Replace("<tr><th>","<tr bgcolor='Silver'><th>")

#Export du fichier HTML
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
