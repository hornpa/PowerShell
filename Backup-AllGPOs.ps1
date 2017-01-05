<#
    .SYNOPSIS
     Sichert alle GPOs.
    .DESCRIPTION
     Die Funktion 'Backup-ALLGPOs_hp' sichert jede GPO in einen separaten Ordner inkl. eines HTML Report.
    .PARAMETER BackupPath
     Gibt den Backup Pfad an, Default ist 'C:\Temp'.
    .EXAMPLE
     Backup-ALLGPOs_hp
     Erstellt einen neuen Unterordner unter 'C:\Temp' mit dem Aktuellen Datum und Uhrzeit.
     In desem die GPOs gesichert werden.
    .EXAMPLE
     Backup-ALLGPOs_hp -BackupPath C:\GPO_Backup
     Erstellt einen neuen Unterordner unter 'C:\GPO_Backup' mit dem Aktuellen Datum und Uhrzeit.
     In desem die GPOs gesichert werden.
    .NOTES
     AUTHOR: 
      Patrik Horn (PHo)
	 HISTORY:
      2016-11-08 - v1.03 - Add Write-Progress
	  2016-10-18 - v1.02 - Change from Function to Script, Added Zip Function, Code Clean up (PHo)
	  2016-06-27 - v1.01 - Add Domain to Backup Path and Remove Special Characters from Export Path (PHo)
      2015-XX-XX - v1.00 - Script created as Function (PHo)    
    .LINK
      http://www.makrofactory.de
      http://www.hornpa.de
#>

[CmdletBinding(SupportsShouldProcess=$False)]

param(
    [parameter(Mandatory=$false)]
    [string]$BackupPath = "C:\MyPolicy" ,
    [parameter(Mandatory=$false)]
    [Switch]$Zip = $true
)

$Module_Name = "GroupPolicy"
IF(Get-Module -ListAvailable -Name $Module_Name) {
    Write-Host "Loading PowerShell Module: $Module_Name"
    Import-Module $Module_Name -Force
    }Else{
    $Msg = "$Module_Name Module is not available on the system, exit"
    Write-Host $Msg -ForegroundColor Red
    Return $Msg
}

Write-Verbose "Setting variables"
$Date = Get-Date -Format yyyy-MM-dd_HH-mm
$GPOs = Get-GPO -All | Sort-Object DisplayName
$ProgressBar_Summary = $GPOs.Count
$ProgressBar_Current = 0


Write-Verbose "Check if backup folder exists, if not it will be created"
IF (!(Test-Path $BackupPath)){
    Write-Verbose "Backup Folder will be created"
    New-Item -Path $BackupPath -ItemType Directory | Out-Null
}

Write-Verbose "Create Sub Folder for GPO export"
$BackupName = $Date + "_" + $env:USERDOMAIN + "_MsGPOs"
$BackupFolder = $BackupPath +"\" + $BackupName 
New-Item -Path $BackupFolder -ItemType Directory | Out-Null

Write-Verbose "Create GPO Report and Backup"
Foreach ($GPO in $GPOs){

    $Msg = " $($GPO.DisplayName) ... ($ProgressBar_Current / $ProgressBar_Summary)"

    Write-Progress -Activity "Export GPO..." -Status $Msg -PercentComplete ([math]::Round((100*$ProgressBar_Current)/$ProgressBar_Summary))

    $GPOName = $GPO.DisplayName

    $pattern = '[^a-zA-Z0-9()_+.,-]'

    $GPOName = $GPOName -replace $pattern, ' ' 

    Write-Verbose "Create Folder"

    $TempFolder =  $BackupFolder+"\"+$GPOName

    New-Item -Path $TempFolder -ItemType Directory | Out-Null

    Write-Verbose "Export GPO Report"

    $TempReportFile = $TempFolder+"\"+$GPOName +".html"

    Get-GPOReport -Guid $GPO.Id -ReportType Html -Path $TempReportFile | Out-Null

    Write-Verbose "Backup GPO"

    Backup-GPO -Guid $GPO.Id -Path $TempFolder | Out-Null

    Start-Sleep -Seconds 1

    $ProgressBar_Current++

}


IF ($Zip){

    Write-Verbose "Creating zip file"

    $Zip_Source_Path = $BackupFolder

    $Zip_File_Name = $BackupName

    $Zip_Path = $BackupPath

    Copy-Item -Path $Zip_Source_Path -Destination $env:TEMP -Recurse -Force

    $Zip_Tmp_Path = $env:TEMP + "\" + $Zip_File_Name

    $Zip_Dest_Path_File = $Zip_Path + "\" + $Zip_File_Name + ".zip"

    [Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" ) | Out-Null

    [System.IO.Compression.ZipFile]::CreateFromDirectory(($Zip_Tmp_Path+"\"), $Zip_Dest_Path_File) | Out-Null

    Remove-Item -Path $Zip_Tmp_Path -Recurse -Force | Out-Null

    Remove-Item -Path $Zip_Source_Path -Recurse -Force | Out-Null

}