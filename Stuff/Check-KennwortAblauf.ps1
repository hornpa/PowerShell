<#
    .SYNOPSIS
     Check-KennwortAblauf
    .DESCRIPTION
     Prüft ob das Benutzer Kennwort in x Tage abläuft.
    .PARAMETER Username
     Hier kann ein Benutzer mitgegeben werden.
     Standardmäßig ist hier der aktulle Benutzer hinterlegt.
    .PARAMETER Days
     Gibt den Zeitraum in Tagen an.
     Standardmäßig sind 7 Tage hinterlegt.
    .NOTES
		Author: 
		 Patrik Horn (PHo)
		Link:	
		 www.hornpa.de
		History:
		 2016-10-30 - v1.02 - Code Cleanup (PHo)
         2016-08-11 - v1.01 - Code change for PasswordExpriation (Pho)
         2016-04-25 - v1.00 - Script created (PHo)
#>
[CmdletBinding(SupportsShouldProcess=$False)]
param(
	[Parameter(Position=1)]
	[string]$UserName = $env:USERNAME,
    [Parameter(Position=2)]
    [int]$Days = 7,
    [Parameter(Position=3)]
    [int]$LogSizeMB = 1
	[Parameter(Position=4)]
	[string]$LogPath = $env:APPDATA\Scripts\Check-KennwortAblauf.log,
    )
Write-Verbose "Clear Error Variable Count"
$Error.Clear()

Write-Verbose "Setting System Varaible"
$scriptName = Split-Path $MyInvocation.MyCommand -Leaf
$scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
# Intiliaze GUI
[System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null

# Willkommens Nachricht
$Msg = "Skript gestartet um $(Get-Date -Format "HH:mm") am $(Get-Date -Format "yyyy-MM-dd") für Bentuzer $($env:USERNAME)."
Write-Output $Msg | Out-File $LogPath -Append 

Write-Verbose "Check Folder"
IF (!(Test-Path -path (Split-Path $LogPath))) {
	Write-Verbose "Create Folder"
	New-Item (Split-Path $LogPath) -Type Directory | Out-Null
}

Write-Verbose "Check if File exist"
IF (Test-Path $LogPath){
#Check  Size
    IF (((Get-Item "$LogPath").length/1MB) -ge $LogSizeMB){
		Remove-Item $LogPath -ErrorAction SilentlyContinue
    }
}

Write-Verbose "Check the OS"
$OS = (Get-WmiObject win32_operatingsystem) #| FL *
Switch -Wildcard ($OS.Version) {
    6.3* {
    $Messages = "Betriebssystem Windows Server 2012 R2 wurde erkannt."
    Write-Output $Messages | Out-File $LogPath -Append 
    }
    Default {
    $Messages = "Das Betriebssystem wird nicht durch dieses Skript unterstützt, Abbruch !"
    Write-Output $Messages | Out-File $LogPath -Append 
    Exit
    }
}

$Messages = "Benutzername lautet $UserName"
Write-Output $Messages | Out-File $LogPath -Append 
If ((get-aduser -identity $UserName -properties *).passwordneverexpires){
	$Messages = "Ihr Kennwort läuft nie ab."
    Write-Output $Messages | Out-File $LogPath -Append 
	Start-Sleep -s 2
	}Else{
	$today = Get-Date
	$expireson = (Get-ADUser -identity $UserName -Properties msDS-UserPasswordExpiryTimeComputed |  select @{ Name = "ExpirationDate"; Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}).ExpirationDate
	$daystoexpire=[math]::round((New-TimeSpan $(get-date -month $($today).Month -day $($today).Day -year $($today).Year) $(get-date -month $($expireson).Month -day $($expireson).Day -year $($expireson).Year)).TotalDays)
	If ($daystoexpire -lt 0){
        $daystoexpire = $daystoexpire * -1
        }
	$Messages = "Ihr Kennwort läuft in " + $daystoexpire + " Tagen ab."
    Write-Output $Messages | Out-File $LogPath -Append 
	Start-Sleep -s 2
	If ($daystoexpire -le $Days){
        $Messages = "Ihr Kennwort läuft in $daystoexpire Tagen ab.`n`nZum Ändern drücken Sie bitte folgende Tasten`n`nSTRG + ALT + ENTF`n`nund wählen 'Kennwort ändern'."
        [System.Windows.Forms.MessageBox]::Show( $Messages, "Kennwort Information", 0)
        Write-Output $Messages | Out-File $LogPath -Append 
        }
}

# Abschieds Nachricht
$Msg = "Skript beendet um $(Get-Date -Format "HH:mm") am $(Get-Date -Format "yyyy-MM-dd")."
Write-Output $Msg | Out-File $LogPath -Append 