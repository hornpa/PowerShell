<#
    .SYNOPSIS
     Importiert Anwendungen zum MDT.
    .DESCRIPTION
     Die Funktion 'Import-MDTApplications_hp' Anwendungen in dem Deploymentshare.
    .PARAMETER MDTFolder
     Gibt den Pfad zum MDT Deployment Share an.
    .PARAMETER Import_Path
     Legt fest in welchen Ordner die Importiere Anwendungen liegen.
    .PARAMETER Delete
     Nach Importieren der Anwendungen werden diese aus dem Import_Path gelöscht.
    .EXAMPLE
     Import-MDTApplications_hp -MDTFolder D:\TestDeploymentShare -Import_Path S:\Tausch\MDT_Import\Applications
     In diesem Beispiel werden alle Anwendungen und deren Einstellungen importiert.
    .NOTES
     AUTHOR: 
      Patrik Horn (PHo)
	 HISTORY:
	  2016-12-09 - v1.01 - Change from Function to Script, added Support for PSWrapper.ps1  (PHo)
      2015-12-28 - v1.00 - Script created as Function (PHo)    
    .LINK
      http://www.hornpa.de
#>

[CmdletBinding(SupportsShouldProcess=$True)]

param(

    [Parameter(Mandatory=$False,Position=1)]
    [string]$MDT_Path = "\\s01cne01\mdt_deployment\Produktion",


    [Parameter(Mandatory=$False,Position=2)]
    [string]$Import_Path = "\\s01cne01\mdt_deployment\KVBW\Import\Applications",

    [Parameter(Mandatory=$False)]
    [switch]$Delete,

    [Parameter(Mandatory=$False)]
    [string]$MDT_PSDriveName = "MTD01" 

)

# Load System.Windows.Forms
Add-Type –AssemblyName System.Windows.Forms

# Import PowerShell Module
try
{

    Import-Module "$env:ProgramFiles\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1" -Force
    New-PSDrive -Name $MDT_PSDriveName -PSProvider "MDTProvider" -Root $MDT_Path -Description "Deployment Share" | Out-Null

}
catch
{

    $msg_title = "Abbruch"
    $msg = "Konnte MDT PowerShell Module nicht laden!"
    [System.Windows.Forms.MessageBox]::Show($msg,$msg_title,0,[System.Windows.Forms.MessageBoxIcon]::Error)
    Write-Error 'The MDT PS module could not be loaded correctly, exit'
    Exit

}

$Dir = $Import_Path
$Applications = Get-ChildItem -Path $Dir -Include Deploy-Application.ps1,PSWrapper.ps1 -Recurse

Foreach ($Application in $Applications)
{
    [xml]$ApplicationSettings = Get-Content -Path (($Application.FullName | Split-Path -Resolve)+"\MDTApplication.xml")

    $Directory =  Split-Path $Application.FullName 
    $LongName = Split-Path -Path $Directory -parent
    $AppName = Split-Path -Path $LongName -Leaf
    $Version = Split-Path $Directory -Leaf
    $Publisher = $ApplicationSettings.Application.Publisher
    $Enable = $ApplicationSettings.Application.Enable
    $Reboot = $ApplicationSettings.Application.Reboot
    $Path = $ApplicationSettings.Application.Path
    $Hide = $ApplicationSettings.Application.Hide
    $Install = Split-Path $Application.FullName -Leaf
    $CommandLine = "PowerShell.exe -ExecutionPolicy ByPass -File $Install"

    Try
    {
        Import-MDTApplication -Path ($MDT_PSDriveName + ":\Applications") `
        -enable "True"  `
        -Name "$AppName-$Version"  `
        -ShortName "$AppName-$Version"  `
        -Version $Version  `
        -Publisher $Publisher `
        -Language ""  `
        -CommandLine $CommandLine  `
        -WorkingDirectory ".\Applications\$AppName\$Version"  `
        -ApplicationSourcePath $Directory  `
        -DestinationFolder "$AppName\$Version" `
        -ErrorVariable CheckImport `
        -ErrorAction Stop
    }
    Catch [Microsoft.BDD.PSSnapIn.DeploymentPointException]
    {
        Write-Host "Breits vorhande!"
    }


    IF (Test-Path  ($MDT_PSDriveName+":\"+$Path))
    {
        Write-Verbose "Ordner $Path ist bereits vorhanden"
    }
    Else
    {
        Write-Verbose "Ornder $Path ist nocht vorhanden, wird angelegt"
        New-Item -Path ($MDT_PSDriveName+":\"+$Path) -ItemType Directory -Force
    }

    Set-ItemProperty ($MDT_PSDriveName+":\Applications\$AppName-$Version") -Name hide -Value $Hide
    Set-ItemProperty ($MDT_PSDriveName+":\Applications\$AppName-$Version") -Name Reboot -Value $Reboot
    Set-ItemProperty ($MDT_PSDriveName+":\Applications\$AppName-$Version") -Name enable -Value $Enable
    Set-ItemProperty ($MDT_PSDriveName+":\Applications\$AppName-$Version") -Name Publisher -Value $Publisher

    Write-Verbose "Move Application to $Path"
    Move-Item ($MDT_PSDriveName+":\Applications\$AppName-$Version") ($MDT_PSDriveName+":\"+$Path)

    Write-Verbose "AppName $AppName"
    Write-Verbose "Version $Version"
    Write-Verbose "Publisher $Publisher"
    Write-Verbose "FullPath is $Application.FullName"
    Write-Verbose "Directory is $Directory"
    Write-Verbose "Install $Install"
    Write-Verbose "Commandline $Commandline"
    Write-Verbose "MDT Name is $AppName-$Version" 
    Write-Verbose "MDT Path is $Path" 
    Write-Verbose "MDT PS Drive Name is $MDT_PSDriveName" 

    IF (($Delete -eq $True) -and ($CheckImport.Count -eq '0'))
    {
        Remove-Item -Path $Name -Force -Recurse
    }
    else
    {
        Write-Verbose "Wurde nicht gelöscht, da es einen Fehler beim Importieren gab."
    }

    Write-Host "------"
}