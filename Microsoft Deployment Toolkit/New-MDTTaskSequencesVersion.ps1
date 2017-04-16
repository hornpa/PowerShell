<#
.SYNOPSIS
    Erstellt eine neue Version einer Task Sequence auf Basis der alten.
.DESCRIPTION
    Die Funktion 'New-MDTTaskSequencesVersion_hp' Erstellt eine neue Version einer Task Sequence auf Basis der alten.
    Es ist wichtig das sowohl der Name als auch ID bereits eine v1.0 (_v1.0) besitzen da ansonsten es zu fehler kommt.
.PARAMETER MDTFolder
    Gibt den Pfad zum MDT Deployment Share an
.EXAMPLE
    New-MDTTaskSequencesVersion_hp 
.EXAMPLE
    New-MDTTaskSequencesVersion_hp -MDTFolder D:\TestDeploymentShare
.NOTES
    AUTHOR: 
    Patrik Horn (PHo)
    HISTORY:
    2016-12-08 - v1.01 - Change from Function to Script, complety redesign code, added gui  (PHo)
    2015-12-28 - v1.00 - Script created as Function (PHo)    
.LINK
    http://www.hornpa.de
#>

[CmdletBinding(SupportsShouldProcess=$True)]

param(

    [Parameter(Mandatory=$False,Position=1)]
    [string]$MDT_Path = "\\s01cne01\mdt_deployment\Produktion",

    [Parameter(Mandatory=$False,Position=2)]
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

$ListTS = Get-ChildItem -Path ($MDT_PSDriveName+":"+"Task Sequences") -Recurse | ?{!$_.PSIsContainer} | Select ID, Version, PSChildName, PSParentPath,guid | Sort-Object ID| Out-GridView -Title "Select Task Sequences" -PassThru
Foreach ($TS in ($ListTS | Sort-Object ID -Unique)) {
    
    Write-Verbose "Running: $($TS.PSChildName)"

    $SelectedTS = Get-ChildItem -Path ($MDT_PSDriveName+":"+"Task Sequences") -Recurse | ?{$_.ID -match $TS.ID} | Sort -Unique

    $Current_Date = Get-Date -Format yyyy-MM-dd
    $Current_Time = Get-Date -Format HH:mm
    $Current_Name = $SelectedTS.Name.Split("_")[0]
    $Current_Name_Version = $SelectedTS.Name.Split("_")[1]
    $Current_Name_Date = $SelectedTS.Name.Split("_")[2]

    $Current_ID_Name = $SelectedTS.ID.Split("_")[0]

    $Current_Release = $SelectedTS.Version.Split(".")[0]
    $Current_Version = $SelectedTS.Version.Split(".")[1]

    #region GUI Ask for new Major Release or new Version
    $Box_Title = "Frage an den Benutzer"
    $Box_Msg = "Soll ein neues Major Release erstellt?" + [System.Environment]::NewLine + 
               """$($SelectedTS.Name)"""
    $Box_Result = [System.Windows.Forms.MessageBox]::Show($Box_Msg,$Box_Title,3,[System.Windows.Forms.MessageBoxIcon]::Question)
 
    If ($Box_Result -eq "Yes")
    {
        Write-Verbose "Es wird ein neues Release erstellt."
        [string]$New_Release = [decimal]$Current_Release + 1
        [string]$New_Version = "0"

    }
    elseif ($Box_Result -eq "No")
    {
        Write-Verbose "Es wird kein neues Release erstellt."
        [string]$New_Release = $Current_Release
        [string]$New_Version = [decimal]$Current_Version + 1
    }
    else
    {
        $msg_title = "Abbruch"
        $msg = "Vorgang wurde abgebrochen"
        [System.Windows.Forms.MessageBox]::Show($msg,$msg_title,0)
        Exit
    }
    #endregion

    $New_ID = $Current_ID_Name + "_V" + $New_Release + "." + $New_Version
    $New_Name = $Current_Name + "_V" + $New_Release + "." + $New_Version + "_" + $Current_Date
    $New_MDT_Path =  $MDT_PSDriveName + ":\" + $SelectedTS.PSParentPath.Split("::")[3]
    $New_Comment = "Erstellt durch " + $env:USERDNSDOMAIN + "\" + $env:USERNAME + " am " + $Current_Date + " um " + $Current_Time 

    Write-Verbose "New name: $New_Name"
    Write-Verbose "New ID: $New_ID"
    Write-Verbose "New release number: $New_Release"
    Write-Verbose "New version number: $New_Version"
    Write-Verbose "New MDT path: $New_MDT_Path"

    # New Task Sequence
    try
    {
        Import-MDTTaskSequence -Name $New_Name -ID $New_ID -Version ($New_Release + "." + $New_Version) -Template "Server.xml" -Path $New_MDT_Path -Comments $New_Comment -ErrorAction Stop | Out-Null
        Write-Verbose "Successfully created new Task Sequence"
    }
    catch
    {
        Write-Verbose "Error could not created new Task Sequence"
        $msg_title = "Abbruch"
        $msg = "Die Task Sequece konnte nicht erstellt werden!"
        [System.Windows.Forms.MessageBox]::Show($msg,$msg_title,0,[System.Windows.Forms.MessageBoxIcon]::Error)
        Break
    }

    # Copy old TS to new TS
    try
    {
        Copy-Item -Path ($MDT_Path+"\Control\"+$SelectedTS.ID+"\*.*") -Destination ($MDT_Path+"\Control\"+$New_ID) -Force -ErrorAction Stop | Out-Null
        Write-Verbose "Successfully copyed files from old Task Sequence to new Task Sequence"
    }
    catch
    {
        Write-Verbose "Error could not copyed files from old Task Sequence to new Task Sequence"
        $msg_title = "Abbruch"
        $msg = "Die Datein von der alten Task Sequence konnten nicht kopiert werden! Bitte händisch prüfen!"
        [System.Windows.Forms.MessageBox]::Show($msg,$msg_title,0,[System.Windows.Forms.MessageBoxIcon]::Error)
    }

    Write-Verbose "----------"

}

Write-Verbose "Remove PS Drive"
Remove-PSDrive -Name $MDT_PSDriveName