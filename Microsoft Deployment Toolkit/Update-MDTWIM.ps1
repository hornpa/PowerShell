<#

.SYNOPSIS
    Update-MDTWIM.
.DESCRIPTION
    Integriert alle Patch (MSU und CAB) aus dem Pfad in eine WIM File.
.PARAMETER Count
    Legt die Anzahl der durchläufe fest, wie oft probiert werden soll Updates zu integrieren.
.PARAMETER UpdatesPath
    Pfad zu den Updates.
.PARAMETER MountPath
    Pfad zu Mount Punkt, sollte er nicht vorhanden sein wird er erstellt.
.PARAMETER WimFile
    Pfad inkl. Name zum WIM File welches aktualisiert werden soll.
.PARAMETER WimIndex
    Gibt an welche Version gepatcht werden soll, es muss eine ausgewählt werden.
.NOTES
    AUTHOR: 
    Patrik Horn (PHo)
    HISTORY:
    2017-04-14 - v1.01 - Added Feature installation and Cleanup WIM File (PHo)
    2017-04-01 - v1.01 - Code optimaize (PHo)
    2017-03-23 - v1.00 - Script created as Function (PHo)    
.LINK
    http://www.makrofactory.de
    http://www.hornpa.de

#>

[CmdletBinding(SupportsShouldProcess=$False)]

param(


    [Parameter(Mandatory=$false)]
    [string]$MDT_Path = "D:\DeploymentShare",

    [Parameter(Mandatory=$False)]
    [string]$MountPath = “D:\Mount”,

    [Parameter(Mandatory=$False)]
    [string]$MDT_PSDriveName = "MTD01",

    [Parameter(Mandatory=$False)]
    [string]$Count = 3 

)

# Load System.Windows.Forms
Add-Type –AssemblyName System.Windows.Forms

#region Import PowerShell Module
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
#endregion


$OS = Get-ChildItem -Path ($MDT_PSDriveName+":\Operating Systems") -Recurse | Where-Object { ($_.PSIsContainer -like "False") -and (!([string]::IsNullOrEmpty($_.GetPhysicalSourcePath()))) } | Sort-Object Name | Out-GridView -Title "Select Applications" -PassThru

# $OS | FL *
$Date = (Get-Date -Format yyyyHHMM)
$ProgressBarAcitvity_Summary = 8
$UpdatesPath = $MDT_Path + "\_Files\CDS\Update\Source"
$WimIndex = $OS.ImageIndex
$WimName = ($OS.Source | Split-Path -Leaf)
$Source_WimFile = (Get-ChildItem $OS.GetPhysicalSourcePath() -Filter "install.wim" -Recurse).FullName
$Tmp_Dest_WimFile_Folder = $MDT_Path + "\_Files\CDS\Update\Temp"
$Dest_WimFile = $Tmp_Dest_WimFile_Folder + "\" + $WimName + "_" + $Date + ".wim"

#region Check if Folder exists
IF (!(Test-Path $MountPath))
{

    Write-Verbose "$MountPath not Found, Folder will be created."
    New-Item -Path $MountPath -ItemType Directory -Force | Out-Null

}

IF (!(Test-Path $Tmp_Dest_WimFile_Folder))
{

    Write-Verbose "$Tmp_Dest_WimFile_Folder not Found, Folder will be created."
    New-Item -Path $Tmp_Dest_WimFile_Folder -ItemType Directory -Force | Out-Null

}
#endregion

Write-Progress -Activity "Copy WIM (1/$ProgressBarAcitvity_Summary)..."
Copy-Item -Path $Source_WimFile -Destination $Dest_WimFile

Write-Progress -Activity "Mounting WIM (2/$ProgressBarAcitvity_Summary)..."
$Command_Mount = DISM /Mount-Wim /WimFile:$Dest_WimFile /index:$WimIndex /Mountdir:$MountPath

Write-Progress -Activity "Install Features (3/$ProgressBarAcitvity_Summary)..."
$SourceSXS = ($OS.GetPhysicalSourcePath() + "\sources\sxs")
$Command_Feature = DISM /Image:$MountPath /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:$SourceSXS 

Write-Progress -Activity "Loading Updates (4/$ProgressBarAcitvity_Summary)..."
$UpdateArray = Get-ChildItem $UpdatesPath -Include "*msu", "*cab" -Recurse
$CurrentCount = 0
Do
{

    $CurrentCount++
    Write-Verbose "Current Run is $CurrentCount " 

    #region Integration Windows Updates ...

    $ProgressBar_Current = 1
    $ProgressBar_Summary = $UpdateArray.Count

    ForEach ($Updates in $UpdateArray)
    {

        $Msg = " $($Updates.Name) ... ($ProgressBar_Current / $ProgressBar_Summary)"
        Write-Progress -Activity "Integration Windows Updates (4/$ProgressBarAcitvity_Summary) - Running $CurrentCount from $Count..." -Status $Msg -PercentComplete ([math]::Round((100*$ProgressBar_Current)/$ProgressBar_Summary))
        $ProgressBar_Current++

        $Command_Update = DISM /image:$MountPath /Add-Package /Packagepath:$Updates

        Start-Sleep –s 1

    }
    #endregion

}
Until($CurrentCount -eq $Count)

Write-Progress -Activity "Cleanup WIM File (5/$ProgressBarAcitvity_Summary)..."
$Command_Cleanup_AnalyzeComponentStore = Dism /Image:$MountPath /Cleanup-Image /AnalyzeComponentStore
$Command_Cleanup_StartComponentCleanup = Dism /Image:$MountPath /Cleanup-Image /StartComponentCleanup
$Command_Cleanup_SPSuperseded = Dism /Image:$MountPath /Cleanup-Image /SPSuperseded

Write-Progress -Activity "Unmount WIM (6/$ProgressBarAcitvity_Summary)..."
$Command_Unmount = DISM /Unmount-Wim /Mountdir:$MountPath /commit

Write-Progress -Activity "Cleanup System (7/$ProgressBarAcitvity_Summary)..."
$Command_Cleanup = DISM /Cleanup-Wim

Write-Progress -Activity "Import Image to MDT (8/$ProgressBarAcitvity_Summary)..."
New-Item -Path (($OS.PSParentPath | Split-Path -NoQualifier) + " - Update $Date") -ItemType Directory | Out-Null
Import-MDTOperatingSystem -Path (($OS.PSParentPath | Split-Path -NoQualifier) + " - Update $Date")  -SourceFile $Dest_WimFile -DestinationFolder ($WimName + "_" + $Date) -Move
Foreach ($Element in (Get-ChildItem (($OS.PSParentPath | Split-Path -NoQualifier) + " - Update $Date")))
{

    IF (!($Element.ImageIndex -eq $WimIndex))
    {
    
    Write-Verbose "Removing $Element"
    Remove-Item ($Element.PSPath | Split-Path -NoQualifier)

    }


}


