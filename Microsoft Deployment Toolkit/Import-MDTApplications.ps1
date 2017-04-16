<#
.SYNOPSIS
    Import Application into Microsoft Deployment Toolkit.
.DESCRIPTION
    This function seach for PSWrapper.ps1 and Deploy-Application.ps1 in a specific path and add them to the Microsoft Deployment Toolkit as Application.
.PARAMETER MDTFolder
    Path to the Deployment Share.
.PARAMETER Import_Path
    Path to the source files (recursive).
.PARAMETER Delete
    After the Application was successfully add the source files will be deleted.
.EXAMPLE
    Import-MDTApplications_hp -MDT_Path D:\TestDeploymentShare -Import_Path S:\Tausch\MDT_Import\Applications
    Powershell is search for Application in the Import Path and added them to Deployment Share, without deleting after the import
.EXAMPLE
    Import-MDTApplications_hp -MDT_Path D:\TestDeploymentShare -Import_Path S:\Tausch\MDT_Import\Applications -Delete
    Powershell is search for Application in the Import Path and added them to Deployment Share, after the successfully import the source files will be deleted.
.NOTES
    AUTHOR: 
    Patrik Horn (PHo)
    HISTORY:
    2017-04-11 - v1.02 - Added PSWraper XML support, added new Output (PHo)
    2016-12-09 - v1.01 - Change from Function to Script, added Support for PSWrapper.ps1  (PHo)
    2015-12-28 - v1.00 - Script created as Function (PHo)    
.LINK
    http://www.hornpa.de
#>

[CmdletBinding(SupportsShouldProcess=$True)]

param(

    [Parameter(Mandatory=$False,Position=1)]
    [string]$MDT_Path = "D:\DeploymentShare2",

    [Parameter(Mandatory=$False,Position=2)]
    [string]$Import_Path = "D:\Temp\Import",

    [Parameter(Mandatory=$False)]
    [switch]$Delete = $False,

    [Parameter(Mandatory=$False)]
    [string]$MDT_PSDriveName = "MTD01" 

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

#region Setting Variable
$Dir = $Import_Path
$Applications = Get-ChildItem -Path $Dir -Include Deploy-Application.ps1,PSWrapper.ps1 -Recurse
$LogPath = $MDT_Path + "\_Files\CDS\" + "ImportApplications.csv"
$ProgressBar_Summary = $Applications.Count
$ProgressBar_Current = 0
$Result = @()
#endregion

Foreach ($Application in $Applications)
{

    $Directory =  Split-Path $Application.FullName 
    $LongName = Split-Path -Path $Directory -parent
    $AppName = Split-Path -Path $LongName -Leaf
    $Version = Split-Path $Directory -Leaf
    $Install = Split-Path $Application.FullName -Leaf
    $CommandLine = "PowerShell.exe -ExecutionPolicy ByPass -File $Install"
    $DateAdded = Get-Date -Format "dd-MM-yyyy hh:mm"

    $Msg = " $AppName - $Version ... ($ProgressBar_Current / $ProgressBar_Summary)"
    Write-Progress -Activity "Import Application..." -Status $Msg -PercentComplete ([math]::Round((100*$ProgressBar_Current)/$ProgressBar_Summary))
    $ProgressBar_Current++

    $Application_ImportPath = Get-ChildItem -Path ($Application.FullName | Split-Path -Resolve)
    Switch -Wildcard ($Application_ImportPath.Name)
    {
        "*MDTApplication*"
        { 
        
            Write-Verbose "MDTApplication"
            [xml]$ApplicationSettings = Get-Content -Path (($Application.FullName | Split-Path -Resolve)+"\MDTApplication.xml") 
            $Path = $ApplicationSettings.Application.Path
            $Publisher = $ApplicationSettings.Application.Publisher
            $Enable = $ApplicationSettings.Application.Enable
            $Reboot = $ApplicationSettings.Application.Reboot
            $Hide = $ApplicationSettings.Application.Hide
            $Name = "$AppName-$Version"

        }
        "*PSApplication*" 
        { 
        
            Write-Verbose "PSApplication"
            [xml]$ApplicationSettings = Get-Content -Path (($Application.FullName | Split-Path -Resolve)+"\PSApplication.xml") 
            $Version1 = $ApplicationSettings.Application.Version
            $Product = $ApplicationSettings.Application.Product
            $Vendor = $ApplicationSettings.Application.Vendor
            $Path = $ApplicationSettings.Application.MDTSettings.Path + "\" + $Vendor
            $Publisher = $ApplicationSettings.Application.Publisher
            $Enable = $ApplicationSettings.Application.MDTSettings.Enable
            $Reboot = $ApplicationSettings.Application.MDTSettings.Reboot
            $Hide = $ApplicationSettings.Application.MDTSettings.Hide
            $Name = "$Vendor $Product $Version1"

        }
    }

    Try
    {
        $ImportMDTApp = Import-MDTApplication -Path ($MDT_PSDriveName + ":\Applications") `
        -enable "True"  `
        -Name "$Name"  `
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
        $Tmp_Result = "successfully added"

        Write-Verbose "Checking MDT Folder, IF NOT try to create..."
        IF (!(Test-Path  ($MDT_PSDriveName+":\"+$Path)))
        {
            Write-Verbose "MDT folder $Path is missing, will be created."
            New-Item -Path ($MDT_PSDriveName+":\"+$Path) -ItemType Directory -Force | Out-Null
        }

        $MDTPath = ($ImportMDTApp.PSParentPath -split ("$MDT_PSDriveName"+":"))[1]

        Write-Verbose "Set Application Properties"
        Set-ItemProperty ($MDT_PSDriveName+":\$MDTPath") -Name hide -Value $Hide
        Set-ItemProperty ($MDT_PSDriveName+":\$MDTPath") -Name Reboot -Value $Reboot
        Set-ItemProperty ($MDT_PSDriveName+":\$MDTPath") -Name enable -Value $Enable
        Set-ItemProperty ($MDT_PSDriveName+":\$MDTPath") -Name Publisher -Value $Publisher

        Write-Verbose "Move Application to $Path"
        Move-Item ($MDT_PSDriveName+":\$MDTPath") ($MDT_PSDriveName+":\"+$Path)

        IF (($Delete -eq $True) -and ($CheckImport.Count -eq '0'))
        {
            Remove-Item -Path $Directory -Force -Recurse
        }
        else
        {
            Write-Verbose "Error, could not be deleted becoue there was an error by importing."
        }

    }
    Catch [Microsoft.BDD.PSSnapIn.DeploymentPointException]
    {
        $Tmp_Result = "Application already exist!"
    }

    # Output
    $Tmp_Result_Entry = New-Object PSObject -Property @{
        AppName = $AppName
        Version = $Version  
        Install = $Install 
        SourcePath = $($Application.FullName)
        User = ($env:USERDNSDOMAIN+"\"+$env:USERNAME)
        Timestamp = $DateAdded
        Result = $Tmp_Result
    }

    $Result += $Tmp_Result_Entry
    
}

$Result | Export-Csv -Path $LogPath -NoTypeInformation -Append

Return $Result