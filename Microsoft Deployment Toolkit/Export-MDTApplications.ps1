<#
.SYNOPSIS
    Export-MDTApplications_hp
.DESCRIPTION
    Die Funktion 'Export-MDTApplications_hp' alle Anwendungen aus MDT und deren Einstellungen als Zip Datei.
.PARAMETER MDTFolder
    Gibt den Pfad zum MDT Deployment Share an
.PARAMETER Destination
    Legt fest in welchen Ordner die Exporte gespeichert werden sollen
.EXAMPLE
    Export-MDTApplications_hp -Destination D:\MDT_Export
    In diesem Beispiel werden alle Anwendungen und deren Einstellungen innerhalb MDT Exportiert.
.NOTES
    AUTHOR: 
    Patrik Horn (PHo)
    HISTORY:
    2017-04-11 - v1.01 - Added PSWraper XML support, added new Output (PHo)
    2016-11-24 - v1.00 - Script created as Script (PHo)    
.LINK
    http://www.hornpa.de
#>

[CmdletBinding(SupportsShouldProcess=$True)]
param(

    [Parameter(Mandatory=$false)]
    [string]$MDT_Path = "D:\DeploymentShare2",

    [Parameter(Mandatory=$false)]
    [string]$Destination = "D:\Temp\Export",

    [Parameter(Mandatory=$False)]
    [switch]$All,

    [Parameter(Mandatory=$False)]
    [string]$MDT_PSDriveName = "MTD01" ,

    [Parameter(Mandatory=$False)]
    [switch]$Force = $True

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

#region Select Apps

IF($All)
{

    $Apps =  Get-ChildItem -Path ($MDT_PSDriveName+":\Applications") -Recurse | Where-Object { ($_.PSIsContainer -like "False") -and (!([string]::IsNullOrEmpty($_.GetPhysicalSourcePath()))) } | Sort-Object Name

}
Else
{

    $Apps = Get-ChildItem -Path ($MDT_PSDriveName+":\Applications") -Recurse | Where-Object { ($_.PSIsContainer -like "False") -and (!([string]::IsNullOrEmpty($_.GetPhysicalSourcePath()))) } | Sort-Object Name | Out-GridView -Title "Select Applications" -PassThru

}

#endregion

# Variables
$Path_Temp = $env:TEMP + "\MDT_Export_Applications"
$Result = @()
$LogPath = $MDT_Path + "\_Files\CDS\" + "ExportApplications.csv"
$ProgressBar_Summary = $Apps.Count
$ProgressBar_Current = 0

Foreach ($App in $Apps)
{

    $Tmp_Name = $App.Name
    $Tmp_ShortName = $App.ShortName
    $Tmp_Publisher = $App.Publisher
    $Tmp_Cmndl = $App.CommandLine
    $Tmp_Path = $App.GetPhysicalSourcePath()
    $DateAdded = Get-Date -Format "dd-MM-yyyy hh:mm"
    $App_Tmp_Path_Copy = $Path_Temp + "\" + ($Tmp_Path | Split-Path -Resolve | Split-Path -Leaf) + "\" + ($Tmp_Path | Split-Path -Leaf)

    $Msg = " $Tmp_Name ... ($ProgressBar_Current / $ProgressBar_Summary)"
    Write-Progress -Activity "Export Application..." -Status $Msg -PercentComplete ([math]::Round((100*$ProgressBar_Current)/$ProgressBar_Summary))
    $ProgressBar_Current++

    Switch -Wildcard ($Tmp_Cmndl)
    {
        "*Deploy-Application.ps1*"
        { 
        
            Write-Verbose "Detected Deploy-Application.ps1"

            $Tmp_Version = $App.Version
            $Tmp_Reboot =  $App.Reboot
            $Tmp_Enable = $App.enable
            $Tmp_Hide = $App.hide

            New-Item -Path $App_Tmp_Path_Copy -Type Directory | Out-Null

            #region Setting File

            # Create The Document
            $XmlWriter = New-Object System.XMl.XmlTextWriter(($App_Tmp_Path_Copy+"\MDTApplication.xml"),$Null)
            # Set The Formatting
            $xmlWriter.Formatting = "Indented"
            $xmlWriter.Indentation = "4"
            # Write the XML Decleration
            $xmlWriter.WriteStartDocument()
            # Set the XSL
            $XSLPropText = "type='text/xsl' href='style.xsl'"
            $xmlWriter.WriteProcessingInstruction("xml-stylesheet", $XSLPropText)
            # Write Root Element
            $xmlWriter.WriteStartElement("Application")
            # Write the Document
            $xmlWriter.WriteElementString("Name",$Tmp_Name)
            $xmlWriter.WriteElementString("ShortName",$Tmp_ShortName)
            $xmlWriter.WriteElementString("Version",$Tmp_Version)
            IF ([string]::IsNullOrEmpty($Tmp_Publisher)){
                $xmlWriter.WriteElementString("Publisher","Empty")
                }Else{
                $xmlWriter.WriteElementString("Publisher",$Tmp_Publisher)
            }
            IF ([string]::IsNullOrEmpty($Tmp_Hide)){
                $xmlWriter.WriteElementString("Hide","False")
                }Else{
                $xmlWriter.WriteElementString("Hide","True")
            }
            IF ([string]::IsNullOrEmpty($Tmp_Enable)){
                $xmlWriter.WriteElementString("Enable","False")
                }Else{
                $xmlWriter.WriteElementString("Enable","True")
            }
            IF ([string]::IsNullOrEmpty($Tmp_Reboot)){
                $xmlWriter.WriteElementString("Reboot","False")
                }Else{
                $xmlWriter.WriteElementString("Reboot","True")
            }
            $xmlWriter.WriteElementString("Path",($App.PSParentPath.TrimStart("MicrosoftDeploymentToolkit\MDTProvider::MDT001:")))
            # Write Close Tag for Root Element
            $xmlWriter.WriteEndElement > $null # <-- Closing RootElement
            # End the XML Document
            $xmlWriter.WriteEndDocument()
            # Finish The Document
            $xmlWriter.Finalize
            $xmlWriter.Flush > $null
            $xmlWriter.Close()

            #endregion

        }
        "*PSWrapper.ps1**" 
        { 
        
            Write-Verbose "Detected PSWrapper.ps1"
                
        }
    }
    
    Copy-Item -Path $Tmp_Path -Destination $App_Tmp_Path_Copy -Recurse

    #region Zip

    Write-Verbose "Creating zip file"

    $Zip_Source_Path = $Path_Temp

    $Zip_File_Name = $Tmp_ShortName

    $Zip_Path = $Destination + "\" + $Tmp_Publisher 

    $Zip_Dest_Path_File = $Zip_Path + "\" + $Zip_File_Name + ".zip"

    IF (!(Test-Path $Zip_Path))
    {
        
        New-Item -Path $Zip_Path -ItemType Directory | Out-Null

    }

    IF ($Force)
    {
        
        Remove-Item -Path $Zip_Dest_Path_File -Force -ErrorAction SilentlyContinue | Out-Null

    }

    [Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" ) | Out-Null

    [System.IO.Compression.ZipFile]::CreateFromDirectory(($Zip_Source_Path+"\"), $Zip_Dest_Path_File) | Out-Null

    Remove-Item -Path $Zip_Source_Path -Recurse -Force | Out-Null

    #endregion
    
    #region Output

    $Tmp_Result_Entry = New-Object PSObject -Property @{
        AppName = $Tmp_Name
        SourcePath = $Tmp_Path
        User = ($env:USERDNSDOMAIN+"\"+$env:USERNAME)
        Timestamp = $DateAdded
    }

    $Result += $Tmp_Result_Entry

    #endregion

}

$Result | Export-Csv -Path $LogPath -NoTypeInformation -Append

Return $Result