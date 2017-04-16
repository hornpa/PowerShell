<#
.SYNOPSIS
    Exportiert alle MDT Task Sequencen.
.DESCRIPTION
    Die Funktion 'Export-MDTTaskSequences_hp' alle Task Seqeuncen aus MDT und deren Einstellungen.
.PARAMETER MDT_Path
    Gibt den Pfad zum MDT Deployment Share an
.PARAMETER Destination
    Legt fest in welchen Ordner die Exporte gespeichert werden sollen
.PARAMETER ALL
    Legt fest alle Task Sequencen exportiert. (Default = False)
.EXAMPLE
    Export-MDTTaskSequences_hp -MDT_Path D:\TestDeploymentShare -Destination D:\MDT_Export -All
    In diesem Beispiel werden alle Task Sequencen und deren Einstellungen innerhalb MDT Exportiert.
.EXAMPLE
    Export-MDTTaskSequences_hp -MDT_Path D:\TestDeploymentShare -Destination D:\MDT_Export
    In diesem Beispiel werden nur die ausgewählten Task Sequencen und deren Einstellungen innerhalb MDT Exportiert.
.NOTES
    AUTHOR: 
    Patrik Horn (PHo)
    HISTORY: 
    2017-04-14 - v1.01 - Swtich from function to script, Code redsign, Added Zip functions. Added Output (PHo)
    2015-12-28 - v1.00 - Created as function (PHo)
.LINK
    http://www.hornpa.de
#>

[CmdletBinding(SupportsShouldProcess=$True)]

param
(
    [Parameter(Mandatory=$False)]
    [string]$MDT_Path = "D:\DeploymentShare",

    [Parameter(Mandatory=$False)]
    [string]$Destination = "D:\Temp\Export\TS",

    [Parameter(Mandatory=$False)]
    [switch]$All,

    [Parameter(Mandatory=$False)]
    [switch]$Force = $True,

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

#region Select TS

IF($All)
{

    $ListTS =  Get-ChildItem -Path ($MDT_PSDriveName+":"+"Task Sequences") -Recurse | ?{!$_.PSIsContainer} | Select ID, Version, PSChildName,PSParentPath,guid

}
Else
{

    $ListTS = Get-ChildItem -Path ($MDT_PSDriveName+":"+"Task Sequences") -Recurse | ?{!$_.PSIsContainer} | Select ID, Version, PSChildName, PSParentPath,guid | Sort-Object ID| Out-GridView -Title "Select Task Sequences" -PassThru

}

#endregion

# Variables
$Path_Temp = $env:TEMP + "\MDT_Export_TaskSequences"
$Result = @()
$LogPath = $MDT_Path + "\_Files\CDS\" + "ExportTSs.csv"
$ProgressBar_Summary = $ListTS.Count
$ProgressBar_Current = 0

Foreach ($TS in ($ListTS | Sort-Object ID -Unique)) 
{

    $SelectedTS = Get-ChildItem -Path ($MDT_PSDriveName+":"+"Task Sequences") -Recurse | ?{$_.guid -match $TS.guid} | Sort -Unique
    $TSPropertys = Get-ItemProperty -Path ((Split-Path $SelectedTS.PSParentPath -NoQualifier)+"\"+$SelectedTS.PSChildName)

    $TSName = $TSPropertys.PSChildName
    $TSID = $TSPropertys.ID
    $TSEnable = $TSPropertys.enable
    $TSProviderPath = (Split-Path $TSPropertys.PSParentPath -NoQualifier) | Split-Path -NoQualifier
    $TSFolderPath = $SelectedTS.GetPhysicalSourcePath()
    $Export_Path_Tmp = $Path_Temp +"\"+$TSID

    $Msg = " $TSName ($TSID) ... ($ProgressBar_Current / $ProgressBar_Summary)"
    Write-Progress -Activity "Export Task Sequence..." -Status $Msg -PercentComplete ([math]::Round((100*$ProgressBar_Current)/$ProgressBar_Summary))
    $ProgressBar_Current++

    New-Item -Path $Export_Path_Tmp -ItemType Directory | Out-Null

    Copy-Item -Path ($TSFolderPath+"\*") -Destination $Export_Path_Tmp | Out-Null
 
    #region Setting File

    # Create XML
    Write-Verbose "Creating XML File"
    $XmlWriter = New-Object System.XMl.XmlTextWriter(($Export_Path_Tmp+"\MDTTaskSequence.xml"),$Null)
    # Set The Formatting
    $xmlWriter.Formatting = "Indented"
    $xmlWriter.Indentation = "4"
    # Write the XML Decleration
    $xmlWriter.WriteStartDocument()
    # Set the XSL
    $XSLPropText = "type='text/xsl' href='style.xsl'"
    $xmlWriter.WriteProcessingInstruction("xml-stylesheet", $XSLPropText)
    # Write Root Element
    $xmlWriter.WriteStartElement("TaskSequence")
    # Write the Document
    $xmlWriter.WriteElementString("Name",$TSName)
    IF ([string]::IsNullOrEmpty($TSEnable))
    {
    
        $xmlWriter.WriteElementString("Enable","False")
    
    }
    Else
    {
    
        $xmlWriter.WriteElementString("Enable","True")
    
    }
    # Write Close Tag for Root Element
    $xmlWriter.WriteEndElement > $null # <-- Closing RootElement
    # End the XML Document
    $xmlWriter.WriteEndDocument()
    # Finish The Document
    $xmlWriter.Finalize
    $xmlWriter.Flush > $null
    $xmlWriter.Close()

    #endregion
    
    #region Zip

    Write-Verbose "Creating zip file"

    $Zip_Source_Path = $Path_Temp

    $Zip_File_Name = $TSID

    $Zip_Path = $Destination

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
        Name = $TSName 
        ID = $TSID
        User = ($env:USERDNSDOMAIN+"\"+$env:USERNAME)
        Timestamp = $DateAdded
    }

    $Result += $Tmp_Result_Entry

    #endregion

} 

$Result | Export-Csv -Path $LogPath -NoTypeInformation

Return $Result