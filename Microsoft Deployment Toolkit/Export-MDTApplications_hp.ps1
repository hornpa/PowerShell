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
      2016-11-24 - v1.00 - Script created as Script (PHo)    
    .LINK
      http://www.hornpa.de
#>

[CmdletBinding(SupportsShouldProcess=$True)]
param(

    [Parameter(Mandatory=$false)]
    [string]$Root = "S:\DeploymentShare",

    [Parameter(Mandatory=$false)]
    [string]$Destination = "S:\___Backup"

)

$MS_MDT_PSDriveName = "MDT001"
$MDTDSRoot = $Root

Try {

    Import-Module "$env:ProgramFiles\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1" -Force
    New-PSDrive -Name $MS_MDT_PSDriveName -PSProvider "MDTProvider" -Root $MDTDSRoot -Description "Deployment Share"

	Write-Verbose "Modules succesfully loaded."
	
	}Catch{
	
	Write-Host "Modules could not loaded." -ForegroundColor Red
    Write-Host "Error Message: " -ForegroundColor Red
    Write-Host $Error[0].Exception.Message -ForegroundColor Red
    Exit

}

$Path_Temp = $env:TEMP + "\MDT_Export_Applications"
$Apps = Get-ChildItem -Path ($MS_MDT_PSDriveName+":\Applications") -Recurse | Where-Object { ($_.PSIsContainer -like "False") -and (!([string]::IsNullOrEmpty($_.GetPhysicalSourcePath()))) } | Sort-Object Name

$ProgressBar_Summary = $Apps.Count
$ProgressBar_Current = 0

Foreach ($App in $Apps){
    
    $Tmp_Name = $App.Name
    $Tmp_Version = $App.Version
    $Tmp_ShortName = $App.ShortName -replace ("-" + $Tmp_Version),""
    $Tmp_Publisher = $App.Publisher
    $Tmp_Reboot =  $App.Reboot
    $Tmp_Enable = $App.enable
    $Tmp_Hide = $App.hide
    $Tmp_Path = $App.GetPhysicalSourcePath()

    $Msg = " $Tmp_ShortName / $Tmp_Version ... ($ProgressBar_Current / $ProgressBar_Summary)"

    Write-Progress -Activity "Export Application..." -Status $Msg -PercentComplete ([math]::Round((100*$ProgressBar_Current)/$ProgressBar_Summary))

    $App_Tmp_Path_Copy = $Path_Temp + "\" + $Tmp_ShortName + "\" + $Tmp_Version

    Copy-Item -Path $Tmp_Path -Destination $App_Tmp_Path_Copy -Recurse

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

    #region Zip

    Write-Verbose "Creating zip file"

    $Zip_Source_Path = $Path_Temp

    $Zip_File_Name = $Tmp_ShortName + "-" + $Tmp_Version

    $Zip_Path = $Destination + "\" + $Tmp_Publisher 

    $Zip_Dest_Path_File = $Zip_Path + "\" + $Zip_File_Name + ".zip"

    IF (!(Test-Path $Zip_Path)){
        
        New-Item -Path $Zip_Path -ItemType Directory | Out-Null

    }

    [Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" ) | Out-Null

    [System.IO.Compression.ZipFile]::CreateFromDirectory(($Zip_Source_Path+"\"), $Zip_Dest_Path_File) | Out-Null

    Remove-Item -Path $Zip_Source_Path -Recurse -Force | Out-Null

    #endregion

    $ProgressBar_Current++

}