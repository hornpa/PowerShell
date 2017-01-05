<#
    .SYNOPSIS
     Export-CitrixPolicies
    .DESCRIPTION
     Export all Citrix Policies in a specific path.
    .NOTES
     AUTHOR: 
      Patrik Horn (PHo)
     HISTORY:
      2016-10-19 - v1.0 - Release (Patrik Horn)
    .LINK
      http://www.hornpa.de
#>

[cmdletbinding()]

Param(
    [parameter(Mandatory=$false)]
    [String]$Path = "C:\MyPolicy" ,
    [parameter(Mandatory=$false)]
    [String]$DCName = "localhost" ,
    [parameter(Mandatory=$false)]
    [Switch]$Zip = $true ,
    [parameter(Mandatory=$false)]
    [String]$CtxGroupPolicyModulePath = "C:\Program Files (x86)\Citrix\Scout\Current\Utilities\Citrix.GroupPolicy.Commands.psm1"
)

Write-Verbose "Load Modules"
Add-PSSnapin Citrix.Common.GroupPolicy
Import-Module $CtxGroupPolicyModulePath -Force

Write-Verbose "Create PSDrive $CTX_PS_Drive"
$CTX_PS_Drive = "CTX_Export_Drive"
New-PSDrive $CTX_PS_Drive –PSProvider CitrixGroupPolicy –Root \ -Controller $DCName | Out-Null

Write-Verbose "Complementary path varibale"
$Backup_Path_Folder_Name = (Get-Date -Format yyyy-MM-dd_HH-mm) + "_" + $env:USERDOMAIN + "_CtxPolicies"
$Backup_Path_Policies = $Path + "\" + $Backup_Path_Folder_Name

Write-Verbose "Check if backup folder exists, if not it will be created"
IF (!(Test-Path $Backup_Path_Policies)){
    New-Item -Path $Backup_Path_Policies -ItemType Directory | Out-Null
}

Write-Verbose "Get all Polices for User and Computers"
$Policies_User = Get-ChildItem ($CTX_PS_Drive+":\User")
$Policies_Computer = Get-ChildItem ($CTX_PS_Drive+":\Computer")

Write-Verbose "Export all Policies for the Citrix Site"
Export-CtxGroupPolicy -DriveName $CTX_PS_Drive -FolderPath $Backup_Path_Policies

Write-Verbose "Export User Policies for the Citrix Site in separate Folders"
Foreach ($Policie in $Policies_User){

    $Tmp_Folder_Path = $Backup_Path_Policies + "\User\" + $Policie.Name

    New-Item -Path $Tmp_Folder_Path -ItemType Directory | Out-Null

    Export-CtxGroupPolicy -DriveName $CTX_PS_Drive -PolicyName $Policie.Name -FolderPath $Tmp_Folder_Path

}

Write-Verbose "Export Computer Policies for the Citrix Site in separate Folders"
Foreach ($Policie in $Policies_Computer){

    $Tmp_Folder_Path = $Backup_Path_Policies + "\Computer\" + $Policie.Name

    New-Item -Path $Tmp_Folder_Path -ItemType Directory | Out-Null

    Export-CtxGroupPolicy -DriveName $CTX_PS_Drive -PolicyName $Policie.Name -FolderPath $Tmp_Folder_Path

}

Write-Verbose "Remove PSDrive $CTX_PS_Drive"
Remove-PSDrive $CTX_PS_Drive | Out-Null

IF ($Zip){

    Write-Verbose "Creating zip file"

    $Zip_Source_Path = $Backup_Path_Policies

    $Zip_File_Name = $Backup_Path_Folder_Name

    $Zip_Path = $Path

    Copy-Item -Path $Zip_Source_Path -Destination $env:TEMP -Recurse -Force

    $Zip_Tmp_Path = $env:TEMP + "\" + $Zip_File_Name

    $Zip_Dest_Path_File = $Zip_Path + "\" + $Zip_File_Name + ".zip"

    [Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" ) | Out-Null

    [System.IO.Compression.ZipFile]::CreateFromDirectory(($Zip_Tmp_Path+"\"), $Zip_Dest_Path_File) | Out-Null

    Remove-Item -Path $Zip_Tmp_Path -Recurse -Force | Out-Null

    Remove-Item -Path $Zip_Source_Path -Recurse -Force | Out-Null
    
}