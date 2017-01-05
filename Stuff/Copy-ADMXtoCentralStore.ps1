<#
    .SYNOPSIS
     Copy-ADMXtoCentralStore
    .DESCRIPTION
     Copy ADMX and ADML File to the Central Store.
    .NOTES
     AUTHOR: 
      Patrik Horn (PHo)
     HISTORY:
      2016-10-26 - v1.0 - Release (PHo)
    .LINK
      http://www.hornpa.de
#>
[cmdletbinding()]

Param(
    [parameter(Mandatory=$false)]
    [String]$ADMXPath = "\\g026spwmk1nas03.g026dpk1.g026.intern\SPKInstall\Citrix\W2K12\_Skripte\ADMX",
    [parameter(Mandatory=$false)]
    [String]$CentralStorePath = "\\$env:USERDNSDOMAIN\SYSVOL\$env:USERDNSDOMAIN\Policies\PolicyDefinitions"
)

Write-Verbose "Clear Error Variable Count"
$Error.Clear()

IF (!(Test-Path $CentralStorePath)){
    Write-Warning "No Central Store Found"
    Exit
}

Try {

    # Kopiert nur "*.admx" und "*.adml" Dateien in den Central Store
    Get-ChildItem -Path $ADMXPath  | Copy-Item -Destination $CentralStorePath -Include *admx ,*adml -Recurse -Force	
    $Msg = "All files copyed."
    Write-Host $Msg

	}Catch{
	
	Write-Host "Something went wrong, please check." -ForegroundColor Red
    Write-Host "Error Message: " -ForegroundColor Red
    Write-Host $Error[0].Exception.Message -ForegroundColor Red
    Exit

}