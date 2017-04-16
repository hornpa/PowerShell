<#
    .SYNOPSIS
     Create-OUs
    .DESCRIPTION
     Creates OU in Active Directory from a CSV File.
    .NOTES
     AUTHOR: Patrik Horn
     HISTORY:
     2016-09-20 - v1.0 - Release (PHo)
    .LINK
     http://www.hornpa.de
#>

[cmdletbinding()]

Param(
    [parameter(Mandatory=$false)]
    [String]$Path = (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)+"\OUs.csv"
)

Write-Verbose "Clear Error Variable Count"
$Error.Clear()

Write-Verbose "Checking for Windows Feature"
IF (!((Get-WindowsFeature RSAT-AD-PowerShell).InstallState -like "Installed")){
Write-Host "Cloud not found Windows Feature (RSAT-AD-PowerShell), Script will be aborted!" -ForegroundColor Red
Exit
}

Try {

	$Entrys = Import-Csv $Path -Delimiter ";"
	Write-Verbose "CSV File succesfully loaded."
	
	}Catch{
	
	Write-Host "CSV File could not loaded." -ForegroundColor Red
    Write-Host "Error Message: " -ForegroundColor Red
    Write-Host $Error[0].Exception.Message -ForegroundColor Red

}

Foreach ($Element in $Entrys){

	Write-Verbose "Create OU..."

    $OUName = $Element.Name
    $OUPath = $Element.Path
    $OUDescription = $Element.Description

    Write-Verbose "OU Name: $OUName"
    Write-Verbose "OU Path: $OUPath"
    Write-Verbose "OU Description: $OUDescription"

    Try {
        
        New-ADOrganizationalUnit -Name $OUName -Path $OUPath -Description $OUDescription

        Write-Host "OU $OUName in $OUPath created." -ForegroundColor Green

        }Catch{

        Write-Host "OU $OUName in $OUPath could not created." -ForegroundColor Red
        Write-Host "Error Message: " -ForegroundColor Red
        Write-Host $Error[0].Exception.Message -ForegroundColor Red

    }

    
}
