<#
    .SYNOPSIS
     Backup-SQLDatabases
    .DESCRIPTION
     Creates Backup for all Databases in a Instance.
    .NOTES
     AUTHOR: 
      Patrik Horn (PHo)
     HISTORY:
      2016-09-20 - v1.0 - Release (PHo)
    .LINK
      http://www.hornpa.de
#>

[cmdletbinding()]

Param(
    [parameter(Mandatory=$false)]
    [String]$Server = $env:COMPUTERNAME,
    [parameter(Mandatory=$false)]
    [String]$DBInstance = "Default",
    [parameter(Mandatory=$true)]
    [String]$BackupPath = (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)

)

Write-Verbose "Clear Error Variable Count"
$Error.Clear()

Write-Verbose "Get Date"
$Date = Get-Date -Format yyyyMMdd_HHmm

Write-Verbose "Get Location"
$PS_Location = Get-Location

Try {

	Import-Module "sqlps" -DisableNameChecking
	Write-Verbose "Module succesfully loaded."
	
	}Catch{
	
	Write-Host "Module could not loaded." -ForegroundColor Red
    Write-Host "Error Message: " -ForegroundColor Red
    Write-Host $Error[0].Exception.Message -ForegroundColor Red
    Exit

}

Write-Verbose "Set Location"
Set-Location -Path $PS_Location

Try{

    $DBs = Get-ChildItem "SQLSERVER:\SQL\$Server\$DBInstance\Databases"
	Write-Verbose "DB succesfully connected."

    }catch{

	Write-Host "DB could not connected." -ForegroundColor Red
    Write-Host "Error Message: " -ForegroundColor Red
    Write-Host $Error[0].Exception.Message -ForegroundColor Red
    Exit
}

IF ($DBInstance -like "Default"){

    Write-Verbose "MS SQL Default Instance"
    $DBConnection = $Server

    }Else{

    Write-Verbose "MS SQL Costum Instance"
    $DBConnection = $Server+"\"+$DBInstance

}

Foreach ($DB in $DBs){

	Write-Verbose "Create Backup..."

    $DBName = $DB.Name
    $BackupFile = ($BackupPath+"\"+$DBName+"_"+$Date+".bak")

    Try {
        
        Backup-SqlDatabase -BackupAction Database -ServerInstance $DBConnection -Database $DBName -BackupFile $BackupFile
        Write-Host "Successfully created Backup for $DBName in $BackupFile ." -ForegroundColor Green

        }Catch{

        Write-Host "Could not created Backup for $DBName in $BackupFile" -ForegroundColor Red
        Write-Host "Error Message: " -ForegroundColor Red
        Write-Host $Error[0].Exception.Message -ForegroundColor Red

    }

    IF ($DB.RecoveryModel -like "Simple"){
        
        Write-Host "Recovery Model is Simple, skip Backup for Transactionlog" -ForegroundColor Yellow

        }Else{

            Try {
        
                Backup-SqlDatabase -BackupAction Log -ServerInstance $DBConnection -Database $DBName -BackupFile $BackupFile
                Write-Host "Successfully created Transactionlog Backup for $DBName in $BackupFile ." -ForegroundColor Green

                }Catch{

                Write-Host "Could not created Transactionlog Backup for $DBName in $BackupFile" -ForegroundColor Red
                Write-Host "Error Message: " -ForegroundColor Red
                Write-Host $Error[0].Exception.Message -ForegroundColor Red

            }

    }

    
}