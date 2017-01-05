<#
    .SYNOPSIS
     Create-ODBCConnections
    .DESCRIPTION
     Creates ODBC Connections from a XML File.
    .NOTES
     AUTHOR: 
      Patrik Horn (PHo)
     HISTORY:
      2016-09-23 - v1.0 - Release (PHo)
    .LINK
      http://www.hornpa.de
#>

[cmdletbinding()]

Param(
    [parameter(Mandatory=$false)]
    [String]$Path = (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent),
    [parameter(Mandatory=$false)]
    [String]$Name = "\Create-ODBCConnections.xml"
)

#$VerbosePreference = "Continue" # Continue = Shows Verbose Output / SilentlyContinue = No Verbose Output
#$ErrorActionPreference = 'Stop'

Write-Verbose "Clear Error Variable Count"
$Error.Clear()

Write-Verbose "Loading XML File"
[XML]$Settings_Global = Get-Content -Path ($Path+"\"+$Name)
$Connetions = $Settings_Global.ODBC.Connecntion

Foreach ($Connetion in $Connetions){
    
    Write-Verbose "Running $($ODBC_Name)..."

    Write-Verbose "Check Platform..."
    Switch -Wildcard ($Connetion.Platform){
        "32"{
            Write-Verbose "Platform is 32-Bit"
            $ODBC_Platform = "32-bit"
            $RegistryPath = "HKLM:\SOFTWARE\Wow6432Node\ODBC\ODBC.INI\"
        }
        "64"{
            Write-Verbose  "Platform is 36-Bit"
            $ODBC_Platform = "64-bit"
            $RegistryPath = "HKLM:\SOFTWARE\ODBC\ODBC.INI\"
        }
        Default{
            Write-Host "Fehler die Platform wurde nicht erkannt, bitte Prüfen!" -ForegroundColor Red
            Exit
        }
    }

    Write-Verbose "Check DsnType..."
    Switch -Wildcard ($Connetion.DsnType){
        "System"{
            Write-Verbose "DsnType is System"
            $ODBC_DnsType = "System"
        }
        "User"{
            Write-Verbose  "DsnType is User"
            $ODBC_DnsType = "User"
        }
        Default{
            Write-Host "Fehler der Type wurde nicht erkannt, bitte Prüfen!" -ForegroundColor Red
            Exit
        }
    }


    Write-Verbose "Check Driver..."
    $ODBC_DriverName = $Connetion.Driver
    IF (Get-OdbcDriver -Name $ODBC_DriverName -Platform $ODBC_Platform){
    Write-Verbose "Driver available"
    }Else{
    Write-Host "Driver not available" -ForegroundColor Red
    #Exit
    }

    Write-Verbose "Settings Variables"
    $ODBC_Name = $Connetion.Name
    $ODBC_Databases = $Connetion.Database
    $ODBC_Server = $Connetion.Server
    $ODBC_AnsiNPW = $Connetion.AnsiNPW
    $ODBC_LastUser = $Connetion.LastUser
    $ODBC_Description = $Connetion.Description

    $ODBC_Reg_Path = $RegistryPath+$ODBC_Name

    Write-Verbose "Creating ODBC Connection..."
    Try{
        Add-OdbcDsn -Name $ODBC_Name -DriverName $ODBC_DriverName -DsnType $ODBC_DnsType -Platform $ODBC_Platform -ErrorAction Stop
        Set-ItemProperty -Path $ODBC_Reg_Path -Name Database -Value $ODBC_Databases -ErrorAction Stop
        Set-ItemProperty -Path $ODBC_Reg_Path -Name Server -Value $ODBC_Server -ErrorAction Stop
        Set-ItemProperty -Path $ODBC_Reg_Path -Name Description -Value $ODBC_Description -ErrorAction Stop
        Set-ItemProperty -Path $ODBC_Reg_Path -Name LastUser -Value $ODBC_LastUser -ErrorAction Stop
        Set-ItemProperty -Path $ODBC_Reg_Path -Name AnsiNPW -Value $ODBC_AnsiNPW -ErrorAction Stop
        Set-ItemProperty -Path $ODBC_Reg_Path -Name Trusted_Connection -Value "No" -ErrorAction Stop
        Write-Host "ODBC Verbindung $ODBC_Name wurde erfolgreich angelegt." -ForegroundColor Green
        }Catch{
        Write-Host "ODBC Verbindung $ODBC_Name wurde nicht angelegt, bitte werte Prüfen!" -ForegroundColor Red
    }

}