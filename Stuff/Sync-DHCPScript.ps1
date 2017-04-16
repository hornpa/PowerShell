<#
.SYNOPSIS
    Sync-DHCPScript
.Description
    Starten den Failover Sync für alle Scopes und repliziert zusätzlich die DHCP Server Option
.NOTES
    Author: 
    Patrik Horn (PHo)
    Link:	
    www.hornpa.de
    History:
    2017-02-22 - v1.2 - Redesign Code, some Bug fixing and new Features (Log, SyncServerOption,..) (PHo)
    2016-01-16 - v1.1 - Added Server Option Sync (PHo)
    2015-11-26 - v1.0 - Script created (PHo)
#>

[cmdletbinding()]

Param(
 
 	[parameter(Mandatory=$false)]
    [String]$LogPath = "$env:windir\Logs\Scripts",

    [parameter(Mandatory=$false)]
    [String]$SyncServerOption = $true

)

#region -----------------------------------------------------------[Pre-Initialisations]------------------------------------------------------------	

    #Set Error Action to Silently Continue
    $ErrorActionPreference = 'Stop'

    #Set Verbose Output
    $VerbosePreference = "SilentlyContinue" # Continue = Shows Verbose Output / SilentlyContinue = No Verbose Output

    #Get Start Time
    $StartPS = (Get-Date)

    #Set Enviorements
    Write-Verbose "Set Variable with MyInvocation"
    $scriptName_PS = Split-Path $MyInvocation.MyCommand -Leaf
    $scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    $scriptHelp = Get-Help "$scriptDirectory\$scriptName_PS" -Full
    $scriptName_SYNOPSIS = $scriptHelp.SYNOPSIS
    $scriptName_NOTES =  $scriptHelp.alertSet.alert.text
    $scriptEventlogSource = "PSScripts"
    $scriptLog = $LogPath + "\" + $scriptName_SYNOPSIS + ".log"

    # Checks if the Folder already exists
    if (!(Test-Path -path $LogPath))
    {
		Write-Verbose "Create Log Folder"
		New-Item $LogPath -Type Directory | Out-Null
    }

    # Check if Eventlog Source is avabile
    if ([System.Diagnostics.EventLog]::SourceExists($scriptEventlogSource ))
    {

	    Write-Verbose "Eventlog Source vorhanden"

    }
    Else
    {

	    Write-Verbose "Eventlog Source nicht vorhanden, wird erstellt..."
	    New-EventLog -LogName Application -Source $scriptEventlogSource 

    }

#endregion

#region -----------------------------------------------------------[Functions]----------------------------------------------------------------------	

    # Nothing to do

#endregion

#region -----------------------------------------------------------[Main-Initialisations]-----------------------------------------------------------

    # Start Entry in Eventlog
    $Msg = "Führe Script ""$scriptName_SYNOPSIS"" aus. Debug Log ist unter ""$scriptLog"" zu finden."
    Write-EventLog -LogName Application -Source $scriptEventlogSource -EntryType Information -EventId 1 -Message $Msg

    # Willkommens Nachricht
    $Msg = "Skript gestartet um $(Get-Date -Format "HH:mm") am $(Get-Date -Format "yyyy-MM-dd") für Bentuzer $($env:USERNAME)."
    Write-Output $Msg | Out-File $scriptLog 

#endregion

#region -----------------------------------------------------------[Execution]----------------------------------------------------------------------

    #region Check if DHCP is installed
        $Msg = "Prüfe DHCP Feature"
    Write-Output $Msg | Out-File $scriptLog -Append

    IF ((Get-WindowsFeature DHCP).Installed)
    { 
    
        $Msg = "DHCP Server installiert!"
        Write-Output $Msg | Out-File $scriptLog -Append

    }
    Else
    {

        $Msg = "Kein DHCP Server installiert, abbruch!"
        Write-Output $Msg | Out-File $scriptLog -Append
        Write-EventLog -LogName Application -Source $scriptEventlogSource -EntryType Error -EventId 1 -Message $Msg
        Exit

    }

    #endregion

    #region Check if only one Failover is configured, only when SyncServerOption is true 

    IF ($SyncServerOption)
    {

        $DHCPServerFailover = Get-DhcpServerv4Failover

        IF ($DHCPServerFailover.PartnerServer.Count -eq 1)
        {
            $Msg = "Es ist nur ein Failover Partner konfiguriert"
            Write-Output $Msg | Out-File $scriptLog -Append

        }
        Else
        {

            $Msg = "Es sind mehrere Failover Partner konfiguriert. Das wird derzeit nicht supported, abbruch!"
            Write-Output $Msg | Out-File $scriptLog -Append
            Write-EventLog -LogName Application -Source $scriptEventlogSource -EntryType Error -EventId 1 -Message $Msg
            Exit

        }

    }

    #endregion

    #region Replicate all Failover Scopes

    Try
    {

        Invoke-DhcpServerv4FailoverReplication -ComputerName $env:COMPUTERNAME -Force
        $Msg = "Alle Scopes wurden repliziert"
        Write-Output $Msg | Out-File $scriptLog -Append

    }
    Catch
    {

        $Msg = "Es trat ein Fehler bei der Replizierung auf, bitte prüfen! " + $Error[0].FullyQualifiedErrorId
        Write-Output $Msg | Out-File $scriptLog -Append

    }

    #endregion

    #region Replicate Server Options, only when SyncServerOption is true

    IF ($SyncServerOption)
    {

        $Computername = $DHCPServerFailover.PartnerServer
        $Msg = "Partner Server lautet:  $Computername"
        Write-Output $Msg | Out-File $scriptLog -Append

        #region Delete all DHCP Server Option on Partner
        $Msg = "Lösche DHCP Server Option..."
        Write-Output $Msg | Out-File $scriptLog -Append
        Try
        {

            Get-DhcpServerv4OptionValue -ComputerName $Computername -All | Remove-DhcpServerv4OptionValue -ComputerName $Computername
            $Msg = "Erfolgreich"
            Write-Output $Msg | Out-File $scriptLog -Append

        }
        Catch
        {

            $Msg = "Es trat ein Fehler auf, bitte prüfen!" + $Error[0].FullyQualifiedErrorId
            Write-Output $Msg | Out-File $scriptLog -Append

        }
        #endregion

        #region Add DHCP Classes
        $Msg = "Füge DHCP Classes hinzu... "
        Write-Output $Msg | Out-File $scriptLog -Append
        $DHCPServerClassv4 = Get-DhcpServerv4Class
        Foreach ($Entry in $DHCPServerClassv4)
        {
            $Msg = "Name: $($Entry.Name)" + [System.Environment]::NewLine + `                    "Type: $($Entry.Type)"+ [System.Environment]::NewLine + `
                    "Data: $($Entry.Data)"+ [System.Environment]::NewLine + `
                    "Description: $($Entry.Description)"

            Try
            {
    
                Add-DhcpServerv4Class -ComputerName $Computername -Name $Entry.Name -Type $Entry.Type -Data $Entry.Data -Description $Entry.Description -ErrorAction SilentlyContinue
                $Msg = "Erfolgreich hinzugefügt"
                Write-Output $Msg | Out-File $scriptLog -Append
    
            }
            Catch
            {
    
                $Msg = "Konnte nicht hinzugefügt werden. " + $Error[0].FullyQualifiedErrorId
                Write-Output $Msg | Out-File $scriptLog -Append

    
            }
    
            Write-Output "-----" | Out-File $scriptLog -Append
    
            Remove-Variable Entry -ErrorAction SilentlyContinue

        }
        #endregion

        #region Add DHCP Server Option Definition
        $Msg = "Füge DHCP Option Definition hinzu... "
        Write-Output $Msg | Out-File $scriptLog -Append
        $DHCPServerOptionDefinitionv4 = Get-DhcpServerv4OptionDefinition -All
        Foreach ($Entry in $DHCPServerOptionDefinitionv4)
        {
            $Msg = "Name: $($Entry.Name)" + [System.Environment]::NewLine + `
                    "OptionId: $($Entry.OptionId)" + [System.Environment]::NewLine + `
                    "VendorClass: $($Entry.VendorClass)" + [System.Environment]::NewLine + `
                    "Type: $($Entry.Type)" + [System.Environment]::NewLine + `
                    "Description: $($Entry.Description)"
    
            Try
            {
    
                Add-DhcpServerv4OptionDefinition -ComputerName $Computername -Name $Entry.Name -OptionId $Entry.OptionId -Type $Entry.Type -VendorClass $Entry.VendorClass -Description $Entry.Description -ErrorAction SilentlyContinue
                $Msg = "Erfolgreich hinzugefügt"
                Write-Output $Msg | Out-File $scriptLog -Append
    
            }
            Catch
            {
    
                $Msg = "Konnte nicht hinzugefügt werden. " + $Error[0].FullyQualifiedErrorId
                Write-Output $Msg | Out-File $scriptLog -Append

    
            }

            Write-Output "-----" | Out-File $scriptLog -Append
    
            Remove-Variable Entry -ErrorAction SilentlyContinue

        }
        #endregion

        #region Add DHCP Server Option
        $Msg = "Füge DHCP Server Option hinzu... " 
        Write-Output $Msg | Out-File $scriptLog -Append
        $DHCPServerOptionv4 = Get-DhcpServerv4OptionValue -All
        Foreach ($Entry in $DHCPServerOptionv4)
        {

            $Msg = "Name: $($Entry.Name)" + [System.Environment]::NewLine + `
                    "OptionId: $($Entry.OptionId)" + [System.Environment]::NewLine + `
                    "Value: $($Entry.Value)" + [System.Environment]::NewLine + `
                    "UserClass: $($Entry.UserClass)" + [System.Environment]::NewLine + `
                    "VendorClass: $($Entry.VendorClass)"

            Try
            {
    
                Set-DhcpServerv4OptionValue -ComputerName $Computername -OptionId $Entry.OptionId -Value $Entry.Value -UserClass $Entry.UserClass -VendorClass $Entry.VendorClass -Force -ErrorAction SilentlyContinue
                $Msg = "Erfolgreich hinzugefügt"
                Write-Output $Msg | Out-File $scriptLog -Append
    
            }
            Catch
            {
    
                $Msg = "Konnte nicht hinzugefügt werden. " + $Error[0].FullyQualifiedErrorId
                Write-Output $Msg | Out-File $scriptLog -Append

    
            }

            Write-Output "-----" | Out-File $scriptLog -Append

            Remove-Variable Entry -ErrorAction SilentlyContinue

        }
        #endregion

    }

    #endregion

#endregion

#region -----------------------------------------------------------[End]----------------------------------------------------------------------------	

    # Abschieds Nachricht
    $Msg = "Skript beendet um $(Get-Date -Format "HH:mm") am $(Get-Date -Format "yyyy-MM-dd")."
    Write-Output $Msg | Out-File $scriptLog -Append 

    # End Entry in Eventlog
    $Msg = "Das Script ""$scriptName_SYNOPSIS"" wurde ausgeführt. Debug Log ist unter ""$scriptLog"" zu finden."
    Write-EventLog -LogName Application -Source $EventlogSource -EntryType Information -EventId 1 -Message $Msg

#endregion