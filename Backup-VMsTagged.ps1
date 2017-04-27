<#

.SYNOPSIS
	Backup-VMsTagged
.DESCRIPTION
	Export all VMs with "#Backup" in the descripton to a specific path.
.NOTES
	AUTHOR: 
	Patrik Horn (PHo)
	HISTORY:
	2017-04-23 - v1.00 - Script created (PHo)    
.LINK
	http://www.hornpa.de

#>

[cmdletbinding()]

Param
(

    [parameter(Mandatory=$false)]
    [String]$Path = "\\test.intra\Source\___Backup",

    [parameter(Mandatory=$false)]
    [int]$KeepLastBackups = 2,

    [parameter(Mandatory=$false)]
    [String]$LogPath = "$env:windir\Logs\Scripts"

)

#region -----------------------------------------------------------[Pre-Initialisations]------------------------------------------------------------	

    #Set Error Action to Silently Continue
    $ErrorActionPreference = "Stop" # Stop is recommend for Try and Catch blocks

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
    $scriptLog = $LogPath + "\" + $scriptName_SYNOPSIS + ".log"
    $scriptResultOutput = $LogPath + "\" + $scriptName_SYNOPSIS + ".csv"
    $scriptResult = @()

    # Checks if the Folder already exists
    if (!(Test-Path -path $LogPath))
    {
		Write-Verbose "Create Log Folder"
		New-Item $LogPath -Type Directory | Out-Null
    }

#endregion

#region -----------------------------------------------------------[Functions]----------------------------------------------------------------------	

    # Nothing to do

#endregion

#region -----------------------------------------------------------[Main-Initialisations]-----------------------------------------------------------

    # Welcome message
    $Msg = "Skript gestartet um $(Get-Date -Format "HH:mm") am $(Get-Date -Format "yyyy-MM-dd") durch Bentuzer $($env:USERNAME)."
    Write-Output $Msg | Out-File $scriptLog 

#endregion

#region -----------------------------------------------------------[Execution]----------------------------------------------------------------------

    # Settings
    $BackupPath = $Path
    $ExportPath = $BackupPath + "\" + (Get-Date -Format yyyy-MM-dd)

    #region Check beckup folder
    IF (Test-Path -Path $ExportPath)
    {
        $Ms = "Export folder $ExportPath exist"
        Write-Verbose $Msg
    
    }
    Else
    {

        $Ms = "Export folder don't $ExportPath exist, it will be created"
        Write-Verbose $Msg
        New-Item -Path $ExportPath -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
    
    }
    #endregion

    #region Backup VM when tagged found in description

    $VMs = Get-VM

    $ProgressBar_Summary = $VMs.Count
    $ProgressBar_Current = 1

    Foreach ($VM in $VMs)
    {

        $Msg = " $($VM.Name) ... ($ProgressBar_Current / $ProgressBar_Summary)"
        Write-Progress -Activity "Checking VM..." -Status $Msg -PercentComplete ([math]::Round((100*$ProgressBar_Current)/$ProgressBar_Summary))
        $ProgressBar_Current++

        IF (($VM.Notes) -match '#Backup')
        {

            $Msg = " $($VM.Name) ... ($ProgressBar_Current / $ProgressBar_Summary)"
            Write-Progress -Activity "Exporting VM..." -Status $Msg -PercentComplete ([math]::Round((100*$ProgressBar_Current)/$ProgressBar_Summary))
            Try
            {

                $Cmdlet = Export-VM -Path $ExportPath -Name ($VM.Name) -CaptureLiveState CaptureCrashConsistentState
                $Result_VMExport = "Successfully"


            }
            Catch
            {

                # Remove Folder if there was an error while exporting the VM
                Remove-Item ($ExportPath + "\" + $VM.Name) -Recurse -ErrorAction SilentlyContinue
                $Result_VMExport = "Error, $($Error[0].CategoryInfo)"

            }
        
        }
        Else
        {

            $Result_VMExport = "No Tag"
            Write-Verbose $Msg
        
        }

        #region Output / Result

        $LoopResult = New-Object PSObject
        
        Add-Member -InputObject $LoopResult NoteProperty -Name "VM Name" -Value $($VM.Name)
        Add-Member -InputObject $LoopResult NoteProperty -Name "VM Export Path" -Value $($ExportPath)
        Add-Member -InputObject $LoopResult NoteProperty -Name "VM Export Result" -Value $($Result_VMExport)

        $scriptResult += $LoopResult

        #endregion

    }
    #endregion

    #region Cleanup old Backups
    $KeepLastBackups = $KeepLastBackups + 1

    Try
    {
        
        Get-ChildItem $BackupPath -Directory | Sort-Object CreationTime -Descending | Select-Object -Skip $KeepLastBackups | Remove-Item -Recurse -Force
        $Msg = "Succussfully Cleanup"
        Write-Verbose $Msg
    
    }
    Catch
    {

        $Msg = "Error Cleanup, $($Error[0].CategoryInfo)"
        Write-Verbose $Msg

    }
    #endregion


#endregion

#region -----------------------------------------------------------[End]----------------------------------------------------------------------------	

    #region Write Result

    $scriptResult | Export-Csv -Path $scriptResultOutput -NoTypeInformation -Append

    #endregion

    #region End message

    $Msg = "Skript beendet um $(Get-Date -Format "HH:mm") am $(Get-Date -Format "yyyy-MM-dd")."
    Write-Output $Msg | Out-File $scriptLog -Append 

    #endregion

#endregion