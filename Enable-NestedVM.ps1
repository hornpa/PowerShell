<#

.SYNOPSIS
    Enable-NestedVM
.Description
    Checks VM for nesting comatability and configures if not properly setup.
.NOTES
    Author: 
    Patrik Horn (PHo)
    Link:	
    www.hornpa.de
    History:
    2017-04-18 - v1.0 - Script created (PHo)

#>

[cmdletbinding()]

Param
(
 
 	[parameter(Mandatory=$false)]
    [String]$LogPath = "$env:windir\Logs\Scripts",

    [parameter(Mandatory=$false)]
    [String]$ComputerName = "localhost"

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

    #region Select VMs

    $VMs = Get-VM -ComputerName $ComputerName | Out-GridView -Title "Select VM" -PassThru

    if([string]::IsNullOrEmpty($VMs))
    {

        $Msg = "No VM selected, Exit"
        Return $Msg
        Exit;

    }

    #endregion

    $ProgressBar_Summary = $VMs.Count
    $ProgressBar_Current = 1

    Foreach ($VM in $VMs)
    {
    
        $VM = Get-VM -Name $VM.Name -ComputerName $ComputerName

        $Msg = " $($VM.Name) ... ($ProgressBar_Current / $ProgressBar_Summary)"
        Write-Progress -Activity "Enabled Nested on VM..." -Status $Msg -PercentComplete ([math]::Round((100*$ProgressBar_Current)/$ProgressBar_Summary))
        $ProgressBar_Current++

        #region Shutdown VM

        $VM_Shutdown_Time = 60
        IF ($VM.State -ne 'Off' -or $VM.State -eq 'Saved')
        {
        
            $Msg = "VM is running or saved, will be shutting down in $VM_Shutdown_Time seconds"
            Write-Verbose $Msg
            Start-Sleep -Seconds $VM_Shutdown_Time 
            
            Try
            {
            
                Stop-VM -VMName $VM.Name -ComputerName $ComputerName
                $Result_Shutdown = "Successfully shutdown"

            }
            Catch
            {
            
                $Result_Shutdown = "Could not be shutdown VM!"

            }
           
            
        }
        Else
        {
        
            $Result_Shutdown = "VM is not running or saved"            

        }

        #endregion

        #region Expose Virtualization Extensions

        IF (((Get-VMProcessor -VM $vm).ExposeVirtualizationExtensions))
        {
        
            $Result_VMProcessor = "Exprose Virtualization Extensions is enabeld"

        }
        Else
        {
        
            $Msg = "Exprose Virtualization Extensions is not enabeld..."
            Write-Verbose $Msg

            Try
            {
            
                Set-VMProcessor -VMName $VM.Name -ComputerName $ComputerName -ExposeVirtualizationExtensions $true
                $Result_VMProcessor = "Successfully enabled"

            }
            Catch
            {
            
                $Result_VMProcessor = "Could not be enabled!"

            }

        }
        
        #endregion
        
        #region Dynamic Memory

        IF ($VM.DynamicMemoryEnabled -eq $true)
        {
        
            $Msg = "Dynamic Memory is enabled, will be disabeld"
            Write-Verbose $Msg
                        
            Try
            {
            
                Set-VMMemory -VMName $VM.Name -ComputerName $ComputerName -DynamicMemoryEnabled $false 
                $Result_DynamicMemory = "Successfully disabeld"

            }
            Catch
            {
            
                $Result_DynamicMemory = "Could not be disabeld!"

            }
           
            
        }
        Else
        {
        
            $Result_DynamicMemory = "Dynamic Memory is disabeld"
        
        }

        #endregion

        #region MAC Address Spoofing

        $NetworkAdapters = Get-VMNetworkAdapter $VM.Name

        Foreach ($NetworkAdapter in $NetworkAdapters)
        {

            IF ($NetworkAdapter.MacAddressSpoofing -eq 'Off')
            {
        
                $Msg = "MAC Address Spoofing is off (nested guests won't have network)."
                Write-Verbose $Msg
                        
                Try
                {
            
                    Get-VMNetworkAdapter -VMName $VM.Name | Set-VMNetworkAdapter -MacAddressSpoofing On
                    $Result_MacAddressSpoofing = "Successfully enabled"

                }
                Catch
                {
            
                    $Result_MacAddressSpoofing = "Could not be enabled!"

                }
           
            
            }
            Else
            {
        
                $Result_MacAddressSpoofing = "MAC Address Spoofing is on."

            }

        }

        #endregion

        #region Check VM Memory size

        $VM_Memory_recommendation = 4294967296
        IF ($VM.MemorySize -lt $VM_Memory_recommendation)
        {
        
            $Result_MemorySize = "VM memory is set less than 4GB, without 4GB or more, you may not be able to start VMs."
            
        }
        Else
        {
        
            $Result_MemorySize = "VM memory is equal or more than 4GB Memory."
        
        }

        #endregion

        #region Check Checkpoint

        IF ((Get-VMCheckpoint $VM.Name).Count -ge 1)
        {
        
            $Result_VMCheckpoint = "VM has one or more checkpoints, please remove them and make a new one."
            
        }
        Else
        {
        
            $Result_VMCheckpoint = "VM has no checkpoints, all fine."
        
        }

        #endregion

        #region Output / Result

        $LoopResult = New-Object PSObject
        
        Add-Member -InputObject $LoopResult NoteProperty -Name "VM Name" -Value $($VM.Name)
        Add-Member -InputObject $LoopResult NoteProperty -Name "Shutdown" -Value $Result_Shutdown
        Add-Member -InputObject $LoopResult NoteProperty -Name "VM Processor" -Value $Result_VMProcessor
        Add-Member -InputObject $LoopResult NoteProperty -Name "Dynamic Memory" -Value $Result_DynamicMemory
        Add-Member -InputObject $LoopResult NoteProperty -Name "Mac Address Spoofing" -Value $Result_MacAddressSpoofing
        Add-Member -InputObject $LoopResult NoteProperty -Name "Memory Size" -Value $Result_MemorySize
        Add-Member -InputObject $LoopResult NoteProperty -Name "Checkpoint" -Value $Result_VMCheckpoint

        $scriptResult += $LoopResult

        #endregion

    }

    #region Write Result

    $scriptResult | Export-Csv -Path $scriptResultOutput -NoTypeInformation -Append

    #endregion

#endregion

#region -----------------------------------------------------------[End]----------------------------------------------------------------------------	

    # End message
    $Msg = "Skript beendet um $(Get-Date -Format "HH:mm") am $(Get-Date -Format "yyyy-MM-dd")."
    Write-Output $Msg | Out-File $scriptLog -Append 

#endregion