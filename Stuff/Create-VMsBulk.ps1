<#
    .SYNOPSIS
     Create-VMsBulk.
    .DESCRIPTION
     Erstellt VMs anhand einer CSV.
    .NOTES
     AUTHOR: 
      Patrik Horn (PHo)
	 HISTORY:
      2016-12-25 - v1.00 - Script created (PHo)    
    .LINK
      http://www.hornpa.de
#>

[CmdletBinding(SupportsShouldProcess=$False)]

param(
    [parameter(Mandatory=$false)]
    [string]$PathtoCSV = "C:\Temp\VMs.csv"
   )

$Result = @{}

$VMs = Import-Csv -Path $PathtoCSV -Delimiter ";"

Foreach ($VM in $VMs){

    $VM_Name = $VM.VM_Name
    $VM_Path = $VM.VM_Path
    $VM_CPU = $VM.VM_CPU
    [int64]$VM_RAM = 1GB*($VM.VM_RAM)
    [int64]$VM_HDD_Size = 1GB*($VM.VM_HDD_Size)
    $VM_HDD_Path = ("$VM_Path\$VM_Name\Virtual Hard Disks\$VM_Name" +"_HDD0.vhdx")
    $VM_Switch = $VM.VM_Switch
    $VM_ISO = $VM.VM_ISO
     
    Write-Host -NoNewline "Running: " 
    Write-Host $VM_Name
    
    # Create VM
    Write-Host -NoNewline " - Create VM.."
    Try
    {
        New-VM -Name $VM_Name -Path $VM_Path -MemoryStartupBytes $VM_RAM -NewVHDPath $VM_HDD_Path -NewVHDSizeBytes $VM_HDD_Size -SwitchName $VM_Switch -Generation 2
        $Msg = "Successfully"
        Write-Host $Msg -ForegroundColor Green
        $Tmp_Result_VM = $Msg
    }
    Catch
    {
        $Msg = "Error: " + $Error[0].Exception.Message
        Write-Host $Msg -ForegroundColor Red
        $Tmp_Result_VM = $Msg
    }

    # Add DVD Drive
    Write-Host -NoNewline " - Add DVD Drive.."
    Try
    {
        Get-VM -Name $VM_Name | Add-VMDvdDrive -Path $VM_ISO
        $Msg = "Successfully"
        Write-Host $Msg -ForegroundColor Green
        $Tmp_Result_Add_DVDDrive = $Msg
    }
    Catch
    {
        $Msg = "Error: " + $Error[0].Exception.Message
        Write-Host $Msg -ForegroundColor Red
        $Tmp_Result_Add_DVDDrive = $Msg
    }

    # Configure VM
    Write-Host -NoNewline " - Configure VM.."
    Try
    {
        Get-VM -Name $VM_Name | Set-VM -ProcessorCount $VM_CPU -AutomaticStopAction Save -AutomaticStartAction StartIfRunning
        $Msg = "Successfully"
        Write-Host $Msg -ForegroundColor Green
        $Tmp_Result_VM_Settings = $Msg
    }
    Catch
    {
        $Msg = "Error: " + $Error[0].Exception.Message
        Write-Host $Msg -ForegroundColor Red
        $Tmp_Result_VM_Settings = $Msg
    }

    # Change Boot Order
    Write-Host -NoNewline " - Change Boot Order.."
    Try
    {
        $VM_BootOrder_network = Get-VMNetworkAdapter -VMName $VM_Name
        $VM_BootOrder_vhd = Get-VMHardDiskDrive -VMName $VM_Name
        $VM_BootOrder_dvd = Get-VMDvdDrive -VMName $VM_Name
        Set-VMFirmware -VMName $VM_Name -BootOrder $VM_BootOrder_dvd,$VM_BootOrder_vhd,$VM_BootOrder_network
        $Msg = "Successfully"
        Write-Host $Msg -ForegroundColor Green
        $Tmp_Result_BootOrder = $Msg
    }
    Catch
    {
        $Msg = "Error: " + $Error[0].Exception.Message
        Write-Host $Msg -ForegroundColor Red
        $Tmp_Result_BootOrder = $Msg
    }

    #region Result
    $Tmp_Result_Summary = @{
        "VM" = $VM_Name
        "Create VM" = $Tmp_Result_VM
        "Add DVD Drive" = $Tmp_Result_Add_DVDDrive
        "Set CPU, RAM, Settings" = $Tmp_Result_VM_Settings
        "Change Boot Order" = $Tmp_Result_BootOrder
    }

    $Result.Add($VM_Name,$Tmp_Result_Summary)

    #endregion

}

$Path_Result = "C:\Temp"+"\"+(Get-Date -Format yyyy-MM-dd_HHmm)+"_Result.xml"

$Result | Export-Clixml -Path $Path_Result
