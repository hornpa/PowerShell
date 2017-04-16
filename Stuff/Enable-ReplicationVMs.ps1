<#
    .SYNOPSIS
     Enable-VMReplications
    .DESCRIPTION
     Enable Replications for all VMs on a Host
    .NOTES
     AUTHOR: Patrik Horn
     HISTORY:
     2016-10-14 - v1.0 - Release (PHo)
    .LINK
     http://www.makrofactory.de
#>

[cmdletbinding()]

Param(
    [parameter(Mandatory=$false)]
    [String]$Host_Source_Server = "$env:COMPUTERNAME.$env:USERDNSDOMAIN",
    [parameter(Mandatory=$true)]
    [String]$Host_Replication_Server = "Z013SPWMK1HV301.Z013DPK1.Z013.INTERN",
    [String]$Host_Replication_Share = "G$\VM-Replica",
    [String]$Host_Replication_Local_Path = "G:\VM-Replica"
)

[String]$Host_Replication_Port = "80"
[String]$Host_Replication_Auth = "Kerberos"
[String]$Host_Replication_Sync_Intervall = "900"

Write-Verbose "Clear Error Variable Count"
$Error.Clear()

$VMstoEnableReplica = Get-VM * | Where-Object { $_.ReplicationState -like "Disabled" } | Sort-Object Name

$Destination_Path = "\\$Host_Replication_Server\$Host_Replication_Share"
$Source_Path = "$Host_Source_Server"

Foreach ($VM in $VMstoEnableReplica){

    $VM_Name = $VM.VMName
    $VM_Path = $VM.Path

    Write-Host "Enabling Replica for $VM_Name"

    Write-Host "Copy VM to Replication Host"
    Copy-Item -Path $VM_Path -Destination $Destination_Path -Recurse -Force

    Write-Host "Import VM on Replication Host"
    $VM_GUID = Get-ChildItem ($Destination_Path+"\"+$VM_Name+"\Virtual Machines") -Filter "*.vmcx"
    Import-VM -Path ($Host_Replication_Local_Path+"\"+$VM_Name+"\Virtual Machines\"+$VM_GUID.Name) -Register -ComputerName $Host_Replication_Server

    Write-Host "Enable Replication on VM"
    Enable-VMReplication -ComputerName $Host_Replication_Server -VMName $VM_Name -AsReplica
    Enable-VMReplication -ComputerName $Host_Source_Server -VMName $VM_Name  -ReplicaServerName $Host_Replication_Server -ReplicaServerPort $Host_Replication_Port -AuthenticationType $Host_Replication_Auth
    Set-VMReplication -ComputerName $Host_Source_Server -VMName $VM_Name -ReplicationFrequencySec $Host_Replication_Sync_Intervall
    Start-VMInitialReplication -ComputerName $Host_Source_Server -VMName $VM_Name -UseBackup

}




