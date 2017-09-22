<#
    .SYNOPSIS
     Get-FolderSize
    .DESCRIPTION
     Outputs the Size of the sub Folders.
    .PARAMETER FolderPath
     Path to the Folder.
    .PARAMETER LogPath
     Path to the Log.
    .PARAMETER Log
     When true a Log file will be written.
    .PARAMETER Console
     When true a Console Output is enabeld.
    .NOTES
     AUTHOR: 
      Patrik Horn (PHo)
	 HISTORY:
      2017-09-04 - v1.00 - Script created as Function (PHo)    
    .LINK
      http://www.makrofactory.de
      http://www.hornpa.de
#>

[CmdletBinding(SupportsShouldProcess=$False)]

param
(

    [parameter(Mandatory=$false)]
    [String]$FolderPath = "U:\Profiles" ,

    [parameter(Mandatory=$false)]
    [String]$LogPath = "C:\Temp" ,

    [parameter(Mandatory=$false)]
    [Switch]$Log = $true,

    [parameter(Mandatory=$false)]
    [Switch]$Console = $true

)

$Result = @()
$Sizes = 0

#region Get Folder Size

ForEach ($Item in (Get-ChildItem $FolderPath))
{
    If ($Item.PSIsContainer) 
    {
        $subFolderItems = Get-ChildItem $Item.FullName -recurse -force | Where-Object {$_.PSIsContainer -eq $false} | Measure-Object -property Length -sum | Select-Object Sum
    }
    Else 
    {
        Write-Host "No Subfolder"
    }
    $SizeFormated = "{0:N2}" -f ($subFolderItems.Sum / 1MB)

    $LoopResult = New-Object PSObject
        
    Add-Member -InputObject $LoopResult NoteProperty -Name "Folder" -Value $($Item)
    Add-Member -InputObject $LoopResult NoteProperty -Name "Size" -Value $($SizeFormated)

    $Result += $LoopResult
}

#endregion

$Date = Get-Date -Format "yyyy-MM-dd" 
$Time = Get-Date -Format "HHmm" 

#region Write Log File

IF ($Log)
{
    $OutFileName = "SubFolderSize"
    $OutputFile = $LogPath + "\" + $Date + "_" + $Time + "_" + $OutFileName + ".txt"

    $OutFile = $Result | Sort-Object -Property Size
    "Date:  $Date" | Out-File $OutputFile
    "Time:   $Time" | Out-File $OutputFile -Append
    "Path:   $FolderPath" | Out-File $OutputFile -Append
    $OutFile | Out-File $OutputFile -Append
}

#endregion

#region Write Output to Console

IF ($Console)
{
    Write-Output $outfile
}

#endregion