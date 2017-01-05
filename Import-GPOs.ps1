<#
    .SYNOPSIS
     Import-GPOs
    .DESCRIPTION
     Import GPOs from a Path.
    .NOTES
     AUTHOR:
       Patrik Horn (PHo)
     HISTORY:
      2016-11-22 - v1.1 - Added Progressbar, Return Summary, Suffix and Präfix (PHo)
      2016-10-27 - v1.0 - Release (PHo)
    .LINK
      http://www.hornpa.de
#>

[cmdletbinding()]

Param(
    [parameter(Mandatory=$false)]
    [String]$Path = "C:\Users\hornpa\Desktop\Import_GPOs",

    [parameter(Mandatory=$false)]
    [String]$Präfix = "",

    [parameter(Mandatory=$false)]
    [String]$Suffix = ""

)

Write-Verbose "Clear Error Variable Count"
$Error.Clear()

Try {
          
    Import-Module GroupPolicy -Force
	Write-Verbose "Modules succesfully loaded."
	
	}Catch{
	
	Write-Host "Modules could not loaded." -ForegroundColor Red
    Write-Host "Error Message: " -ForegroundColor Red
    Write-Host $Error[0].Exception.Message -ForegroundColor Red
    Exit

}

Write-Verbose "Loading sub Folders"
Try {

    $GPOs = Get-ChildItem $Path -Directory -ErrorAction Stop
	Write-Verbose "GPOs loaded."
	
	}Catch{
	
	Write-Host "Not Path Found." -ForegroundColor Red
    Write-Host "Error Message: " -ForegroundColor Red
    Write-Host $Error[0].Exception.Message -ForegroundColor Red
    Exit

}

Write-Verbose "Setting Variable"
$Result = @()
$ProgressBar_Summary = $GPOs.Count
$ProgressBar_Current = 0

ForEach ($GPO in $GPOs) {
    
    Write-Verbose "Setting Path Variable for gpreport.xml"
    $GPO_ID = (Get-ChildItem $GPO.FullName -Directory).Name
    $GPO_Folder = $GPO.FullName
    $GPO_Name = $Präfix + $GPO.Name + $Suffix

    $Msg = " $($GPO_Name) ... ($ProgressBar_Current / $ProgressBar_Summary)"

    Write-Progress -Activity "Importing GPO..." -Status $Msg -PercentComplete ([math]::Round((100*$ProgressBar_Current)/$ProgressBar_Summary))

    $ProgressBar_Current++

    Try{
     
        Import-GPO -BackupId $GPO_ID -TargetName $GPO_Name -path $GPO_Folder -CreateIfNeeded -ErrorAction Stop | Out-Null

        Write-Verbose "Import successfully"

        $Tmp_Result = "succesfully"
                                
        }Catch{

        Write-Verbose "Error"

        $Tmp_Result = "Error: " + $Error[0].Exception.Message

    }

    $Tmp_Result_Entry = New-Object PSObject -Property @{
        GPO = $GPO_Name
        Path = $GPO_Folder
        Result = $Tmp_Result
    }

    $Result += $Tmp_Result_Entry
 
}

Return $Result