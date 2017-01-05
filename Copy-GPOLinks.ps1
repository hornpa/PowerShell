<#
    .SYNOPSIS
     Copy-GPOLinks
    .DESCRIPTION
     Copy GPO Links from one OU to antoher.
     User Formt "OU=Win2012R2,OU=RDS,OU=Citrix VDA,OU=Servers,OU=Test,DC=test,DC=intra" for Source and DestinationOU
    .NOTES
     AUTHOR: 
      Patrik Horn (PHo)
     HISTORY:
      2016-09-21 - v1.0 - Release (PHo)
    .LINK
      http://www.hornpa.de
#>

[cmdletbinding()]

Param(
    [parameter(Mandatory=$true)]
    [String]$SourceOU,
    [parameter(Mandatory=$true)]
    [String]$DestinationOU

)

Write-Verbose "Clear Error Variable Count"
$Error.Clear()

Try {

	Import-Module GroupPolicy -Force
	Write-Verbose "Module succesfully loaded."
	
	}Catch{
	
	Write-Host "Module could not loaded." -ForegroundColor Red
    Write-Host "Error Message: " -ForegroundColor Red
    Write-Host $Error[0].Exception.Message -ForegroundColor Red
    Exit

}

Try {

	# Get the linked GPOs 
    $linked = (Get-GPInheritance -Target $SourceOU).gpolinks 
	
	}Catch{
	
	Write-Host "Could not loaded GPO list." -ForegroundColor Red
    Write-Host "Error Message: " -ForegroundColor Red
    Write-Host $Error[0].Exception.Message -ForegroundColor Red
    Exit

}

# Loop through each GPO and link it to the target 
Foreach ($link in $linked){
 
    $guid = $link.GPOId 
    $order = $link.Order 
    $enabled = $link.Enabled

    if ($enabled){ 
        $enabled = "Yes" 
        }else{ 
        $enabled = "No" 
    } 

    # Create the link on the target 
    New-GPLink -Guid $guid -Target $DestinationOU -LinkEnabled $enabled -confirm:$false 
    # Set the link order on the target 
    Set-GPLink -Guid $guid -Target $DestinationOU -Order $order -confirm:$false
     
}