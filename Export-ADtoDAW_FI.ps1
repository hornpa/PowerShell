<#
    .SYNOPSIS
     Export-ADtoDAW_FI
    .DESCRIPTION
     Export Users from AD for Import-CSV to DAW
    .NOTES
     AUTHOR: 
	  MakroFactory Gmbh & Co. KG.
	   Thomas Neumann (TNe)
	   Patrik Horn (PHo)
	 LINK:
      http://www.makrofactory.de
     HISTORY:
      2016-11-10 - v1.0 - Release (PHo / TNe)
#>

$SourceDomain = "sskmg.intern"

#Current Path
$currentExecutionPath = $MyInvocation.MyCommand.Definition.Replace($MyInvocation.MyCommand.Name, "")
$csvreportfile = "$currentExecutionPath\AD-Export-$(get-date -f yyyy-MM-dd-hh-mm).csv"

#import the ActiveDirectory Module 
    Import-Module ActiveDirectory 


Get-ADUser -Server $SourceDomain -Filter * -Properties GivenName,Surname,SamAccountName,EmailAddress,Department,State |  
       Select-Object @{Label = "GUIAccount.Firstname";Expression = {$_.GivenName}},
                     @{Label = "GUIAccount.Lastname";Expression = {$_.Surname}},
                     @{Label = "GUIAccount.Name";Expression = {$_.SamAccountName}},
                     @{Label = "GUIAccount.EMailAddress";Expression = {$_.EmailAddress}},
                     @{Label = "GUIAccount.Department";Expression = {$_.Department}}|
	   
	   
#Export CSV report 
                  Export-Csv -Encoding UTF8 -Path $csvreportfile -NoTypeInformation     






                  Get-ADUser a013adm1 | fl *
