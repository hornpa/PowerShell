<#
    .SYNOPSIS
     Enable-CitrixVDALogging
    .DESCRIPTION
     Enable Citrix VDA Logging on the System.
    .NOTES
     AUTHOR: 
      Patrik Horn (PHo)
     HISTORY:
      2016-10-20 - v1.0 - Release (Patrik Horn)
    .LINK
      http://www.hornpa.de
#>

[cmdletbinding()]

Param(
    [parameter(Mandatory=$false)]
    [String]$LogFileName = 'D:\Static\Cds\Broker.log' ,
    [parameter(Mandatory=$false)]
    [String]$LogLevel = '16' ,
    [parameter(Mandatory=$false)]
    [Switch]$OverwriteLogFile = $true ,
    [parameter(Mandatory=$false)]
    [String]$Path_Config_File = "$env:ProgramFiles\Citrix\Virtual Desktop Agent\BrokerAgent.exe.config"
)

#Functions
Function Add-XMLNode
{
    PARAM
    (
        $Key,
        $Value,
        $xml
    )

    begin{}

    process
    {
        $xml.CreateElement('add') | ForEach-Object { # Could have used a variable defenition here but wanted to disable console output with out-null
            $newAppSetting = $_
        $xml.Configuration.AppSettings.AppendChild($newAppSetting) | Out-Null}
        $newAppSetting.SetAttribute('key',"$($Key)")
        $newAppSetting.SetAttribute('value', "$($Value)")

        Return $xml
    }
    end{}
}

$Service_Name = "BrokerAgent"
Try{
    Stop-Service $Service_Name -ErrorAction Stop
    }Catch{
    Throw $_        
}

Try{
    $XML = [XML] (Get-Content -Path $Path_Config_File)
    }Catch{
    Throw $_        
}

$Node = $XML.configuration.appSettings.add
If ($node.key -eq 'LogFileName'){
    $Current_Path = $(($XML.configuration.appSettings.ChildNodes | Where-Object {$_.key -eq "LogFileName"}).Value)
    $Msg = "Logging for $Service_Name is already enabled. Nothing to do here. The current path is $Current_Path"
    Write-Warning $Msg
    Return
}

$XML = Add-XMLNode -Key 'LogFileName' -Value $LogFileName -xml $xml
$XML = Add-XMLNode -Key 'LogLevel' -Value $LogLevel -xml $xml        

IF ($OverwriteLogFile){
    $XML = Add-XMLNode -Key 'OverwriteLogFile' -Value '1' -xml $xml        
}

$XML.Save($Path_Config_File)

Try{
    Start-Service $Service_Name -ErrorAction Stop
    Write-Output "Logging for $Service_Name is now enabled. The path is $($LogFileName)."
    }Catch{
    Throw $_        
}