<#
    .SYNOPSIS
     Disable-CitrixVDALogging
    .DESCRIPTION
     Disable Citrix VDA Logging on the System.
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
    [String]$Path_Config_File = "$env:ProgramFiles\Citrix\Virtual Desktop Agent\BrokerAgent.exe.config"
)

#Functions
Function Remove-XMLNode
{
    PARAM
    (
        $node,
        $xml
    )
    begin{}

    process
    {
        $xml.configuration.appSettings.RemoveChild($node) | Out-Null

        Return $xml
    }
    end{}
}


Try{
    $XML = [XML] (Get-Content -Path $Path_Config_File)
    }Catch{
    Throw $_        
}

$Node = $XML.configuration.appSettings.add
If (!($node.key -eq 'LogFileName')){
    $Msg = "Logging for $Service is already disabled. Nothing to do here."
    Write-Warning $Msg
    Return
}

Switch($Node.key){
    LogFileName {
        $node = $xml.configuration.appSettings.ChildNodes | Where-Object {$_.key -eq 'LogFileName'}
        $xml = Remove-XMLNode -node $node -xml $xml            
    }
    LogLevel {
        $node = $xml.configuration.appSettings.ChildNodes | Where-Object {$_.key -eq 'LogLevel'}
        $xml = Remove-XMLNode -node $node -xml $xml            
    }
    OverwriteLogFile {
        $node = $xml.configuration.appSettings.ChildNodes | Where-Object {$_.key -eq 'OverwriteLogFile'}
        $xml = Remove-XMLNode -node $node -xml $xml 
    }
}

$XML.Save($Path_Config_File)
