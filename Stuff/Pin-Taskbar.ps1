function Pin-Taskbar{ 
<#
    .SYNOPSIS
     Pin-Taskbar
	.Description
     Pin  or Unpin a Link or Exe to the Taskbar
    .PARAMETER Item
     Path to Exe or Link.
    .PARAMETER Action
     Should it Pinned or Unpinned.
    .EXAMPLE
     Pin-Taskbar -Item "C:\VRZ\StartScreen\#2#\FI-Task-Manager.lnk" -Action Pin
    .EXAMPLE
     Pin-Taskbar -Item "C:\VRZ\StartScreen\#2#\FI-Task-Manager.lnk" -Action Unpin
    .NOTES
		Author: 
		 Patrik Horn (PHo)
		Link:	
		 www.hornpa.de
		History:
      	 2016-11-15 - v1.0 - Script created (PHo)
#>

    [cmdletbinding()]

    Param(

        [parameter(Mandatory=$false)]
        [ValidateScript({Test-Path $_})] 
        [String]$Item ,

	    [parameter(Mandatory=$false)]
        [ValidateSet(“Pin”,”Unpin”)] 
        [String]$Action

    )

    $Language = (Get-Culture).Name
    Switch ($Language){
        "de-DE" {
            $Search_Pin = "An Tas&kleiste anheften"
            $Search_Unpin =  "Von Tas&kleiste lösen"
        }
        "en-US" {
            $Search_Pin = "Pin to Tas&kbar"
            $Search_Unpin =  "Unpin from Tas&kbar"
        }
        Default{
        Write-Error -Message "Language is not supported, only en-US and de-DE is supported" -ErrorAction Stop
        }
    }

    $Shell = New-Object -ComObject "Shell.Application"
    $ItemParent = Split-Path -Path $Item -Parent
    $ItemLeaf = Split-Path -Path $Item -Leaf
    $Folder = $Shell.NameSpace($ItemParent)
    $ItemObject = $Folder.ParseName($ItemLeaf)
    $Verbs = $ItemObject.Verbs()

    switch($Action){
        "Pin"   {$Verb = $Verbs | Where-Object -Property Name -EQ $Search_Pin}
        "Unpin" {$Verb = $Verbs | Where-Object -Property Name -EQ $Search_Unpin}
        default {Write-Error -Message "Invalid action, should be Pin or Unpin" -ErrorAction Stop}
    }

    IF (Test-Path -Path "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\$ItemLeaf"){
        $Msg = "Already exist"
        Write-Warning -Message $Msg
        Return $Msg
        }Else{   
        If($Verb -eq $null){
            $Msg = "That action is not currently available on this item"
            Write-Error -Message $Msg
            Return $Msg
            }else{
            $Result = $Verb.DoIt()
            $Msg = "Successfully"
            Return $Msg
        }
    }

}