<#
    .SYNOPSIS
     Download-eLuxPackagesFI
    .DESCRIPTION
     Die Funktion 'Download-eLuxPackagesFI' ladet die IDF Datei einer bestimmten eLux Version und deren Packages herunter.
    .PARAMETER IDFName
     Legt fest welche eLux IDF Release
    .PARAMETER IDFAdress
     Mit diesem Parameter kann der Default Wert überschrieben werden. (Default = 7.228.228.50)
    .PARAMETER Directory
     Mit diesem Parameter kann der Default Wert überschrieben werden. (Default = %tmp%)
    .PARAMETER eluxVersion
     Mit diesem Parameter kann der Default Wert überschrieben werden. (Default = UC_RP5)
    .EXAMPLE
     Download-eLuxPackagesFI.ps1 -IDFName FI000043
     Erstellt einen Ordner unter "%tmp%\FI000043" und speichert dort alle Dateien ab.
    .EXAMPLE
     Download-eLuxPackagesFI.ps1 -IDFName FI000043 -Directory C:\Temp
     Erstellt einen Ordner unter "C:\Temp\FI000043" und speichert dort alle Dateien ab.
    .EXAMPLE
     Download-eLuxPackagesFI.ps1 -IDFName FI000043 -Directory C:\Temp -IDFAdress 7.242.128.23 -eluxVersion UC_RL
     In diesem Beispiel wird ein andere Server (IDFAdress) mit gegeben und eine neue eLux Version (eLuxVersion)
    .NOTES
		Author: 
         Patrik Horn (PHo)
		Link:	
         http://www.makrofactory.de
         http://www.hornpa.de
		History:
         2017-03-17 - v1.04 - Bug fixing: creates folder even if he could not download files. Adding support for https urls. Some code cleanup
         2016-10-24 - v1.03 - Some code cleanup, bug fix missing gz files in UC_RP5 and switch from module to script(PHo)
         2016-10-17 - v1.02 - Support for UC_RP5 Packages (PHo)
         2015-03-09 - v1.01 - Some bug fix (PHo)
      	 2014-XX-XX - v1.00 - Script created (PHo)
#>

[CmdletBinding(SupportsShouldProcess=$False)]
param(

    [Parameter(Mandatory=$false,Position=1)]
    [string]$IDFName = 'FI00P209',

    [string]$IDFAdress = '7.228.228.50',

    [string]$Directory = "$env:TEMP",

    [string]$eluxVersion = 'UC_RP5',

    [switch]$https = $false

)

# Http or Https
IF ($https)
{
    
    Write-Verbose "Using https protocoll"
    $webprotocoll = "https"

}
else
{
    
    Write-Verbose "Using http protocoll"
    $webprotocoll = "http"

}

# URL for IDF
$URL_IDF = "$webprotocoll"+"://$IDFAdress/eluxng/$eluxVersion/$IDFName.idf"
$DL_IDF = "$Folder\IDF\$IDFName.idf"
Write-Verbose "IDF URL: $URL_IDF"
Write-Verbose "IDF Datei: $DL_IDF"

# Download IDF from HTTP
try 
{

    Invoke-WebRequest -Uri $URL_IDF | Out-Null

}
catch [System.Net.WebException]
{

    Write-Host -ForegroundColor Red "Datei wurde nicht gefunden"
    Write-Host -ForegroundColor Red "Bitte IDFName und / oder IDFAdress prüfen!"
    Write-Host -ForegroundColor Yellow "Warning: " + $Error[0].Exception
    return

}

# Check Folder
$Folder = "$Directory\$IDFName"
IF (Test-Path $Folder) 
{

    Write-Host "Information: Ordner ist vorhanden."
    New-Item -ItemType directory -Path "$Folder\IDF" -ErrorAction SilentlyContinue | Out-Null
    New-Item -ItemType directory -Path "$Folder\Packages" -ErrorAction SilentlyContinue | Out-Null

}
else
{

    Write-Host "Information: Ordner ist nicht vorhanden und wird erstellt."
    New-Item -ItemType directory -Path $Folder -ErrorAction SilentlyContinue | Out-Null
    New-Item -ItemType directory -Path "$Folder\IDF" -ErrorAction SilentlyContinue | Out-Null
    New-Item -ItemType directory -Path "$Folder\Packages" -ErrorAction SilentlyContinue | Out-Null

}

# Download IDF from HTTP
try 
{

    Invoke-WebRequest -Uri $URL_IDF -OutFile $DL_IDF

}
catch [System.Net.WebException]
{

    Write-Host -ForegroundColor Red "Datei wurde nicht gefunden"
    Write-Host -ForegroundColor Red "Bitte IDFName und / oder IDFAdress prüfen!"
    return

}

# Load IDF File 
$eLuxIDFFile = Get-Content $DL_IDF
$eLuxPackageList = @()
$eLuxPackagesEPM = @()

# Filter IDF List for EPM
$EPMList = $eLuxIDFFile -match "EPM"
$EPMList2 = $EPMList.Split(",") -match ".epm"

# Trim EPM List and load in Array
Foreach ($element in $EPMList2) 
{

    $TMP = $element.TrimStart("[EPM:")
    $TMP2 = $TMP.TrimEnd("]")
    $eLuxPackagesEPM += $TMP2

}

# Download EPM Packages
Write-Host -BackgroundColor DarkGreen -ForegroundColor White "Section: EPM Packages..."
Foreach ($element in $eLuxPackagesEPM)
{

    Write-Host "Downloading: $element"
    $URL_Package = "$webprotocoll"+"://$IDFAdress/eluxng/$eluxVersion/$element"
    $DL_Package = "$Folder\Packages\$element"
    Write-Verbose "EPM Package URL: $URL_Package"
    Write-Verbose "EPM Package Datei: $DL_Package"
    Invoke-WebRequest -Uri $URL_Package -OutFile $DL_Package

}

# Download FPM Packages
Write-Host -BackgroundColor DarkGreen -ForegroundColor White "Section: FPM Packages..."
$eLuxPackagesFPM = Get-ChildItem -Path "$Folder\Packages" -Filter *.epm
Foreach ($element in $eLuxPackagesFPM.Name)
{

    # Get Content
    $EPM_List = Get-Content -Path "$Folder\Packages\$element"
    # Filter Content 
    $FPMList = $EPM_List.Split(",") | Where-Object {($_ -match ".$eluxVersion-") -and ($_ -match ".FPM") }
    #Download FPM Packages
    Write-Host -BackgroundColor DarkGray -ForegroundColor White "Information: Lade FPM Packages für $element"
    foreach ($FPMPackages in $FPMList)
    {

        Write-Host "Downloading: $FPMPackages"
        $URL_Package = "$webprotocoll"+"://$IDFAdress/eluxng/$eluxVersion/$FPMPackages"
        $DL_Package = "$Folder\Packages\$FPMPackages"
        Write-Verbose "EPM Package URL: $URL_Package"
        Write-Verbose "EPM Package Datei: $DL_Package"
        Invoke-WebRequest -Uri $URL_Package -OutFile $DL_Package

        IF($FPMPackages -like "installrp*")
        {
            $GZPackage = $FPMPackages -replace "fpm","gz"
            Write-Host "Downloading: $GZPackage"
            $URL_Package = "$webprotocoll"+"://$IDFAdress/eluxng/$eluxVersion/$GZPackage"
            $DL_Package = "$Folder\Packages\$GZPackage"
            Write-Verbose "EPM Package URL: $URL_Package"
            Write-Verbose "EPM Package Datei: $DL_Package"
            Invoke-WebRequest -Uri $URL_Package -OutFile $DL_Package

        }

    }

}
# Summary
Write-Host -BackgroundColor DarkGreen -ForegroundColor White "Section: Summary..."
Write-Host -NoNewline "Ordner Pfad: "
Write-Host -ForegroundColor Green " $Folder"
Write-Host -NoNewline "Packages Pfad: "
Write-Host -ForegroundColor Green " $Folder\Packages"