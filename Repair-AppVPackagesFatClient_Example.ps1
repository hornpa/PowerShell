Get-AppvClientPackage -All | Remove-AppvClientPackage
Get-AppvClientPackage | Remove-AppvClientPackage
Sync-AppvPublishingServer 1 -Force

Get-AppvClientPackage | FT Name, PackageId,VersionId