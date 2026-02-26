#Requires -RunAsAdministrator
#Requires -Version 5.1

if (-not (Test-Path -Path $PSScriptRoot\Office\Data\*\stream.x64.x-none.dat))
{
	Write-Information -MessageData "" -InformationAction Continue
	Write-Verbose -Message "There aren't neccessary Office files to install. Please do not move Install.ps1 from where all files are." -Verbose
}

if (-not (Test-Path -Path $PSScriptRoot\Config.xml))
{
	Write-Information -MessageData "" -InformationAction Continue
	Write-Verbose -Message "There is no Config.xml in $PSScriptRoot." -Verbose
}

Write-Information -MessageData "" -InformationAction Continue
Write-Verbose -Message "Installing downloaded Microsoft 365 components..." -Verbose
Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/configure `"$PSScriptRoot\Config.xml`"" -Wait
