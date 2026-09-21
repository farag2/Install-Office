<#
	.SYNOPSIS
	Download Microsoft Office 2024 and Microsoft 365

	.PARAMETER Branch
	ProPlus2024Volume for Microsoft Office 2024

	.PARAMETER Branch
	O365ProPlusRetail for Microsoft 365

	.PARAMETER Channel
	PerpetualVL2024 for Microsoft Office 2024

	.PARAMETER Channel
	Current for Microsoft 365

	.PARAMETER Channel
	SemiAnnual for Microsoft 365

	.PARAMETER Components
	Access, OneDrive, Outlook, Word, Excel, PowerPoint, Teams, OneNote, Publisher, Project 2024, Visio 2024

	.EXAMPLE Download Office 2024 with the Excel, Word components
	Download.ps1 -Branch ProPlus2024Volume -Channel PerpetualVL2024 -Components Excel, Word

	.EXAMPLE Download Office 365 with the Excel, Word, PowerPoint components
	Download.ps1 -Branch O365ProPlusRetail -Channel Current -Components Excel, OneDrive, Outlook, PowerPoint, Teams, Word

	.NOTES
	The Current and SemiAnnual channels are valid for O365ProPlusRetail only
	The PerpetualVL2024 channel is valid for ProPlus2024Volume only

	.LINK
	https://config.office.com/deploymentsettings
#>
[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true)]
	[ValidateSet("ProPlus2024Volume", "O365ProPlusRetail")]
	[string]
	$Branch,

	[Parameter(Mandatory = $true)]
	[ValidateSet("Current", "PerpetualVL2024", "SemiAnnual")]
	[string]
	$Channel,

	[Parameter(Mandatory = $true)]
	[ValidateSet("Access", "OneDrive", "Outlook", "Word", "Excel", "OneNote", "Publisher", "PowerPoint", "Teams", "ProjectPro2024Volume", "VisioPro2024Volume")]
	[string[]]
	$Components
)

#Requires -Version 5.1

# Channels allowed for every branch
# ValidateSet cannot depend on another parameter's value, so the pair is checked here
$ValidChannels = @{
	ProPlus2024Volume = @("PerpetualVL2024")
	O365ProPlusRetail = @("Current", "SemiAnnual")
}

if ($ValidChannels[$Branch] -notcontains $Channel)
{
	Write-Information -MessageData "" -InformationAction Continue
	Write-Warning -Message "The `"$Channel`" channel cannot be used with `"$Branch`". Available: $($ValidChannels[$Branch] -join ", ")"
	exit
}

if (-not (Test-Path -Path "$PSScriptRoot\Default.xml"))
{
	Write-Information -MessageData "" -InformationAction Continue
	Write-Warning -Message "Default.xml doesn't exist"
	exit
}

if ($PSVersionTable.PSVersion.Major -eq 5)
{
	# Progress bar can significantly impact cmdlet performance
	# https://github.com/PowerShell/PowerShell/issues/2138
	$Script:ProgressPreference = "SilentlyContinue"

	[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
}

[xml]$Config = Get-Content -Path "$PSScriptRoot\Default.xml" -Encoding Default -Force

($Config.Configuration.Add.Product | Where-Object -FilterScript {$_.ID -eq ""}).ID = $Branch
$Config.Configuration.Add.Channel = $Channel

foreach ($Component in $Components)
{
	switch ($Component)
	{
		{$_ -in @("Access", "Excel", "OneNote", "Outlook", "PowerPoint", "Publisher", "Word")}
		{
			# Default.xml may have been edited manually, so the node is not guaranteed to exist
			$Node = $Config.SelectSingleNode("//ExcludeApp[@ID='$Component']")
			$Node.ParentNode.RemoveChild($Node)
		}
		{$_ -in @("ProjectPro2024Volume", "VisioPro2024Volume")}
		{
			$Product = $Config.Configuration.Add.AppendChild($Config.CreateElement("Product"))
			$Product.SetAttribute("ID", $Component)

			$Language = $Product.AppendChild($Config.CreateElement("Language"))
			$Language.SetAttribute("ID", "MatchOS")
		}
		"OneDrive"
		{
			$OneDrive = Get-Package -Name "Microsoft OneDrive" -ProviderName Programs -Force -ErrorAction Ignore
			if (-not $OneDrive)
			{
				$OneDriveSetup = switch ((Get-CimInstance -ClassName Win32_OperatingSystem).Caption)
				{
					{$_ -match 10}
					{
						"$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
						break
					}
					{$_ -match 11}
					{
						"$env:SystemRoot\System32\OneDriveSetup.exe"
						break
					}
				}

				if ($OneDriveSetup -and (Test-Path -Path $OneDriveSetup))
				{
					Write-Information -MessageData "" -InformationAction Continue
					Write-Verbose -Message "OneDrive Installing" -Verbose

					Start-Process -FilePath $OneDriveSetup
				}
				else
				{
					Write-Information -MessageData "" -InformationAction Continue
					Write-Verbose -Message "OneDrive Downloading" -Verbose

					try
					{
						# Parse XML to get the URL
						# https://go.microsoft.com/fwlink/p/?LinkID=844652
						$Parameters = @{
							Uri             = "https://g.live.com/1rewlive5skydrive/OneDriveProductionV2"
							UseBasicParsing = $true
							Verbose         = $true
						}
						$OneDriveURL = (Invoke-RestMethod @Parameters).root.update.amd64binary.url | Select-Object -Index 1
					}
					catch
					{
						Write-Information -MessageData "" -InformationAction Continue
						Write-Verbose -Message "Connection could not be established with https://g.live.com/1rewlive5skydrive/OneDriveProductionV2" -Verbose

						exit
					}

					try
					{
						$Parameters = @{
							Uri             = $OneDriveURL
							OutFile         = "$PSScriptRoot\OneDriveSetup.exe"
							UseBasicParsing = $true
							Verbose         = $true
						}
						Invoke-WebRequest @Parameters
					}
					catch
					{
						Write-Information -MessageData "" -InformationAction Continue
						Write-Verbose -Message "Connection could not be established with https://g.live.com/1rewlive5skydrive/OneDriveProductionV2" -Verbose

						exit
					}

					Write-Information -MessageData "" -InformationAction Continue
					Write-Verbose -Message "OneDrive was downloaded to $PSScriptRoot" -Verbose
				}
			}
		}
		"Teams"
		{
			Write-Information -MessageData "" -InformationAction Continue
			Write-Verbose -Message "Teams Downloading" -Verbose

			try
			{
				# https://www.microsoft.com/microsoft-teams/download-app
				$Parameters = @{
					Uri             = "https://statics.teams.cdn.office.net/production-windows-x86/lkg/MSTeamsSetup.exe"
					OutFile         = "$PSScriptRoot\MSTeamsSetup.exe"
					UseBasicParsing = $true
					Verbose         = $true
				}
				Invoke-WebRequest @Parameters
			}
			catch
			{
				Write-Information -MessageData "" -InformationAction Continue
				Write-Verbose -Message "Connection could not be established with https://statics.teams.cdn.office.net/production-windows-x86/lkg/MSTeamsSetup.exe" -Verbose

				exit
			}

			Write-Information -MessageData "" -InformationAction Continue
			Write-Verbose -Message "Teams was downloaded to $PSScriptRoot" -Verbose
		}
	}
}

$Config.Save("$PSScriptRoot\Config.xml")

# Microsoft blocks Russian and Belarusian regions for Office downloading
# https://docs.microsoft.com/en-us/windows/win32/intl/table-of-geographical-locations
# https://en.wikipedia.org/wiki/2022_Russian_invasion_of_Ukraine
$HomeLocation = (Get-WinHomeLocation).GeoId
if ($HomeLocation -in @(203, 29))
{
	# Set to Poland
	$Script:Region = $HomeLocation
	Set-WinHomeLocation -GeoId 191

	Write-Information -MessageData "" -InformationAction Continue
	Write-Warning -Message "Region changed to Poland"

	$Script:RegionChanged = $true
}

# It is needed to remove these keys to bypass Russian and Belarusian region blocks
$Paths = @(
	"HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Experiment",
	"HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs",
	"HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentEcs"
)
Remove-Item -Path $Paths -Recurse -Force -ErrorAction Ignore

# Download Office Deployment Tool
# https://www.microsoft.com/en-us/download/details.aspx?id=49117
if (-not (Test-Path -Path "$PSScriptRoot\setup.exe"))
{
	try
	{
		$Parameters = @{
			Uri             = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
			OutFile         = "$PSScriptRoot\setup.exe"
			UseBasicParsing = $true
			Verbose         = $true
		}
		Invoke-WebRequest @Parameters
	}
	catch
	{
		Write-Information -MessageData "" -InformationAction Continue
		Write-Verbose -Message "Connection could not be established with https://officecdn.microsoft.com/pr/wsus/setup.exe" -Verbose

		exit
	}
}

Write-Information -MessageData "" -InformationAction Continue
Write-Verbose -Message "Downloading... Please do not close any console windows." -Verbose

# Start downloading to the Office folder
Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/download `"$PSScriptRoot\Config.xml`"" -Wait

if ($Script:RegionChanged)
{
	# Set to original region ID
	Set-WinHomeLocation -GeoId $Script:Region

	Write-Information -MessageData "" -InformationAction Continue
	Write-Warning -Message "Region changed to original one"
}

Write-Information -MessageData "" -InformationAction Continue
Write-Verbose -Message "Office downloaded. Please run `"$PSScriptRoot\Install.ps1`" file with administrator privileges." -Verbose
