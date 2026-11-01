<#
	.SYNOPSIS
	Remove Microsoft Office completely with all leftovers

	.DESCRIPTION
	Office 2007-2024 and Microsoft 365 Apps: Click-to-Run, Windows Installer and Microsoft Store editions,
	including Visio, Project, language packs, proofing tools and Office runtimes (Access Database Engine, Filter Pack)

	Every known location is cleaned blindly: nothing is checked beforehand, every error is swallowed, nothing is logged
	Kept: documents and Outlook data files (*.pst)

	Exit codes: 0 — done, 3010 — done, restart required

	.EXAMPLE
	.\Remove-Office.ps1
#>
#Requires -Version 5.1
#Requires -RunAsAdministrator

Write-Verbose -Message "Do you want to completely uninstall Microsoft Office/Microsoft 365 and remove any traces?" -Verbose
pause

$ErrorActionPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$ConfirmPreference = "None"

# Terminating errors (.NET exceptions, COM, parameter binding) do not stop the script either
trap
{
	continue
}

#region Functions
# Path, command line or registry data that points into Office
function Test-OfficePath
{
	param
	(
		[Parameter(Mandatory = $false)]
		[string]
		$Path
	)

	if ([string]::IsNullOrEmpty($Path))
	{
		return $false
	}

	$Path = [Environment]::ExpandEnvironmentVariables($Path)
	$Patterns = @(
		"\Program Files\Microsoft Office",
		"\Program Files (x86)\Microsoft Office",
		"\Microsoft Shared\ClickToRun",
		"\Microsoft Shared\OFFICE1",
		"\Microsoft Shared\OfficeSoftwareProtectionPlatform",
		"\Microsoft Shared\Source Engine",
		"\root\Office1",
		"Microsoft Office\root\",
		"Microsoft Office 15",
		"Microsoft Office 16",
		# Windows Installer cache of an Office product: %SystemRoot%\Installer\{90160000-...-0000000FF1CE}\wordicon.exe
		"0000000FF1CE}"
	)
	foreach ($Pattern in $Patterns)
	{
		if ($Path.IndexOf($Pattern, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
		{
			return $true
		}
	}

	return $false
}

# Windows Installer packed GUID (32 chars) to the regular {GUID} form
function ConvertFrom-PackedGuid
{
	param
	(
		[Parameter(Mandatory = $true)]
		[string]
		$PackedGuid
	)

	if ($PackedGuid.Length -ne 32)
	{
		return $null
	}

	# The first three groups are reversed, every byte of the last 8 bytes has its nibbles swapped
	$Chars = $PackedGuid.ToUpperInvariant().ToCharArray()
	$Data4 = for ($i = 16; $i -lt 32; $i += 2)
	{
		$Chars[$i + 1]
		$Chars[$i]
	}
	$Data4 = -join $Data4

	return "{{{0}-{1}-{2}-{3}-{4}}}" -f (-join $Chars[7..0]), (-join $Chars[11..8]), (-join $Chars[15..12]), $Data4.Substring(0, 4), $Data4.Substring(4)
}

# Darwin descriptor (base85 "squished" GUID, first 20 chars) to the regular {GUID} form
function ConvertFrom-SquishedGuid
{
	param
	(
		[Parameter(Mandatory = $true)]
		[string]
		$SquishedGuid
	)

	if ($SquishedGuid.Length -lt 20)
	{
		return $null
	}

	$Alphabet = '!$%&''()*+,-.0123456789=?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[]^_`abcdefghijklmnopqrstuvwxyz{}~'
	$Bytes = New-Object -TypeName byte[] -ArgumentList 16

	for ($Group = 0; $Group -lt 4; $Group++)
	{
		[uint64]$Value = 0
		[uint64]$Power = 1
		for ($i = 0; $i -lt 5; $i++)
		{
			# The original script indexed its 128-item table before checking the char code and crashed on non-ASCII chars
			$Index = $Alphabet.IndexOf($SquishedGuid[$Group * 5 + $i])
			if ($Index -lt 0)
			{
				return $null
			}
			$Value += [uint64]$Index * $Power
			$Power *= 85
		}

		if ($Value -gt [uint32]::MaxValue)
		{
			return $null
		}
		[System.BitConverter]::GetBytes([uint32]$Value).CopyTo($Bytes, $Group * 4)
	}

	return ([guid]::new($Bytes)).ToString("B").ToUpperInvariant()
}

# Any Office 2007+ product code. Accepts both {GUID} and packed forms
function Test-OfficeProductCode
{
	param
	(
		[Parameter(Mandatory = $false)]
		[string]
		$ProductCode
	)

	if ($ProductCode.Length -eq 32)
	{
		# Cheap prefilter: packed "-0000000FF1CE}" always ends with "000000F01FEC"
		if ((-not $ProductCode.EndsWith("000000F01FEC", [System.StringComparison]::OrdinalIgnoreCase)) -and (-not $ExplicitPackedCodes.Contains($ProductCode)))
		{
			return $false
		}
		$ProductCode = ConvertFrom-PackedGuid -PackedGuid $ProductCode
	}

	if ($ProductCode.Length -ne 38)
	{
		return $false
	}
	$ProductCode = $ProductCode.ToUpperInvariant()

	# Checked first: in the original script these checks were unreachable for most GUIDs because of an early Exit Function
	if ($ExplicitProductCodes -contains $ProductCode)
	{
		return $true
	}

	if (-not $ProductCode.EndsWith("0000000FF1CE}"))
	{
		return $false
	}

	# Major version: 12 — Office 2007, 16 — Office 2016 and newer
	$Major = 0
	return [int]::TryParse($ProductCode.Substring(3, 2), [ref]$Major) -and ($Major -ge 12)
}

# Close everything that runs from Office folders
function Stop-OfficeProcess
{
	param
	(
		[Parameter(Mandatory = $false)]
		[string[]]
		$ExtraName
	)

	$Names = $OfficeProcesses + $ExtraName
	$Processes = @(Get-CimInstance -ClassName Win32_Process | Where-Object -FilterScript {($_.ProcessId -ne $PID) -and (($Names -contains $_.Name) -or (Test-OfficePath -Path $_.ExecutablePath))})
	foreach ($Process in $Processes)
	{
		Stop-Process -Id $Process.ProcessId -Force
	}

	if ($Processes)
	{
		Write-Verbose -Message "Closed: $(($Processes.Name | Sort-Object -Unique) -join ', ')" -Verbose

		# Let the processes release file handles
		Start-Sleep -Seconds 3
	}
}

# -Name of Remove-ItemProperty expands wildcards, so the name is escaped. The default value is "(default)"
function Remove-RegistryValue
{
	param
	(
		[Parameter(Mandatory = $true)]
		[string]
		$Path,

		[Parameter(Mandatory = $true)]
		[AllowEmptyString()]
		[string]
		$Name
	)

	if ($Name -eq "")
	{
		$Name = "(default)"
	}
	else
	{
		$Name = [WildcardPattern]::Escape($Name)
	}

	Remove-ItemProperty -LiteralPath $Path -Name $Name -Force
}
#endregion Functions

#region Initialization
$WorkFolder = "$env:TEMP\RemoveOffice"
New-Item -Path $WorkFolder -ItemType Directory -Force | Out-Null

$RebootRequired = $false

# MoveFileEx for files in use
$Signature = @{
	Namespace        = "RemoveOffice"
	Name             = "NativeMethods"
	Language         = "CSharp"
	MemberDefinition = @"
[DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
public static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, int dwFlags);
"@
}
if (-not ("RemoveOffice.NativeMethods" -as [type]))
{
	Add-Type @Signature
}

$OfficeProcesses = @(
	"appvshnotify.exe",
	"integratedoffice.exe",
	"integrator.exe",
	"firstrun.exe",
	"communicator.exe",
	"lync.exe",
	"msosync.exe",
	"OneNoteM.exe",
	"mavinject32.exe",
	"perfboost.exe",
	"roamingoffice.exe",
	"officeclicktorun.exe",
	"officeondemand.exe",
	"OfficeC2RClient.exe",
	"ose.exe",
	"ose64.exe",
	"osppsvc.exe"
)

$OfficeExecutables = @(
	"WINWORD.EXE",
	"EXCEL.EXE",
	"POWERPNT.EXE",
	"OUTLOOK.EXE",
	"MSACCESS.EXE",
	"MSPUB.EXE",
	"ONENOTE.EXE",
	"ONENOTEM.EXE",
	"VISIO.EXE",
	"WINPROJ.EXE",
	"lync.exe",
	"GROOVE.EXE",
	"INFOPATH.EXE",
	"SPDESIGN.EXE",
	"MSOSYNC.EXE",
	"OfficeClickToRun.exe",
	"OfficeC2RClient.exe",
	"msoia.exe",
	"SETLANG.EXE"
)

$UninstallPaths = @(
	"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
	"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

$UserProfiles = @(Get-CimInstance -ClassName Win32_UserProfile -Filter "Special = FALSE" | Where-Object -FilterScript {$_.LocalPath})
#endregion Initialization

#region Detection
Write-Verbose -Message "Detecting installed Office" -Verbose

# Invalid subkeys break Windows Installer API enumeration used by msiexec later
$Locations = @(
	@{Hive = [Microsoft.Win32.Registry]::LocalMachine; SubKey = "SOFTWARE\Classes\Installer\Products"},
	@{Hive = [Microsoft.Win32.Registry]::LocalMachine; SubKey = "SOFTWARE\Classes\Installer\Components"},
	@{Hive = [Microsoft.Win32.Registry]::LocalMachine; SubKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"},
	@{Hive = [Microsoft.Win32.Registry]::LocalMachine; SubKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components"}
)
foreach ($Location in $Locations)
{
	$Key = $Location.Hive.OpenSubKey($Location.SubKey)
	if (-not $Key)
	{
		continue
	}

	foreach ($Name in $Key.GetSubKeyNames())
	{
		if ($Name.Length -ne 32)
		{
			Remove-Item -LiteralPath "Registry::$($Key.Name)\$Name" -Recurse -Force
		}
	}
	$Key.Close()
}

# Click-to-Run packages. The original script read the 16.0 and root values into one variable and then overwrote it
$ClickToRunRoots = @(
	@{
		Version = 16
		Root    = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun"
	},
	@{
		Version = 16
		Root    = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\ClickToRun"
	},
	@{
		Version = 15
		Root    = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun"
	}
)

$ClickToRunVersion = $null
$Packages = @()
foreach ($ClickToRunRoot in $ClickToRunRoots)
{
	$PackageFolder = [Microsoft.Win32.Registry]::GetValue($ClickToRunRoot.Root, "PackageFolder", $null)
	if (-not $PackageFolder)
	{
		continue
	}

	$Packages += [PSCustomObject]@{
		Folder = $PackageFolder
		Guid   = [Microsoft.Win32.Registry]::GetValue($ClickToRunRoot.Root, "PackageGUID", $null)
	}

	if (-not $ClickToRunVersion)
	{
		$ClickToRunVersion = $ClickToRunRoot.Version
	}
}

# Default package folders are tried blindly too: the registration may already be gone
$Packages += @(
	[PSCustomObject]@{Folder = "$env:ProgramFiles\Microsoft Office"; Guid = $null},
	[PSCustomObject]@{Folder = "$env:ProgramFiles\Microsoft Office 15"; Guid = $null},
	[PSCustomObject]@{Folder = "$env:ProgramFiles\Microsoft Office 16"; Guid = $null},
	[PSCustomObject]@{Folder = "${env:ProgramFiles(x86)}\Microsoft Office"; Guid = $null}
)

# Click-to-Run unified uninstall command
$ArpUninstallString = $null
foreach ($UninstallKey in (Get-ChildItem -LiteralPath $UninstallPaths))
{
	$UninstallString = [string]$UninstallKey.GetValue("UninstallString")
	if ($UninstallString.IndexOf("OfficeClickToRun.exe", [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
	{
		$ArpUninstallString = $UninstallString
		Write-Verbose -Message "Found: $($UninstallKey.GetValue('DisplayName')) $($UninstallKey.GetValue('DisplayVersion'))" -Verbose
	}
}

if ($ClickToRunVersion)
{
	Write-Verbose -Message "Click-to-Run $ClickToRunVersion.0 is registered" -Verbose
}
else
{
	Write-Verbose -Message "No Click-to-Run registration found" -Verbose
}

# Product codes removed regardless of the Office pattern
$ExplicitProductCodes = @(
	# Microsoft Online Services Sign-in Assistant x64 and x86
	"{6C1ADE97-24E1-4AE4-AEDD-86D3A209CE60}",
	"{9520DDEB-237A-41DB-AA20-F2EF2360DCEB}",
	# Office 2016 Click-to-Run package
	"{9AC08E99-230B-47E8-9721-4577B7F124EA}"
)
foreach ($Package in $Packages)
{
	if ($Package.Guid)
	{
		# PackageGUID is stored without braces: the original script compared it with a braced product code and never matched
		$ExplicitProductCodes += "{$($Package.Guid.Trim('{}').ToUpperInvariant())}"
	}
}

$ExplicitPackedCodes = New-Object -TypeName System.Collections.Generic.HashSet[string] -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($ProductCode in $ExplicitProductCodes)
{
	# The packing permutation is its own inverse: applied to the plain 32 hex digits it returns the packed form
	$Packed = ConvertFrom-PackedGuid -PackedGuid $ProductCode.Trim("{}").Replace("-", "")
	[void]$ExplicitPackedCodes.Add($Packed.Trim("{}").Replace("-", ""))
}
#endregion Detection

#region Uninstall
Write-Verbose -Message "Closing Office applications" -Verbose
Stop-OfficeProcess

Write-Verbose -Message "Removing scheduled tasks" -Verbose

# Scheduled tasks: the whole \Microsoft\Office\ folder and the root-level tasks of older builds
$Tasks = @(
	"\FF_INTEGRATEDstreamSchedule",
	"\FF_INTEGRATEDUPDATEDETECTION",
	"\C2RAppVLoggingStart",
	"\Office 15 Subscription Heartbeat",
	"\Microsoft Office 15 Sync Maintenance for {d068b555-9700-40b8-992c-f866287b06c1}",
	"\Office Background Streaming",
	"\Office Subscription Maintenance"
)
Get-ScheduledTask | Where-Object -FilterScript {($_.TaskPath -eq "\Microsoft\Office\") -or ($Tasks -contains "$($_.TaskPath)$($_.TaskName)")} | Unregister-ScheduledTask -Confirm:$false

$Scheduler = New-Object -ComObject Schedule.Service
$Scheduler.Connect()
$Scheduler.GetFolder("\Microsoft").DeleteFolder("Office", 0)

Write-Verbose -Message "Removing Office product keys" -Verbose

# Product keys of Office 2010 (OSPP) and Office 2013+ (SPP or OSPP)
Start-Process -FilePath "$PSScriptRoot\x64\CleanOSPP.exe" -Wait -WindowStyle Hidden

foreach ($ClassName in @("SoftwareLicensingProduct", "OfficeSoftwareProtectionProduct"))
{
	$Parameters = @{
		ClassName = $ClassName
		Filter    = "(ApplicationId = '0ff1ce15-a989-479d-af46-f275c6370663' OR ApplicationId = '59a52881-a989-479d-af46-f275c6370663') AND PartialProductKey IS NOT NULL"
	}
	foreach ($LicenseProduct in (Get-CimInstance @Parameters))
	{
		Write-Verbose -Message "Product key: $($LicenseProduct.Name)" -Verbose

		# UninstallProductKey() takes no arguments. The original script passed ProductKeyID and the error was swallowed
		$null = Invoke-CimMethod -InputObject $LicenseProduct -MethodName UninstallProductKey
	}
}

Write-Verbose -Message "Removing Office sign-in tokens from Credential Manager" -Verbose

# Office sign-in tokens in Credential Manager. Available for the account the script runs under only
foreach ($Line in (& "$env:SystemRoot\System32\cmdkey.exe" /list))
{
	# "Target: LegacyGeneric:target=MicrosoftOffice16_Data:SSPI:user@contoso.com". The label is localized, the target is not
	$Index = $Line.IndexOf("MicrosoftOffice1", [System.StringComparison]::OrdinalIgnoreCase)
	if ($Index -lt 0)
	{
		continue
	}

	$Target = $Line.Substring($Index).Trim()
	& "$env:SystemRoot\System32\cmdkey.exe" /delete:$Target | Out-Null
	& "$env:SystemRoot\System32\cmdkey.exe" /delete:"LegacyGeneric:target=$Target" | Out-Null
}

Write-Verbose -Message "Removing Microsoft Store editions of Office" -Verbose

# Microsoft Store editions of Office and its companion apps. Dependent packages go before the main one
$StorePackages = @(
	"Microsoft.Office.Desktop*",
	"Microsoft.MicrosoftOfficeHub",
	"Microsoft.Office.OneNote"
)
$ProvisionedPackages = @(Get-AppxProvisionedPackage -Online)
foreach ($StorePackage in $StorePackages)
{
	Get-AppxPackage -AllUsers -Name $StorePackage | Sort-Object -Property {$_.Name.Length} -Descending | ForEach-Object -Process {
		Write-Verbose -Message "Store package: $($_.PackageFullName)" -Verbose
		Remove-AppxPackage -Package $_.PackageFullName -AllUsers
	}

	$ProvisionedPackages | Where-Object -FilterScript {$_.DisplayName -like $StorePackage} | ForEach-Object -Process {
		Write-Verbose -Message "Provisioned package: $($_.PackageName)" -Verbose
		Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName | Out-Null
	}
}

# Office Source Engine must be startable and run as LocalSystem for msiexec based removal
# The original script called ChangeStartMode() and Change() unconditionally: VBScript does not short-circuit And
foreach ($Service in (Get-CimInstance -ClassName Win32_Service -Filter "Name LIKE 'ose%'"))
{
	if ($Service.StartMode -eq "Disabled")
	{
		Set-Service -Name $Service.Name -StartupType Manual
	}

	if ($Service.StartName -ne "LocalSystem")
	{
		$null = Invoke-CimMethod -InputObject $Service -MethodName Change -Arguments @{StartName = "LocalSystem"; StartPassword = ""}
	}
}

# Click-to-Run via Office Deployment Tool: local ODT16\ODT15 next to the script, otherwise a signed download (Office 2016+)
$OdtSetup = $null
if ($ClickToRunVersion -eq 15)
{
	# Office 2013 needs the 2013 Office Deployment Tool. Its public download link is dead
	if (Get-Item -LiteralPath "$PSScriptRoot\ODT15\setup.exe")
	{
		$OdtSetup = "$PSScriptRoot\ODT15\setup.exe"
	}
}
elseif ($ClickToRunVersion -eq 16)
{
	if (Get-Item -LiteralPath "$PSScriptRoot\ODT16\setup.exe")
	{
		$OdtSetup = "$PSScriptRoot\ODT16\setup.exe"
	}
	else
	{
		Write-Verbose -Message "Downloading Office Deployment Tool" -Verbose

		# Windows PowerShell 5.1 does not enable TLS 1.2 by default on older systems
		[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

		# The original script downloaded over plain HTTP and ran the file elevated without any check
		$Parameters = @{
			Uri             = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
			OutFile         = "$WorkFolder\setup.exe"
			UseBasicParsing = $true
		}
		Invoke-WebRequest @Parameters

		$Signature = Get-AuthenticodeSignature -FilePath "$WorkFolder\setup.exe"
		if (($Signature.Status -eq "Valid") -and $Signature.SignerCertificate.Subject.Contains("O=Microsoft Corporation"))
		{
			$OdtSetup = "$WorkFolder\setup.exe"
		}
	}
}

if ($OdtSetup)
{
	$ConfigurationXml = @"
<Configuration>
	<Remove All="TRUE" />
	<Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
	Set-Content -LiteralPath "$WorkFolder\RemoveAll.xml" -Value $ConfigurationXml -Encoding UTF8 -Force

	Write-Verbose -Message "Uninstalling Click-to-Run via Office Deployment Tool. It may take up to 15 minutes" -Verbose
	Start-Process -FilePath $OdtSetup -ArgumentList "/configure `"$WorkFolder\RemoveAll.xml`"" -Wait -WindowStyle Hidden
}
elseif ($ArpUninstallString)
{
	# "C:\...\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP sourcetype=None productstoremove=... culture=en-us
	# The original script failed when productstoremove= was not found
	$ExeEnd = $ArpUninstallString.IndexOf("OfficeClickToRun.exe", [System.StringComparison]::OrdinalIgnoreCase) + "OfficeClickToRun.exe".Length
	$FilePath = $ArpUninstallString.Substring(0, $ExeEnd).Trim([char[]]'" ')
	$Arguments = foreach ($Token in $ArpUninstallString.Substring($ExeEnd).Trim([char[]]'" ').Split([char[]]" ", [System.StringSplitOptions]::RemoveEmptyEntries))
	{
		if (-not $Token.StartsWith("productstoremove=", [System.StringComparison]::OrdinalIgnoreCase))
		{
			$Token
		}
	}
	$Arguments = @($Arguments) + "productstoremove=AllProducts" + "displaylevel=False"

	Write-Verbose -Message "Uninstalling Click-to-Run via OfficeClickToRun.exe. It may take up to 15 minutes" -Verbose
	Start-Process -FilePath $FilePath -ArgumentList $Arguments -Wait -WindowStyle Hidden
}

Write-Verbose -Message "Unregistering Click-to-Run integration" -Verbose

# Published component registration of the Click-to-Run packages
# The original script tested a stale variable instead of the package folder, so this practically never ran
foreach ($Package in $Packages)
{
	Remove-Item -Path "$($Package.Folder)\root\Integration\C2RManifest*.xml" -Force

	if ($Package.Guid)
	{
		Start-Process -FilePath "$($Package.Folder)\root\Integration\integrator.exe" -ArgumentList "/U /Extension PackageRoot=`"$($Package.Folder)\root`" PackageGUID=$($Package.Guid)" -Wait -WindowStyle Hidden
		Start-Process -FilePath "$env:ProgramData\Microsoft\ClickToRun\{$($Package.Guid)}\integrator.exe" -ArgumentList "/U /Extension PackageRoot=`"$($Package.Folder)\root`" PackageGUID=$($Package.Guid)" -Wait -WindowStyle Hidden
	}
	Start-Process -FilePath "$($Package.Folder)\root\Integration\integrator.exe" -ArgumentList "/U" -Wait -WindowStyle Hidden
}

Write-Verbose -Message "Searching Windows Installer products of Office" -Verbose

# Every Office 2007+ Windows Installer product: suites, Visio, Project, language packs, proofing tools, Click-to-Run integration components
$MsiProducts = @()
$InstallerProducts = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey("SOFTWARE\Classes\Installer\Products")
if ($InstallerProducts)
{
	foreach ($Name in $InstallerProducts.GetSubKeyNames())
	{
		if (Test-OfficeProductCode -ProductCode $Name)
		{
			$MsiProducts += ConvertFrom-PackedGuid -PackedGuid $Name
		}
	}
	$InstallerProducts.Close()
}

if ($MsiProducts)
{
	Stop-OfficeProcess

	$Count = 0
	foreach ($ProductCode in $MsiProducts)
	{
		$Count++
		Write-Verbose -Message "msiexec /x $ProductCode ($Count/$($MsiProducts.Count))" -Verbose

		$Process = Start-Process -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList "/x $ProductCode REBOOT=ReallySuppress NOREMOVESPAWN=True /qn" -Wait -PassThru
		if (@(1641, 3010) -contains $Process.ExitCode)
		{
			$RebootRequired = $true
		}
	}

	Stop-Service -Name msiserver -Force
}
else
{
	Write-Verbose -Message "No Windows Installer products of Office found" -Verbose
}

Write-Verbose -Message "Removing Office services" -Verbose

# Services
foreach ($ServiceName in @("OfficeSvc", "ClickToRunSvc", "ose", "ose64", "osppsvc"))
{
	$Service = Get-CimInstance -ClassName Win32_Service -Filter "Name = '$ServiceName'"
	if (-not $Service)
	{
		continue
	}
	Write-Verbose -Message "Service: $ServiceName" -Verbose

	Stop-Service -Name $ServiceName -Force
	if ($Service.ProcessId)
	{
		Stop-Process -Id $Service.ProcessId -Force
	}

	# Remove-Service exists in PowerShell 6+ only. 1072: ERROR_SERVICE_MARKED_FOR_DELETE
	& "$env:SystemRoot\System32\sc.exe" delete $ServiceName | Out-Null
	if ($LASTEXITCODE -eq 1072)
	{
		$RebootRequired = $true
	}
}
#endregion Uninstall

#region Registry
Stop-OfficeProcess

Write-Verbose -Message "Loading registry hives of users who are not logged on" -Verbose

# Hives of users who are not logged on are loaded for the time of the registry cleanup only
$UserHives = @()
foreach ($UserProfile in $UserProfiles)
{
	if ($UserProfile.Loaded)
	{
		$UserHives += [PSCustomObject]@{
			Key    = $UserProfile.SID
			Unload = $false
		}
		continue
	}

	& "$env:SystemRoot\System32\reg.exe" load "HKU\RemoveOffice_$($UserProfile.SID)" "$($UserProfile.LocalPath)\NTUSER.DAT" 2>$null | Out-Null
	if ($LASTEXITCODE -eq 0)
	{
		$UserHives += [PSCustomObject]@{
			Key    = "RemoveOffice_$($UserProfile.SID)"
			Unload = $true
		}
	}
}

try
{
	Write-Verbose -Message "Removing Office registry keys of the computer and $($UserHives.Count) user(s)" -Verbose

	# Keys removed entirely
	$Keys = @(
		"HKLM:\SOFTWARE\Microsoft\Office",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office",
		"HKLM:\SOFTWARE\Policies\Microsoft\Office",
		"HKLM:\SOFTWARE\Policies\Microsoft\Cloud\Office",
		"HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\OfficeSoftwareProtectionPlatform",
		# Outlook as a mail, calendar and contacts client
		"HKLM:\SOFTWARE\Clients\Mail\Microsoft Outlook",
		"HKLM:\SOFTWARE\Clients\Calendar\Microsoft Outlook",
		"HKLM:\SOFTWARE\Clients\Contacts\Microsoft Outlook",
		"HKLM:\SOFTWARE\Clients\News\Microsoft Outlook",
		"HKLM:\SOFTWARE\WOW6432Node\Clients\Mail\Microsoft Outlook",
		"HKLM:\SOFTWARE\WOW6432Node\Clients\Calendar\Microsoft Outlook",
		"HKLM:\SOFTWARE\WOW6432Node\Clients\Contacts\Microsoft Outlook",
		"HKLM:\SOFTWARE\WOW6432Node\Clients\News\Microsoft Outlook",
		# Shell integration
		"HKLM:\SOFTWARE\Classes\Protocols\Handler\osf",
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)",
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)",
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)",
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}",
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}",
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}",
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}",
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{B28AA736-876B-46DA-B3A8-84C5E30BA492}",
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\NetworkNeighborhood\NameSpace\{46137B78-0EC3-426D-8B89-FF7C3A458B5E}",
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Microsoft Office Temp Files"
	)
	foreach ($UserHive in $UserHives)
	{
		$Keys += @(
			"Registry::HKEY_USERS\$($UserHive.Key)\Software\Microsoft\Office",
			"Registry::HKEY_USERS\$($UserHive.Key)\Software\Wow6432Node\Microsoft\Office",
			"Registry::HKEY_USERS\$($UserHive.Key)\Software\Policies\Microsoft\Office",
			"Registry::HKEY_USERS\$($UserHive.Key)\Software\Policies\Microsoft\Cloud\Office"
		)
	}
	Remove-Item -LiteralPath $Keys -Recurse -Force

	# Value names are paths into Office
	$NameKeys = @(
		"HKLM:\SOFTWARE\Microsoft\AppV\ISV",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\AppV\ISV",
		"HKLM:\SOFTWARE\Microsoft\AppVISV",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\AppVISV",
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders",
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\SharedDLLs"
	)
	# Value data are paths into Office
	$DataKeys = @(
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"
	)
	foreach ($UserHive in $UserHives)
	{
		$NameKeys += @(
			"Registry::HKEY_USERS\$($UserHive.Key)\Software\Microsoft\AppV\ISV",
			"Registry::HKEY_USERS\$($UserHive.Key)\Software\Microsoft\AppVISV"
		)
		$DataKeys += "Registry::HKEY_USERS\$($UserHive.Key)\Software\Microsoft\Windows\CurrentVersion\Run"
	}

	foreach ($Key in (Get-Item -LiteralPath $NameKeys))
	{
		foreach ($Name in $Key.GetValueNames())
		{
			if (Test-OfficePath -Path $Name)
			{
				Remove-RegistryValue -Path $Key.PSPath -Name $Name
			}
		}
	}

	foreach ($Key in (Get-Item -LiteralPath $DataKeys))
	{
		foreach ($Name in $Key.GetValueNames())
		{
			if ((@("Lync15", "Lync16") -contains $Name) -or (Test-OfficePath -Path ([string]$Key.GetValue($Name))))
			{
				Remove-RegistryValue -Path $Key.PSPath -Name $Name
			}
		}
	}

	Write-Verbose -Message "Removing Add/Remove Programs entries, App Paths and Default Programs registration" -Verbose

	# Add/Remove Programs
	foreach ($UninstallKey in (Get-ChildItem -LiteralPath $UninstallPaths))
	{
		$Name = $UninstallKey.PSChildName
		$IsOffice = $Name.StartsWith("Office12.", [System.StringComparison]::OrdinalIgnoreCase) -or
			$Name.StartsWith("Office14.", [System.StringComparison]::OrdinalIgnoreCase) -or
			$Name.StartsWith("Office15.", [System.StringComparison]::OrdinalIgnoreCase) -or
			$Name.StartsWith("Office16.", [System.StringComparison]::OrdinalIgnoreCase) -or
			(Test-OfficePath -Path ([string]$UninstallKey.GetValue("UninstallString"))) -or
			(($Name.Length -ge 38) -and (Test-OfficeProductCode -ProductCode $Name.Substring(0, 38)))

		if ($IsOffice)
		{
			Remove-Item -LiteralPath $UninstallKey.PSPath -Recurse -Force
		}
	}

	# App Paths
	foreach ($Key in (Get-ChildItem -LiteralPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths"))
	{
		if ((Test-OfficePath -Path ([string]$Key.GetValue(""))) -or (Test-OfficePath -Path ([string]$Key.GetValue("Path"))))
		{
			Remove-Item -LiteralPath $Key.PSPath -Recurse -Force
		}
	}

	# Default Programs registration
	foreach ($Key in (Get-Item -LiteralPath "HKLM:\SOFTWARE\RegisteredApplications", "HKLM:\SOFTWARE\WOW6432Node\RegisteredApplications"))
	{
		foreach ($Name in $Key.GetValueNames())
		{
			$Data = [string]$Key.GetValue($Name)
			if ($Data.StartsWith("Software\Microsoft\Office", [System.StringComparison]::OrdinalIgnoreCase) -or ($Data.IndexOf("\Microsoft Outlook\", [System.StringComparison]::OrdinalIgnoreCase) -ge 0))
			{
				Remove-RegistryValue -Path $Key.PSPath -Name $Name
			}
		}
	}

	#region Windows Installer metadata
	Write-Verbose -Message "Cleaning Windows Installer metadata. It may take a few minutes" -Verbose

	$UserDataSids = @()
	$UserData = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData")
	if ($UserData)
	{
		$UserDataSids = $UserData.GetSubKeyNames()
		$UserData.Close()
	}

	# Subkey names are packed product codes
	$ProductRoots = @(
		@{Hive = [Microsoft.Win32.Registry]::LocalMachine; SubKey = "SOFTWARE\Classes\Installer\Products"},
		@{Hive = [Microsoft.Win32.Registry]::LocalMachine; SubKey = "SOFTWARE\Classes\Installer\Features"}
	)
	# Value names are packed product codes
	$ValueRoots = @(
		@{Hive = [Microsoft.Win32.Registry]::LocalMachine; SubKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes"},
		@{Hive = [Microsoft.Win32.Registry]::LocalMachine; SubKey = "SOFTWARE\Classes\Installer\UpgradeCodes"}
	)
	# REG_MULTI_SZ lists of Darwin descriptors
	$PublishedRoots = @(
		@{Hive = [Microsoft.Win32.Registry]::LocalMachine; SubKey = "SOFTWARE\Classes\Installer\Components"}
	)

	foreach ($Sid in $UserDataSids)
	{
		$ProductRoots += @{Hive = [Microsoft.Win32.Registry]::LocalMachine; SubKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$Sid\Products"}
		$ValueRoots += @{Hive = [Microsoft.Win32.Registry]::LocalMachine; SubKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$Sid\Components"}
	}

	foreach ($UserHive in $UserHives)
	{
		$ProductRoots += @(
			@{Hive = [Microsoft.Win32.Registry]::Users; SubKey = "$($UserHive.Key)\Software\Microsoft\Installer\Products"},
			@{Hive = [Microsoft.Win32.Registry]::Users; SubKey = "$($UserHive.Key)\Software\Microsoft\Installer\Features"}
		)
		$ValueRoots += @{Hive = [Microsoft.Win32.Registry]::Users; SubKey = "$($UserHive.Key)\Software\Microsoft\Installer\UpgradeCodes"}
		$PublishedRoots += @{Hive = [Microsoft.Win32.Registry]::Users; SubKey = "$($UserHive.Key)\Software\Microsoft\Installer\Components"}
	}

	# .NET is used for enumeration: Components may contain 100 000+ subkeys
	foreach ($Location in $ProductRoots)
	{
		$Root = $Location.Hive.OpenSubKey($Location.SubKey)
		if (-not $Root)
		{
			continue
		}

		foreach ($Name in $Root.GetSubKeyNames())
		{
			if (Test-OfficeProductCode -ProductCode $Name)
			{
				Remove-Item -LiteralPath "Registry::$($Root.Name)\$Name" -Recurse -Force
			}
		}
		$Root.Close()
	}

	Write-Verbose -Message "Cleaning upgrade codes and components" -Verbose

	# The original script required the UpgradeCode itself to match the product code pattern, so it practically never cleaned anything
	foreach ($Location in $ValueRoots)
	{
		$Root = $Location.Hive.OpenSubKey($Location.SubKey)
		if (-not $Root)
		{
			continue
		}

		foreach ($SubKeyName in $Root.GetSubKeyNames())
		{
			$Key = $Root.OpenSubKey($SubKeyName)
			if (-not $Key)
			{
				continue
			}

			$Removed = $false
			foreach ($Name in $Key.GetValueNames())
			{
				if (Test-OfficeProductCode -ProductCode $Name)
				{
					Remove-RegistryValue -Path "Registry::$($Key.Name)" -Name $Name
					$Removed = $true
				}
			}

			# Only keys emptied here: empty keys of other products are not ours to remove
			if ($Removed -and ($Key.ValueCount -eq 0) -and ($Key.SubKeyCount -eq 0))
			{
				Remove-Item -LiteralPath "Registry::$($Key.Name)" -Recurse -Force
			}
			$Key.Close()
		}
		$Root.Close()
	}

	Write-Verbose -Message "Cleaning published components" -Verbose

	foreach ($Location in $PublishedRoots)
	{
		$Root = $Location.Hive.OpenSubKey($Location.SubKey)
		if (-not $Root)
		{
			continue
		}

		foreach ($ComponentId in $Root.GetSubKeyNames())
		{
			$Key = $Root.OpenSubKey($ComponentId)
			if (-not $Key)
			{
				continue
			}

			$Removed = $false
			foreach ($Name in $Key.GetValueNames())
			{
				if ($Key.GetValueKind($Name) -ne [Microsoft.Win32.RegistryValueKind]::MultiString)
				{
					continue
				}

				$Entries = @($Key.GetValue($Name))
				$Keep = @(
					foreach ($Entry in $Entries)
					{
						# Entries that cannot be decoded are kept: the original script silently dropped them on rewrite
						$ProductCode = ConvertFrom-SquishedGuid -SquishedGuid $Entry
						if (-not ($ProductCode -and (Test-OfficeProductCode -ProductCode $ProductCode)))
						{
							$Entry
						}
					}
				)

				if ($Keep.Count -eq $Entries.Count)
				{
					continue
				}

				$Removed = $true
				if ($Keep.Count -eq 0)
				{
					Remove-RegistryValue -Path "Registry::$($Key.Name)" -Name $Name
				}
				elseif ($Name -eq "")
				{
					Set-ItemProperty -LiteralPath "Registry::$($Key.Name)" -Name "(default)" -Value ([string[]]$Keep) -Type MultiString
				}
				else
				{
					Set-ItemProperty -LiteralPath "Registry::$($Key.Name)" -Name $Name -Value ([string[]]$Keep) -Type MultiString
				}
			}

			if ($Removed -and ($Key.ValueCount -eq 0) -and ($Key.SubKeyCount -eq 0))
			{
				Remove-Item -LiteralPath "Registry::$($Key.Name)" -Recurse -Force
			}
			$Key.Close()
		}
		$Root.Close()
	}
	#endregion Windows Installer metadata

	#region COM and file type registrations
	$ClassesKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey("SOFTWARE\Classes")
	$DeletedClsids = New-Object -TypeName System.Collections.Generic.HashSet[string] -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)
	$DeletedTypeLibs = New-Object -TypeName System.Collections.Generic.HashSet[string] -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)
	$DeletedProgIds = New-Object -TypeName System.Collections.Generic.HashSet[string] -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)
	$AppIds = New-Object -TypeName System.Collections.Generic.HashSet[string] -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)

	Write-Verbose -Message "Removing Office COM classes. It may take a few minutes" -Verbose

	# CLSID: a server inside Office makes the class an Office class. CodeBase covers managed COM servers hosted by mscoree.dll
	foreach ($ClsidRoot in @("CLSID", "WOW6432Node\CLSID"))
	{
		$Root = $ClassesKey.OpenSubKey($ClsidRoot)
		if (-not $Root)
		{
			continue
		}

		foreach ($Clsid in $Root.GetSubKeyNames())
		{
			$Key = $Root.OpenSubKey($Clsid)
			if (-not $Key)
			{
				continue
			}

			$IsOffice = $false
			foreach ($ServerName in @("InprocServer32", "LocalServer32", "InprocHandler32", "LocalServer"))
			{
				$Server = $Key.OpenSubKey($ServerName)
				if (-not $Server)
				{
					continue
				}

				if ((Test-OfficePath -Path ([string]$Server.GetValue(""))) -or (Test-OfficePath -Path ([string]$Server.GetValue("CodeBase"))))
				{
					$IsOffice = $true
				}
				$Server.Close()
			}

			if ($IsOffice)
			{
				$AppId = [string]$Key.GetValue("AppID")
				if ($AppId)
				{
					[void]$AppIds.Add($AppId)
				}
				[void]$DeletedClsids.Add($Clsid)
			}
			$Key.Close()

			if ($IsOffice)
			{
				Remove-Item -LiteralPath "Registry::$($Root.Name)\$Clsid" -Recurse -Force
			}
		}
		$Root.Close()
	}

	Write-Verbose -Message "Removed COM classes: $($DeletedClsids.Count)" -Verbose
	Write-Verbose -Message "Removing Office type libraries and interfaces" -Verbose

	# Type libraries: a version goes when every registered file is inside Office
	# The original script also removed versions registered only under other LCIDs, because it started from "can delete"
	$Root = $ClassesKey.OpenSubKey("TypeLib")
	if ($Root)
	{
		foreach ($LibId in $Root.GetSubKeyNames())
		{
			$LibKey = $Root.OpenSubKey($LibId)
			if (-not $LibKey)
			{
				continue
			}

			$Removed = 0
			$Left = 0
			foreach ($Version in $LibKey.GetSubKeyNames())
			{
				$VersionKey = $LibKey.OpenSubKey($Version)
				if (-not $VersionKey)
				{
					continue
				}

				$Found = 0
				$Office = 0
				foreach ($Lcid in $VersionKey.GetSubKeyNames())
				{
					if (@("FLAGS", "HELPDIR") -contains $Lcid)
					{
						continue
					}

					$LcidKey = $VersionKey.OpenSubKey($Lcid)
					if (-not $LcidKey)
					{
						continue
					}

					foreach ($Platform in $LcidKey.GetSubKeyNames())
					{
						$PlatformKey = $LcidKey.OpenSubKey($Platform)
						if (-not $PlatformKey)
						{
							continue
						}

						$Found++
						if (Test-OfficePath -Path ([string]$PlatformKey.GetValue("")))
						{
							$Office++
						}
						$PlatformKey.Close()
					}
					$LcidKey.Close()
				}
				$VersionKey.Close()

				if (($Found -gt 0) -and ($Office -eq $Found))
				{
					Remove-Item -LiteralPath "Registry::$($LibKey.Name)\$Version" -Recurse -Force
					$Removed++
				}
				else
				{
					$Left++
				}
			}
			$LibKey.Close()

			if (($Removed -gt 0) -and ($Left -eq 0))
			{
				Remove-Item -LiteralPath "Registry::$($Root.Name)\$LibId" -Recurse -Force
				[void]$DeletedTypeLibs.Add($LibId)
			}
		}
		$Root.Close()
	}

	# Interfaces described by removed type libraries or marshaled by removed proxies
	foreach ($InterfaceRoot in @("Interface", "WOW6432Node\Interface"))
	{
		$Root = $ClassesKey.OpenSubKey($InterfaceRoot)
		if (-not $Root)
		{
			continue
		}

		foreach ($Iid in $Root.GetSubKeyNames())
		{
			$Key = $Root.OpenSubKey($Iid)
			if (-not $Key)
			{
				continue
			}

			$Delete = $false
			$TypeLibKey = $Key.OpenSubKey("TypeLib")
			if ($TypeLibKey)
			{
				$Delete = $DeletedTypeLibs.Contains([string]$TypeLibKey.GetValue(""))
				$TypeLibKey.Close()
			}

			$ProxyKey = $Key.OpenSubKey("ProxyStubClsid32")
			if ($ProxyKey)
			{
				$Delete = $Delete -or $DeletedClsids.Contains([string]$ProxyKey.GetValue(""))
				$ProxyKey.Close()
			}
			$Key.Close()

			if ($Delete)
			{
				Remove-Item -LiteralPath "Registry::$($Root.Name)\$Iid" -Recurse -Force
			}
		}
		$Root.Close()
	}

	Write-Verbose -Message "Removed type libraries: $($DeletedTypeLibs.Count)" -Verbose

	# AppID: Office executables and the AppIDs of removed classes
	foreach ($AppIdRoot in @("HKLM:\SOFTWARE\Classes\AppID", "HKLM:\SOFTWARE\Classes\WOW6432Node\AppID"))
	{
		foreach ($Executable in $OfficeExecutables)
		{
			$AppId = (Get-ItemProperty -LiteralPath "$AppIdRoot\$Executable" -Name AppID).AppID
			if ($AppId)
			{
				[void]$AppIds.Add($AppId)
			}
			Remove-Item -LiteralPath "$AppIdRoot\$Executable" -Recurse -Force
		}
	}
	foreach ($AppId in $AppIds)
	{
		Remove-Item -LiteralPath "HKLM:\SOFTWARE\Classes\AppID\$AppId", "HKLM:\SOFTWARE\Classes\WOW6432Node\AppID\$AppId" -Recurse -Force
	}

	Write-Verbose -Message "Removing Office file associations" -Verbose

	# ProgIDs and Applications\<exe>
	# The whole key goes only when its identity is Office (icon or class) and nothing non-Office is registered in it;
	# otherwise only the Office verbs and icon are removed, so htmlfile, mailto and the like survive
	$ContainerKeys = @("CLSID", "Interface", "TypeLib", "AppID", "Installer", "WOW6432Node", "Applications", "SystemFileAssociations", "MIME")
	$ProtectedKeys = @("*", "AllFilesystemObjects", "Directory", "Drive", "Folder", "Unknown", "DesktopBackground", "LibraryFolder")

	$Candidates = @(
		foreach ($Name in $ClassesKey.GetSubKeyNames())
		{
			if ((-not $Name.StartsWith(".")) -and ($ContainerKeys -notcontains $Name))
			{
				$Name
			}
		}
	)
	$ApplicationsKey = $ClassesKey.OpenSubKey("Applications")
	if ($ApplicationsKey)
	{
		foreach ($Name in $ApplicationsKey.GetSubKeyNames())
		{
			$Candidates += "Applications\$Name"
		}
		$ApplicationsKey.Close()
	}

	foreach ($Candidate in $Candidates)
	{
		$Key = $ClassesKey.OpenSubKey($Candidate)
		if (-not $Key)
		{
			continue
		}

		$OfficeVerbs = @()
		$OtherVerbs = 0
		$Shell = $Key.OpenSubKey("shell")
		if ($Shell)
		{
			foreach ($Verb in $Shell.GetSubKeyNames())
			{
				$Command = $Shell.OpenSubKey("$Verb\command")
				if ($Command -and (Test-OfficePath -Path ([string]$Command.GetValue(""))))
				{
					$OfficeVerbs += $Verb
				}
				else
				{
					$OtherVerbs++
				}

				if ($Command)
				{
					$Command.Close()
				}
			}
			$Shell.Close()
		}

		$IconPath = $null
		$Icon = $Key.OpenSubKey("DefaultIcon")
		if ($Icon)
		{
			$IconPath = [string]$Icon.GetValue("")
			$Icon.Close()
		}

		$Clsid = $null
		$ClsidKey = $Key.OpenSubKey("CLSID")
		if ($ClsidKey)
		{
			$Clsid = [string]$ClsidKey.GetValue("")
			$ClsidKey.Close()
		}

		$IsProtocol = $Key.GetValueNames() -contains "URL Protocol"
		$Key.Close()

		$OfficeIcon = Test-OfficePath -Path $IconPath
		$OfficeClsid = $Clsid -and $DeletedClsids.Contains($Clsid)

		if ($Candidate.StartsWith("Applications\", [System.StringComparison]::OrdinalIgnoreCase))
		{
			$DeleteAll = ($OfficeVerbs.Count -gt 0) -and ($OtherVerbs -eq 0)
		}
		else
		{
			$DeleteAll = ($OfficeIcon -or $OfficeClsid) -and ($OtherVerbs -eq 0) -and (-not $IsProtocol) -and ($ProtectedKeys -notcontains $Candidate)
		}

		if ($DeleteAll)
		{
			Remove-Item -LiteralPath "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\$Candidate" -Recurse -Force
			[void]$DeletedProgIds.Add($Candidate)
			continue
		}

		foreach ($Verb in $OfficeVerbs)
		{
			Remove-Item -LiteralPath "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\$Candidate\shell\$Verb" -Recurse -Force
		}
		if ($OfficeIcon)
		{
			Remove-Item -LiteralPath "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\$Candidate\DefaultIcon" -Recurse -Force
		}
	}

	Write-Verbose -Message "Removed ProgIDs: $($DeletedProgIds.Count)" -Verbose

	# File extensions pointing to removed ProgIDs and handlers
	foreach ($Extension in $ClassesKey.GetSubKeyNames())
	{
		if (-not $Extension.StartsWith("."))
		{
			continue
		}

		$Key = $ClassesKey.OpenSubKey($Extension)
		if (-not $Key)
		{
			continue
		}

		$ExtensionPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\$Extension"
		$DefaultWasOffice = $DeletedProgIds.Contains([string]$Key.GetValue(""))
		if ($DefaultWasOffice)
		{
			Remove-RegistryValue -Path $ExtensionPath -Name ""
		}

		foreach ($SubKeyName in $Key.GetSubKeyNames())
		{
			# .docx\Word.Document.12\ShellNew
			if ($DeletedProgIds.Contains($SubKeyName))
			{
				Remove-Item -LiteralPath "$ExtensionPath\$SubKeyName" -Recurse -Force
				continue
			}

			$SubKey = $Key.OpenSubKey($SubKeyName)
			if (-not $SubKey)
			{
				continue
			}

			switch ($SubKeyName)
			{
				"OpenWithProgids"
				{
					foreach ($Name in $SubKey.GetValueNames())
					{
						if ($DeletedProgIds.Contains($Name))
						{
							Remove-RegistryValue -Path "$ExtensionPath\OpenWithProgids" -Name $Name
						}
					}
				}
				"PersistentHandler"
				{
					if ($DeletedClsids.Contains([string]$SubKey.GetValue("")))
					{
						Remove-Item -LiteralPath "$ExtensionPath\PersistentHandler" -Recurse -Force
					}
				}
				"ShellEx"
				{
					foreach ($HandlerName in $SubKey.GetSubKeyNames())
					{
						$Handler = $SubKey.OpenSubKey($HandlerName)
						if ($Handler -and $DeletedClsids.Contains([string]$Handler.GetValue("")))
						{
							Remove-Item -LiteralPath "$ExtensionPath\ShellEx\$HandlerName" -Recurse -Force
						}

						if ($Handler)
						{
							$Handler.Close()
						}
					}
				}
			}
			$SubKey.Close()
		}

		# An extension created by Office: nothing but generic metadata left
		$Remaining = @($Key.GetValueNames() | Where-Object -FilterScript {@("", "Content Type", "PerceivedType") -notcontains $_})
		$OpenWith = $Key.OpenSubKey("OpenWithProgids")
		$OpenWithEmpty = (-not $OpenWith) -or ($OpenWith.ValueCount -eq 0)
		if ($OpenWith)
		{
			$OpenWith.Close()
		}
		$OnlyOpenWithLeft = ($Key.SubKeyCount -eq 0) -or (($Key.SubKeyCount -eq 1) -and ($Key.GetSubKeyNames() -contains "OpenWithProgids") -and $OpenWithEmpty)
		$Key.Close()

		if ($DefaultWasOffice -and ($Remaining.Count -eq 0) -and $OnlyOpenWithLeft)
		{
			Remove-Item -LiteralPath $ExtensionPath -Recurse -Force
		}
	}
	$ClassesKey.Close()

	# Per-user default app choices pointing to removed ProgIDs
	foreach ($UserHive in $UserHives)
	{
		$FileExts = [Microsoft.Win32.Registry]::Users.OpenSubKey("$($UserHive.Key)\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts")
		if (-not $FileExts)
		{
			continue
		}

		foreach ($Extension in $FileExts.GetSubKeyNames())
		{
			$UserChoice = $FileExts.OpenSubKey("$Extension\UserChoice")
			if (-not $UserChoice)
			{
				continue
			}

			$ProgId = [string]$UserChoice.GetValue("ProgId")
			$UserChoice.Close()
			if ($DeletedProgIds.Contains($ProgId))
			{
				Remove-Item -LiteralPath "Registry::$($FileExts.Name)\$Extension\UserChoice" -Recurse -Force
			}
		}
		$FileExts.Close()
	}

	# Approved shell extensions: known Office handlers and every removed class
	$ApprovedExtensions = @(
		"{B28AA736-876B-46DA-B3A8-84C5E30BA492}",
		"{8B02D659-EBBB-43D7-9BBA-52CF22C5B025}",
		"{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}",
		"{42042206-2D85-11D3-8CFF-005004838597}",
		"{993BE281-6695-4BA5-8A2A-7AACBFAAB69E}",
		"{C41662BB-1FA0-4CE0-8DC5-9B7F8279FF97}",
		"{506F4668-F13E-4AA1-BB04-B43203AB3CC0}",
		"{D66DC78C-4F61-447F-942B-3FB6980118CF}",
		"{46137B78-0EC3-426D-8B89-FF7C3A458B5E}",
		"{8BA85C75-763B-4103-94EB-9470F12FE0F7}",
		"{CD55129A-B1A1-438E-A425-CEBC7DC684EE}",
		"{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}",
		"{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}"
	)
	foreach ($Key in (Get-Item -LiteralPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved"))
	{
		foreach ($Name in $Key.GetValueNames())
		{
			if (($ApprovedExtensions -contains $Name) -or $DeletedClsids.Contains($Name))
			{
				Remove-RegistryValue -Path $Key.PSPath -Name $Name
			}
		}
	}
	#endregion COM and file type registrations
}
finally
{
	Write-Verbose -Message "Unloading registry hives" -Verbose

	# Provider objects keep hive handles open until they are collected
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()

	foreach ($UserHive in ($UserHives | Where-Object -FilterScript {$_.Unload}))
	{
		for ($Attempt = 1; $Attempt -le 5; $Attempt++)
		{
			& "$env:SystemRoot\System32\reg.exe" unload "HKU\$($UserHive.Key)" 2>$null | Out-Null
			if ($LASTEXITCODE -eq 0)
			{
				break
			}

			Start-Sleep -Seconds 1
			[System.GC]::Collect()
		}

		# A hive left loaded makes Windows log the user on with a temporary profile: a restart releases it
		if ($LASTEXITCODE -ne 0)
		{
			$RebootRequired = $true
		}
	}
}
#endregion Registry

#region Files
Write-Verbose -Message "Closing Explorer to release Office shell extensions" -Verbose

# Explorer holds Office shell extensions
Stop-OfficeProcess -ExtraName "explorer.exe", "msiexec.exe"

$Items = @(
	"$env:ProgramFiles\Microsoft Office",
	"$env:ProgramFiles\Microsoft Office 15",
	"$env:ProgramFiles\Microsoft Office 16",
	"$env:CommonProgramFiles\Microsoft Shared\ClickToRun",
	"$env:CommonProgramFiles\Microsoft Shared\OFFICE12",
	"$env:CommonProgramFiles\Microsoft Shared\OFFICE14",
	"$env:CommonProgramFiles\Microsoft Shared\OFFICE15",
	"$env:CommonProgramFiles\Microsoft Shared\OFFICE16",
	"$env:CommonProgramFiles\Microsoft Shared\OfficeSoftwareProtectionPlatform",
	"$env:CommonProgramFiles\Microsoft Shared\Source Engine",
	"$env:ProgramData\Microsoft\ClickToRun",
	"$env:ProgramData\Microsoft\Office",
	"$env:ProgramData\Microsoft\OfficeSoftwareProtectionPlatform",
	# Local installation source of MSI editions
	"$env:SystemDrive\MSOCache",
	"${env:ProgramFiles(x86)}\Microsoft Office",
	"${env:ProgramFiles(x86)}\Microsoft Office 15",
	"${env:ProgramFiles(x86)}\Microsoft Office 16",
	"${env:CommonProgramFiles(x86)}\Microsoft Shared\OFFICE12",
	"${env:CommonProgramFiles(x86)}\Microsoft Shared\OFFICE14",
	"${env:CommonProgramFiles(x86)}\Microsoft Shared\OFFICE15",
	"${env:CommonProgramFiles(x86)}\Microsoft Shared\OFFICE16",
	"${env:CommonProgramFiles(x86)}\Microsoft Shared\OfficeSoftwareProtectionPlatform",
	"${env:CommonProgramFiles(x86)}\Microsoft Shared\Source Engine"
)

# Every profile on the computer, not only the account the script runs under (SYSTEM under Intune or SCCM)
$OutlookFolders = @()
foreach ($UserProfile in $UserProfiles)
{
	$Items += @(
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Office",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Word",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Excel",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\PowerPoint",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Access",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Publisher",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\OneNote",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\MS Project",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Visio",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Templates",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Document Building Blocks",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\QuickStyles",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Bibliography",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Signatures",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Stationery",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\UProof",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Proof",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\AddIns",
		"$($UserProfile.LocalPath)\AppData\Local\Microsoft\Office",
		"$($UserProfile.LocalPath)\AppData\Local\Microsoft\OneNote",
		"$($UserProfile.LocalPath)\AppData\Local\Microsoft\FORMS",
		"$($UserProfile.LocalPath)\AppData\Local\Microsoft\Word",
		"$($UserProfile.LocalPath)\AppData\Local\Microsoft\Excel",
		"$($UserProfile.LocalPath)\AppData\Local\Microsoft\PowerPoint",
		"$($UserProfile.LocalPath)\AppData\Local\Microsoft\Windows\INetCache\Content.MSO",
		"$($UserProfile.LocalPath)\AppData\Local\Microsoft\Windows\INetCache\Content.Word",
		"$($UserProfile.LocalPath)\AppData\Local\Microsoft\Windows\INetCache\Content.Outlook",
		"$($UserProfile.LocalPath)\Microsoft Office",
		"$($UserProfile.LocalPath)\Microsoft Office 15",
		"$($UserProfile.LocalPath)\Microsoft Office 16"
	)

	# Outlook folders: everything except data files (*.pst)
	foreach ($OutlookFolder in @("$($UserProfile.LocalPath)\AppData\Local\Microsoft\Outlook", "$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Outlook"))
	{
		$Items += (Get-ChildItem -LiteralPath $OutlookFolder -Recurse -Force -File | Where-Object -FilterScript {$_.Extension -ne ".pst"}).FullName
		$OutlookFolders += (Get-ChildItem -LiteralPath $OutlookFolder -Recurse -Force -Directory).FullName
		$OutlookFolders += $OutlookFolder
	}
}

Write-Verbose -Message "Removing Office files and folders of the computer and $($UserProfiles.Count) user profile(s)" -Verbose

foreach ($Item in ($Items | Where-Object -FilterScript {$_}))
{
	Remove-Item -LiteralPath $Item -Recurse -Force

	# rd copes with some paths Windows PowerShell cannot remove, e.g. longer than MAX_PATH
	$Leftover = Get-Item -LiteralPath $Item -Force
	if (-not $Leftover)
	{
		continue
	}

	if ($Leftover.PSIsContainer)
	{
		& "$env:SystemRoot\System32\cmd.exe" /c rd /s /q "$Item" 2>$null
		$Leftover = Get-Item -LiteralPath $Item -Force
		if (-not $Leftover)
		{
			continue
		}
	}

	# Files in use: delete on reboot. Children first, because only empty folders can be deleted on reboot
	if ($Leftover.PSIsContainer)
	{
		$Leftovers = @(Get-ChildItem -LiteralPath $Item -Recurse -Force | Sort-Object -Property {$_.FullName.Length} -Descending) + $Leftover
	}
	else
	{
		$Leftovers = @($Leftover)
	}

	Write-Verbose -Message "In use, will be deleted on reboot: $Item" -Verbose

	foreach ($File in $Leftovers)
	{
		# MOVEFILE_DELAY_UNTIL_REBOOT. [NullString]::Value: a plain $null would be marshaled as an empty string and nothing would be scheduled
		if ([RemoveOffice.NativeMethods]::MoveFileEx($File.FullName, [NullString]::Value, 4))
		{
			$RebootRequired = $true
		}
	}
}

Write-Verbose -Message "Removing Office shortcuts" -Verbose

# Shortcuts: targets inside Office, including advertised Windows Installer shortcuts (%SystemRoot%\Installer\{...0000000FF1CE}\...)
# Only the folders shortcuts live in: the original script crawled the whole ProgramData and Users trees
$ShortcutFolders = @(
	"$env:ProgramData\Microsoft\Windows\Start Menu",
	"$env:PUBLIC\Desktop"
)
foreach ($UserProfile in $UserProfiles)
{
	$ShortcutFolders += @(
		"$($UserProfile.LocalPath)\Desktop",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Windows\Start Menu",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Windows\SendTo",
		"$($UserProfile.LocalPath)\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch"
	)
}

$StartMenuFolders = @()
$Shell = New-Object -ComObject WScript.Shell
foreach ($Shortcut in (Get-ChildItem -LiteralPath $ShortcutFolders -Filter "*.lnk" -Recurse -Force -File))
{
	if (Test-OfficePath -Path $Shell.CreateShortcut($Shortcut.FullName).TargetPath)
	{
		Write-Verbose -Message "Shortcut: $($Shortcut.FullName)" -Verbose
		Remove-Item -LiteralPath $Shortcut.FullName -Force

		# Folders emptied here are removed below only inside Start Menu\Programs
		if ($Shortcut.DirectoryName.IndexOf("\Start Menu\Programs\", [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
		{
			$StartMenuFolders += $Shortcut.DirectoryName
		}
	}
}
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Shell)

# Empty folders: exactly the folders collected above, deepest first, never walking up to the parents.
# The original script climbed up while a folder was empty and so could delete an emptied Desktop or Quick Launch folder
foreach ($Folder in (($StartMenuFolders + $OutlookFolders) | Where-Object -FilterScript {$_} | Sort-Object -Unique | Sort-Object -Property Length -Descending))
{
	if (-not (Get-ChildItem -LiteralPath $Folder -Force | Select-Object -First 1))
	{
		Remove-Item -LiteralPath $Folder -Force
	}
}

# Restart the shell as the logged-on users: not elevated and not as SYSTEM, which is what Start-Process from here would do
Start-Sleep -Seconds 3
if (-not (Get-Process -Name explorer))
{
	Write-Verbose -Message "Starting Explorer" -Verbose

	$Parameters = @{
		TaskName  = "RemoveOfficeExplorer"
		Action    = New-ScheduledTaskAction -Execute "$env:SystemRoot\explorer.exe"
		# BUILTIN\Users: runs in the session of every logged-on user
		Principal = New-ScheduledTaskPrincipal -GroupId "S-1-5-32-545" -RunLevel Limited
		Force     = $true
	}
	Register-ScheduledTask @Parameters | Out-Null
	Start-ScheduledTask -TaskName RemoveOfficeExplorer
	Start-Sleep -Seconds 3
	Unregister-ScheduledTask -TaskName RemoveOfficeExplorer -Confirm:$false
}

Remove-Item -LiteralPath $WorkFolder -Recurse -Force
#endregion Files

if ($RebootRequired)
{
	Write-Warning -Message "Office removed. Restart the computer to complete removal"
	exit 3010
}

Write-Verbose -Message "Office removed" -Verbose
exit 0