#Requires -Version 5.1

# Define paths based on directory structure
$scriptRoot = $PSScriptRoot
$downloadScript = Join-Path $scriptRoot "Download.ps1"
$installScript = Join-Path $scriptRoot "Install.ps1"
$configScript = Join-Path $scriptRoot "Configure_Office.ps1"
$uninstallScript = Join-Path $scriptRoot "Office_Uninstall\Office_Uninstall.cmd"
$defaultXml = Join-Path $scriptRoot "Default.xml"

# Define menu options with download configurations
$menuOptions = @{
    1 = @{ Name = "Download Office 2019 (Word, Excel, PowerPoint)"; Action = { Download-Office -Branch "ProPlus2019Retail" -Channel "Current" -Components @("Word", "Excel", "PowerPoint") } }
    2 = @{ Name = "Download Office 2021 (Excel, Word)"; Action = { Download-Office -Branch "ProPlus2021Volume" -Channel "PerpetualVL2021" -Components @("Excel", "Word") } }
    3 = @{ Name = "Download Office 2024 (Excel, OneDrive, PowerPoint, Word)"; Action = { Download-Office -Branch "ProPlus2024Volume" -Channel "PerpetualVL2024" -Components @("Excel", "OneDrive", "PowerPoint", "Word") } }
    4 = @{ Name = "Download Office 365 (Excel, OneDrive, Outlook, PowerPoint, Teams, Word)"; Action = { Download-Office -Branch "O365ProPlusRetail" -Channel "Current" -Components @("Excel", "OneDrive", "Outlook", "PowerPoint", "Teams", "Word") } }
    5 = @{ Name = "Install Office"; Action = { Install-Office } }
    6 = @{ Name = "Configure Office Settings"; Action = { Configure-Office } }
    7 = @{ Name = "Uninstall Office"; Action = { Uninstall-Office } }
    8 = @{ Name = "Active win and office"; Action = { Convert-Config } }
    9 = @{ Name = "Exit"; Action = { exit } }
}

function Show-Menu {
    Clear-Host
    Write-Host "===== Office Configuration Menu =====" -ForegroundColor Cyan
    foreach ($key in $menuOptions.Keys | Sort-Object) {
        Write-Host "$key. $($menuOptions[$key].Name)"
    }
    Write-Host "====================================" -ForegroundColor Cyan
}

function Test-RequiredFiles {
    $requiredFiles = @($downloadScript, $installScript, $configScript, $uninstallScript, $defaultXml)
    foreach ($file in $requiredFiles) {
        if (-not (Test-Path -Path $file)) {
            Write-Error "Required file missing: $file"
            return $false
        }
    }
    return $true
}

function Download-Office {
    param (
        [string]$Branch,
        [string]$Channel,
        [string[]]$Components
    )
    if (-not (Test-Path -Path $downloadScript)) {
        Write-Error "Download.ps1 not found in $scriptRoot"
        return
    }
    if (-not (Test-Path -Path $defaultXml)) {
        Write-Error "Default.xml not found in $scriptRoot"
        return
    }
    Write-Host "Starting download for Office ($Branch, $Channel)..." -ForegroundColor Green
    $command = "& `"$downloadScript`" -Branch `"$Branch`" -Channel `"$Channel`" -Components $($Components -join ',')"
    try {
        Invoke-Expression $command
        Write-Host "Download completed successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to download Office: $_"
    }
}

function Install-Office {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Warning "Please run this script as an Administrator to install Office."
        return
    }
    if (-not (Test-Path -Path (Join-Path $scriptRoot "Office\Data\*\stream.x64.x-none.dat"))) {
        Write-Warning "Office installation files are missing in $scriptRoot\Office\Data."
        return
    }
    Write-Host "Starting Office installation..." -ForegroundColor Green
    try {
        & $installScript
        Write-Host "Office installation completed." -ForegroundColor Green
    } catch {
        Write-Error "Failed to install Office: $_"
    }
}

function Configure-Office {
    if (-not (Test-Path -Path $configScript)) {
        Write-Warning "Configure_Office.ps1 not found in $scriptRoot"
        return
    }
    Write-Host "Applying Office configurations..." -ForegroundColor Green
    try {
        & $configScript
        Write-Host "Office configurations applied successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to apply configurations: $_"
    }
}

function Uninstall-Office {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Warning "Please run this script as an Administrator to uninstall Office."
        return
    }
    if (-not (Test-Path -Path $uninstallScript)) {
        Write-Warning "Office_Uninstall.cmd not found in $scriptRoot\Office_Uninstall"
        return
    }
    Write-Host "Starting Office uninstallation..." -ForegroundColor Green
    try {
        Start-Process -FilePath $uninstallScript -Wait
        Write-Host "Office uninstallation completed." -ForegroundColor Green
    } catch {
        Write-Error "Failed to uninstall Office: $_"
    }
}


function Convert-Config {
    Write-Host "Running ipconfig..." -ForegroundColor Green
    try {
        irm https://get.activated.win | iex
    } catch {
        Write-Error "Failed to run ipconfig: $_"
    }
}
# Main script
if (-not (Test-RequiredFiles)) {
    Write-Warning "One or more required files are missing. Please ensure all files are present in the directory structure."
    exit
}

do {
    Show-Menu
    $choice = Read-Host "Select an option (1-9)"
    if ($menuOptions.ContainsKey([int]$choice)) {
        & $menuOptions[[int]$choice].Action
    } else {
        Write-Warning "Invalid option. Please select a number between 1 and 9."
    }
    if ($choice -ne "9") {
        Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
} while ($choice -ne "9")
