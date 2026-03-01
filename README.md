[Recording 2026-03-01 202517.webm](https://github.com/user-attachments/assets/a817fa3b-d41f-46af-9386-cae939628cdb)

## Download Microsoft 365 (Office 365) via ODT with PowerShell 

## How-to

* Download [latest](https://github.com/farag2/Install-Office/releases/latest) archive and expand it;
* Open PowerShell console as admin and change execution policy

  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy Bypass -Force
  ```

* Open expanded `Office` folder in PowerShell console and choose which Office to download

* Configure which components to download
  * Branch:
    * ProPlus2019Retail
    * ProPlus2021Volume
    * ProPlus2024Volume
    * O365ProPlusRetail
  * Channel
    * Current
    * PerpetualVL2021
    * PerpetualVL2024
    * SemiAnnual
  * Components
    * Access
    * OneDrive
    * Outlook
    * Word
    * Excel
    * OneNote
    * Publisher
    * PowerPoint
    * Teams
    * ProjectPro2019Volume
    * ProjectPro2021Volume
    * ProjectPro2024Volume

  ```powershell
  .\Download.ps1 -Branch ProPlus2019Retail -Channel Current -Components Word, Excel, PowerPoint
  .\Download.ps1 -Branch ProPlus2021Volume -Channel PerpetualVL2021 -Components Excel, Word
  .\Download.ps1 -Branch ProPlus2024Volume -Channel PerpetualVL2024 -Components Excel, OneDrive, PowerPoint, Word
  .\Download.ps1 -Branch O365ProPlusRetail -Channel Current -Components Excel, OneDrive, Outlook, PowerPoint, Teams, Word
  ```

  * There won't be any messages in console while Office is being downloaded;
  * When Office is downloaded you can find a new `Office` folder in script folder.
* Do not move downloaded `Office` folder to another location as `Install.ps1` script has a link to `setup.exe`.
* Run `.\Install.ps1` as admin to install Office you downloaded;

## Addendum

`Office 2019, 2021, 2024, & 365` support `Windows 10` & `Windows 11` only

## Features

<details>
  <summary>List</summary>

* General
* Remove diagnostics tracking scheduled tasks
* Do not send additional diagnostic and usage data to Microsoft
* Disable LinkedIn features in Office applications
* Turn off the cloud features
* Turn on Touch/Mouse Mode

* Word
  * Do not show the Start screen when application starts
  * Do not open e-mail attachments and other uneditable files in reading view
  * Disable Protected View for files originating from the Internet
  * Disable Protected View for files located in potentially unsafe locations
  * Disable Protected View for Outlook attachments
  * Show the ruler
  * Save AutoRecover information every 3 minutes
  * Enable the "Draw" tab
  * Enable the "Developer" tab
  * Remove Adobe Acrobat Pro DC COM Add-ins

* Excel
  * Do not show the Start screen when application starts
  * Disable Protected View for files originating from the Internet
  * Disable Protected View for files located in potentially unsafe locations
  * Disable Protected View for Outlook attachments
  * Save AutoRecover information every 3 minutes
  * Enable the "Draw" tab
  * Enable the "Developer" tab

</details>

## Links

* [Configure Office](https://config.office.com/deploymentsettings)
* Overview of update channels
  * <https://learn.microsoft.com/en-us/deployoffice/overview-update-channels>
  * <https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run>
* [Office Deployment Tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
* [Deploy Office](https://learn.microsoft.com/en-us/deployoffice/deployment-guide-microsoft-365-apps)
* [Uninstall Office (SaRA)](https://www.microsoft.com/en-us/download/details.aspx?id=103391)
* [OffScrubC2R.vbs 2.19](https://github.com/farag2/Office/tree/master/Office_Uninstall)
* [Office Tool Plus](https://github.com/YerongAI/Office-Tool)
