# Windows Installed Apps Scanner (PowerShell 7.6 + 5.1)

A PowerShell tool that scans your PC for installed software and shows the results in a dark-blue HTML report, with a CSV export. It works on **PowerShell 7.6** and on Windows PowerShell **5.1**.

## How to run

1. Copy `InstalledAppsScanner.ps1` and `LaunchInstalledAppsScan.ps1` to `C:\`.
2. Right click on the .ps1 file and un-check 'block file'
<img width="400" height="263" alt="image" src="https://github.com/user-attachments/assets/5d9290a8-114b-4a75-8471-a686b75cc4b2" /> .
2.1: Run powershell command:
```powershell
Unblock-File C:\LaunchInstalledAppsScan.ps1
Unblock-File C:\InstalledAppsScanner.ps1
```
3. Run `LaunchInstalledAppsScan.ps1` to start the scan.


`LaunchInstalledAppsScan.ps1` uses this command:

```powershell
powershell.exe -NoExit -ExecutionPolicy Bypass -File "C:\InstalledAppsScanner.ps1"
```

This starts the scan in Windows PowerShell 5.1. To run it with **PowerShell 7.6**, use `pwsh.exe` instead:

```powershell
pwsh.exe -NoExit -ExecutionPolicy Bypass -File "C:\InstalledAppsScanner.ps1"
```

For the most complete results, run it **as Administrator**. This adds other users' AppData folders and all-users Microsoft Store apps.

<img width="714" height="176" alt="Open with PowerShell" src="https://github.com/user-attachments/assets/0fcecbe3-d799-4773-ae8e-e059e80ff099" />

When the scan finishes, the HTML report opens automatically. Both files are saved to your Desktop:

- `Apps_Inventory_<date>.html`
- `Apps_Inventory_<date>.csv`

## What's new

- **A-Z buttons** to jump to apps by first letter (letters with no matches are greyed out)
- **"Only Non-Microsoft" button** to filter out Microsoft software
- **"Show All Entries" button** to view the full list with no deduplication (duplicate names are hidden by default)
- **Search box**, **source filters** (Installed / Portable / Store) and **click-to-sort** column headers
- **Progress bar** while scanning, showing the current folder and files checked
- **CSV export to Desktop** with every entry and a `Duplicate` column
- **Wider scan coverage**:
  - Registry (HKLM and HKCU)
  - Portable apps
  - AppData
  - Microsoft Store apps

<img width="1254" height="92" alt="Progress bar" src="https://github.com/user-attachments/assets/7b482909-6d02-42f6-a7a4-f9af0649dca3" />

## HTML export

<img width="1670" height="842" alt="HTML export" src="https://github.com/user-attachments/assets/d18cac64-a413-4622-b93f-9fea67fa5c90" />

## Fixes

- **PowerShell 7.6 support:** Microsoft Store apps are now detected on PowerShell 7 (it loads the Appx module through Windows PowerShell compatibility)
- **No more parse errors or closing windows:** the script is pure ASCII, so it no longer breaks on PowerShell 5.1 (the `â€”` errors)
- **Correct sizes on any language setting:** sizes no longer show as `1,5` instead of `1.5` on non-English systems
- **Desktop path:** reports are saved correctly even if your Desktop is redirected to OneDrive
- **"Only Non-Microsoft" button** works correctly
- **Real app names** are shown (the Python naming bug is gone)
- **Professional dark-blue design**
- **Clean CSV backup**
- **Window stays open** at the end and shows any error message

## Notes

- Scanning Program Files, LocalAppData and AppData for `.exe` files can take several minutes. A progress bar shows where it is.
- The Excel (`.xlsx`) export was removed. Use the CSV instead.
Virus total scanned! because made with AI..
https://www.virustotal.com/gui/file/13598e8023e9e66c2a24297607070d566621409612ce0d7b11a9cbf6eff63098
