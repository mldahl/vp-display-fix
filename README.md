# Visual Pinball X Display Order Fix

Automatically identifies the primary display and updates Visual Pinball X display settings in both the Windows Registry and the `VPinballX.ini` file.

---

## 🧭 Overview

This script performs the following actions:

1. Detects the primary monitor connected to your system, including its 0-based index, width, height, color depth, and refresh rate.
2. Verifies the existence of the Visual Pinball X registry path and specific display-related keys.
3. Updates the following registry keys under `HKCU:\Software\Visual Pinball\VP10\Player`:
   - `Display`
   - `Height`
   - `Width`
   - `RefreshRate`
   - `ColorDepth`
4. Verifies the existence of the `VPinballX.ini` file and `[Player]` section.
5. Updates the same display related keys in the `[Player]` section of `VPinballX.ini`.

---

## ⚙️ Parameters

### `-DryRun`
Simulates changes without applying them to the registry or INI file. Useful for testing.

### `-LogToFile`
Redirects all script output to a log file named `VPinballXDisplayFix.log`, located in the same folder as `VPinballX.ini`. Each entry includes a timestamp.

---

## 💻 Powershell Examples

Run normally (apply settings and show output in console):

```powershell
.\VPinballXDisplayFix.ps1
```

To perform a dry run (simulate changes without applying them):
```powershell
.\VPinballXDisplayFix.ps1 -DryRun
```

To run the script and log all output to a file:
```powershell
.\VPinballXDisplayFix.ps1 -LogToFile
```

---

## 🕹️ Using with PinUP Popper

1) Open PinUP Popper Config (\PinUPSystem\PinUpMenuSetup.exe)
2) Go to the "Popper Setup" tab and Select "GlobalConfig"
3) In the GlobalSettings window, selected the "StartUP" tab
4) In the Menu StartUP Script box, add the path to a hidden PowerShell launcher, replacing the file path with the actual location of the script:

```bat
start "" powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\To\VPinballXDisplayFix.ps1" -LogToFile
```

---

## 📝 Notes

- This works well for my setup.  You may need to make minor changes for yours
- Run this script with `-DryRun` first to ensure it works as expected before allow it to make any changes
- Writes to the `HKCU` registry hive and `%APPDATA%\VPinballX\VPinballX.ini`.
  - Make sure to backup the `Visual Pinball` registry key any custom changes to `VPinballX.ini` before running
- Does **not** require Administrator privileges under typical usage
- Ensure Visual Pinball X is **closed** before running this script to prevent it from overwriting changes
