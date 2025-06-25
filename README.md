# Visual Pinball X Display Order Fix

powershell.exe -ExecutionPolicy Bypass -File .\VPinballXDisplayFix.ps1 -DryRun

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
5. Updates the same display-related keys in the `[Player]` section of `VPinballX.ini`.

---

## ⚙️ Parameters

### `-DryRun`
Simulates changes without applying them to the registry or INI file. Useful for testing.

### `-LogToFile`
Redirects all script output to a log file named `vpx_display_fix.log`, located in the same folder as `VPinballX.ini`. Each entry includes a timestamp.

---

## 📝 Notes

- Uses `System.Windows.Forms.Screen` to accurately detect the 0-based primary monitor index.
- Uses `Get-CimInstance` to retrieve the primary display’s refresh rate.
- Writes to the `HKCU` registry hive and `%APPDATA%\VPinballX\VPinballX.ini`.
- Does **not** require Administrator privileges under typical usage.
- Ensure Visual Pinball X is **closed** before running this script to prevent it from overwriting changes.
- Back up your `VPinballX.ini` if you have custom settings.

---

## 💻 Examples

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
