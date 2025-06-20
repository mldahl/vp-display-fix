<#
.SYNOPSIS
    Automatically identifies the primary display and updates Visual Pinball X
    display settings in both the Windows Registry and the VPinballX.ini file.

.DESCRIPTION
    This script performs the following actions:
    1.  Detects the primary monitor connected to your system, including its
        0-based index, width, height, color depth, and refresh rate.
        This detection method should align with how applications like
        PuPWinIDFix would identify displays.
    2.  Verifies the existence of the Visual Pinball X registry path and
        specific display-related keys. If any are missing, the script
        will print an error and exit.
    3.  Updates the 'Display', 'Height', 'Width', 'RefreshRate',
        and 'ColorDepth' registry keys located under
        'HKCU:\Software\Visual Pinball\VP10\Player'. These keys are relevant
        for older VPX versions or for the initial migration of settings to
        the INI file in VPX 10.8+.
    4.  Verifies the existence of the 'VPinballX.ini' file, the '[Player]'
        section within it, and specific display-related keys in that section.
        If any are missing, the script will print an error and exit.
    5.  Updates the 'Display', 'Height', 'Width',
        'RefreshRate', and 'ColorDepth' settings within the
        '[Player]' section of the 'VPinballX.ini' file. This INI file is
        the primary configuration source for Visual Pinball X 10.8 and
        later versions.

.PARAMETER DryRun
    If specified, the script will perform all detection and reporting,
    but will not make any actual changes to the registry or INI file.

.PARAMETER LogToFile
    If specified, all script output (which would normally go to the console)
    will be redirected and appended to a log file instead. The log file
    is named 'vpx_display_fix.log' and is located next to VPinballX.ini.
    Each log entry will be timestamped.

.NOTES
    -   This script uses the System.Windows.Forms.Screen class for display
        detection, which provides accurate 0-based indexing for monitors,
        and its BitsPerPixel for color depth.
    -   It uses Get-CimInstance (WMI) to attempt to retrieve the primary
        display's refresh rate from the video controller. This might not be
        perfectly accurate in all complex multi-GPU setups but is generally
        reliable for common pinball cabinet configurations.
    -   It writes to the HKEY_CURRENT_USER (HKCU) hive of the registry and
        the %APPDATA% directory, which typically do not require Administrator
        privileges. If you encounter permission issues, try running PowerShell
        as Administrator.
    -   Ensure Visual Pinball X is closed before running this script to
        prevent it from overwriting changes when it exits.
    -   Back up your VPinballX.ini file if you have custom settings that
        you are worried about overwriting.

.EXAMPLE
    To run the script normally (update settings and output to console):
    .\Update-VPXDisplaySettings.ps1

.EXAMPLE
    To perform a dry run (simulate changes without applying them):
    .\Update-VPXDisplaySettings.ps1 -DryRun

.EXAMPLE
    To run the script and log all output to a file:
    .\Update-VPXDisplaySettings.ps1 -LogToFile

.EXAMPLE
    To perform a dry run and log output to a file:
    .\Update-VPXDisplaySettings.ps1 -DryRun -LogToFile
#>

param (
    [switch]$DryRun,
    [switch]$LogToFile
)

# --- Configuration ---
$vpxRegPath = "HKCU:\Software\Visual Pinball\VP10\Player"
$vpxIniFileName = "VPinballX.ini"
$vpxIniSection = "Player"
$logFileName = "vpx_display_fix.log"
$logFilePath = Join-Path $env:APPDATA "VPinballX\$logFileName"

# --- Custom Logging Function ---
function Write-LogOutput {
    param (
        [string]$Message,
        [string]$ForegroundColor = "White"
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "[$timestamp] $Message"

    if ($LogToFile) {
        # Ensure log directory exists
        $logDir = Split-Path $logFilePath
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        Add-Content -Path $logFilePath -Value $logEntry
    }
    else {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
}

# --- Helper Functions for INI file manipulation ---
function Read-IniFileContent {
    param (
        [string]$FilePath
    )
    $iniData = [ordered]@{} # Use ordered hashtable to preserve section order
    $currentSection = ""

    if (Test-Path $FilePath) {
        Write-LogOutput "  Reading existing INI file: '$FilePath'" -ForegroundColor DarkCyan
        # Get-Content -Raw ensures entire file is read as a single string, then split by common line endings
        (Get-Content -Path $FilePath -Raw -Encoding UTF8) -split "`r`n|\n" | ForEach-Object {
            $line = $_.Trim()
            if ($line -match "^\[(.*)\]$") { # Section header like [Player]
                $currentSection = $matches[1]
                $iniData[$currentSection] = [ordered]@{} # New ordered hashtable for keys in this section
                if ($currentSection -eq $vpxIniSection) {
                    Write-LogOutput "    Found target section: '[$currentSection]'" -ForegroundColor DarkCyan
                }
            } elseif ($line -match "^([^=;\[\]]+?)\s*=(.*)$") { # Key-value pair like Sound3D = 4
                if ($currentSection) { # Only process keys if we are within a known section
                    $key = $matches[1].Trim()
                    $value = $matches[2].Trim()
                    # Ensure it's not a comment or malformed section-like line that got through
                    if (-not $key.StartsWith(";") -and -not $key.StartsWith("[")) {
                        $iniData[$currentSection][$key] = $value
                        # Write-LogOutput "      Found key '$key'='$value' in section '[$currentSection]'" -ForegroundColor DarkCyan
                    }
                }
            }
            # Lines that are comments, blank, or don't match key=value are skipped in this parsing phase
            # Their preservation is handled by Write-IniFileContent
        }
    } else {
        # This function assumes Test-Path was already done, if not, it will return empty $iniData
        Write-LogOutput "  INI file not found during Read-IniFileContent: '$FilePath'." -ForegroundColor Yellow
    }
    return $iniData
}

function Write-IniFileContent {
    param (
        [string]$FilePath,
        [hashtable]$IniDataToWrite, # Parameter type changed from [ordered] to [hashtable]
        [array]$OriginalLines,      # The raw original lines (to preserve comments/layout)
        [string]$TargetSection,     # The specific section to update (e.g., "Player")
        [switch]$DryRunSwitch       # Dry run flag for this function
    )

    $outputLines = New-Object System.Collections.ArrayList
    $sectionFoundAndUpdated = $false # Tracks if the target section was found and its keys handled
    $inCurrentSectionPass = $false # Flag for iterating through original lines

    foreach ($originalLine in $OriginalLines) {
        $trimmedOriginalLine = $originalLine.Trim()

        # Handle section headers
        if ($trimmedOriginalLine -match "^\[(.*)\]$") {
            # If we were previously in the target section, ensure it's marked as processed
            if ($inCurrentSectionPass) {
                $sectionFoundAndUpdated = $true
            }
            $inCurrentSectionPass = ($matches[1] -eq $TargetSection) # Check if this is the target section
            [void]$outputLines.Add($originalLine) # Always add the original section header
        }
        # Handle key-value pairs within the current section we're processing
        elseif ($inCurrentSectionPass -and $trimmedOriginalLine -match "^([^=;\[\]]+?)\s*=(.*)$") {
            $key = $matches[1].Trim()
            # If this key is one we want to update (it MUST exist in $IniDataToWrite for the target section)
            if ($IniDataToWrite[$TargetSection].Contains($key)) { # FIX: Changed to .Contains()
                # Add the updated value for this key
                [void]$outputLines.Add("$key=$($IniDataToWrite[$TargetSection][$key])")
                Write-LogOutput "    Updated existing key '$key' in section '[$TargetSection]'." -ForegroundColor Green
                # Remove it from IniDataToWrite so it's not checked again (as already written)
                # It's important to remove it from the data passed in ($IniDataToWrite) so it doesn't get added
                # as a "new" key later if we are writing out missing keys.
                $IniDataToWrite[$TargetSection].Remove($key)
            } else {
                [void]$outputLines.Add($originalLine) # Add original line if not one we're updating
            }
        }
        # Add other lines (comments, blank lines) as they are, if not a key-value we're explicitly handling
        else {
            [void]$outputLines.Add($originalLine)
        }
    }

    # After iterating through all original lines, if the target section was found, mark it as handled
    if ($inCurrentSectionPass -and -not $sectionFoundAndUpdated) { # This means the target section was the very last one
        $sectionFoundAndUpdated = $true
    }

    # Write the content
    if ($DryRunSwitch) {
        Write-LogOutput "  Dry run: Would write the following content to '$FilePath':`n$($outputLines -join "`n")" -ForegroundColor DarkYellow
    } else {
        Set-Content -Path $FilePath -Value ($outputLines -join "`n") -Encoding UTF8 -Force -ErrorAction Stop
        Write-LogOutput "'$FilePath' update complete." -ForegroundColor Green
    }
}

# --- Script Start ---
Write-LogOutput "Starting Visual Pinball X display configuration update..." -ForegroundColor Green
if ($DryRun) {
    Write-LogOutput "DRY RUN MODE ENABLED: No changes will be applied to the registry or INI file." -ForegroundColor DarkYellow
}
if ($LogToFile) {
    Write-LogOutput "Logging all output to file: '$logFilePath'" -ForegroundColor Cyan
}

# 1. Get Primary Display Information
# Load the System.Windows.Forms assembly to access screen information
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
}
catch {
    Write-Error "Failed to load System.Windows.Forms assembly. This is required for display detection. Error: $($_.Exception.Message)"
    exit 1
}

$screens = [System.Windows.Forms.Screen]::AllScreens
$primaryScreen = $null
$primaryDisplayIndex = -1

# Iterate through all detected screens to find the primary one
for ($i = 0; $i -lt $screens.Length; $i++) {
    if ($screens[$i].Primary) {
        $primaryScreen = $screens[$i]
        $primaryDisplayIndex = $i
        break
    }
}

# Check if a primary screen was found
if (-not $primaryScreen) {
    Write-Error "Could not find the primary display. Please ensure a primary display is configured in Windows. Exiting."
    exit 1
}

# Extract primary display properties
$displayWidth = $primaryScreen.Bounds.Width
$displayHeight = $primaryScreen.Bounds.Height
$colorDepth = $primaryScreen.BitsPerPixel # Color depth in bits per pixel

$refreshRate = $null
try {
    # Attempt to get refresh rate from Win32_VideoController
    # This assumes the primary video controller corresponds to the primary display.
    $videoController = Get-CimInstance -ClassName Win32_VideoController | Select-Object -First 1 CurrentRefreshRate
    if ($videoController) {
        $refreshRate = $videoController.CurrentRefreshRate
    }
}
catch {
    Write-LogOutput "Warning: Could not query Win32_VideoController for refresh rate. Error: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Ensure $refreshRate is not null for output/INI, provide a default or indicator if not found
if (-not $refreshRate -or $refreshRate -eq 0) { # 0 can sometimes indicate uninitialized or unsupported
    $refreshRate = "60" # Default to 60Hz if not found, as it's common
    Write-LogOutput "Warning: Primary display refresh rate could not be definitively determined. Defaulting to '$refreshRate Hz'." -ForegroundColor Yellow
}


Write-LogOutput "`nDetected Primary Display Details:" -ForegroundColor Cyan
Write-LogOutput "  Windows 0-based Index: $primaryDisplayIndex" -ForegroundColor Cyan
Write-LogOutput "  Width: $displayWidth px" -ForegroundColor Cyan
Write-LogOutput "  Height: $displayHeight px" -ForegroundColor Cyan
Write-LogOutput "  Color Depth (Bits Per Pixel): $colorDepth" -ForegroundColor Cyan
Write-LogOutput "  Refresh Rate: $refreshRate Hz" -ForegroundColor Cyan

# 2. Update Visual Pinball X Registry Keys
Write-LogOutput "`nAttempting to update Visual Pinball X Registry settings at '$vpxRegPath'..." -ForegroundColor Green

try {
    # --- Strict Check: Registry Path Existence ---
    if (-not (Test-Path $vpxRegPath)) {
        Write-Error "Registry path '$vpxRegPath' is missing. Exiting script."
        exit 1
    }

    # Helper function for updating registry with strict existence check
    function Update-VPXRegistryProperty {
        param (
            [string]$Path,
            [string]$Name,
            $Value,
            [switch]$DryRunSwitch
        )
        $fullRegKeyPath = "$Path\$Name"
        # Check if the specific property (key) exists within the path
        if (-not (Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue)) {
            Write-Error "Registry key '$fullRegKeyPath' is missing. Exiting script."
            exit 1
        }

        if ($DryRunSwitch) {
            Write-LogOutput "  Dry run: Would set '$Name' registry key to: '$Value'" -ForegroundColor DarkYellow
        } else {
            Set-ItemProperty -Path $Path -Name $Name -Value $Value -Force -ErrorAction Stop | Out-Null
            Write-LogOutput "  Set '$Name' registry key to: '$Value'" -ForegroundColor Green
        }
    }

    Update-VPXRegistryProperty -Path $vpxRegPath -Name "Display" -Value $primaryDisplayIndex -DryRunSwitch:$DryRun
    Update-VPXRegistryProperty -Path $vpxRegPath -Name "Height" -Value $displayHeight -DryRunSwitch:$DryRun
    Update-VPXRegistryProperty -Path $vpxRegPath -Name "Width" -Value $displayWidth -DryRunSwitch:$DryRun
    Update-VPXRegistryProperty -Path $vpxRegPath -Name "RefreshRate" -Value [int]$refreshRate -DryRunSwitch:$DryRun
    Update-VPXRegistryProperty -Path $vpxRegPath -Name "ColorDepth" -Value $colorDepth -DryRunSwitch:$DryRun

    Write-LogOutput "Registry update for Visual Pinball X complete." -ForegroundColor Green
}
catch {
    Write-Error "Failed to update registry settings for Visual Pinball X. Error: $($_.Exception.Message)"
    Write-LogOutput "This might be due to permissions or if VPX is running and locking the registry keys." -ForegroundColor Yellow
    exit 1 # Exit on registry error
}

# 3. Update VPinballX.ini File
$iniFilePath = Join-Path $env:APPDATA "VPinballX\$vpxIniFileName"

Write-LogOutput "`nAttempting to update VPinballX.ini file at '$iniFilePath'..." -ForegroundColor Green

try {
    # --- Strict Check: INI File Existence ---
    if (-not (Test-Path $iniFilePath)) {
        Write-Error "INI file '$iniFilePath' is missing. Exiting script."
        exit 1
    }

    # Use Get-Content to preserve all lines, including blank ones
    $originalIniLines = Get-Content -Path $iniFilePath -Encoding UTF8 -ErrorAction Stop

    # Read the current INI content into a structured format for checks
    $currentIniData = Read-IniFileContent -FilePath $iniFilePath

    # --- Strict Check: INI Section Existence ---
    if (-not $currentIniData.Contains($vpxIniSection)) { # FIX: Changed from ContainsKey to Contains for top-level ordered hashtable
        Write-Error "Section '[$vpxIniSection]' is missing in '$iniFilePath'. Exiting script."
        exit 1
    }

    # Define the keys to update and their new values (as strings for INI)
    $iniUpdates = [ordered]@{
        "Display"      = "$primaryDisplayIndex"
        "Height"       = "$displayHeight"
        "Width"        = "$displayWidth"
        "RefreshRate"  = "$refreshRate"
        "ColorDepth"   = "$colorDepth"
    }

    # --- Strict Check: INI Section Keys Existence ---
    foreach ($key in $iniUpdates.Keys) {
        if (-not $currentIniData[$vpxIniSection].Contains($key)) { # FIX: Changed from ContainsKey to Contains for section-level ordered hashtable
            Write-Error "Key '$key' is missing in section '[$vpxIniSection]' of '$iniFilePath'. Exiting script."
            exit 1
        }
    }

    # Pass the desired state (which now only contains existing keys thanks to strict checks)
    # and the original lines to the Write-IniFileContent function.
    # We create a new ordered hashtable for IniDataToWrite to ensure we only pass
    # the specific keys we intend to update, preventing unintended additions if this were to somehow change.
    $finalIniDataForWriting = [ordered]@{}
    $finalIniDataForWriting[$vpxIniSection] = [ordered]@{}
    foreach ($key in $iniUpdates.Keys) {
        $finalIniDataForWriting[$vpxIniSection].Add($key, $iniUpdates[$key])
    }

    Write-IniFileContent -FilePath $iniFilePath -IniDataToWrite $finalIniDataForWriting -OriginalLines $originalIniLines -TargetSection $vpxIniSection -DryRunSwitch:$DryRun
}
catch {
    Write-Error "Failed to update '$vpxIniFileName'. Error: $($_.Exception.Message)"
    Write-LogOutput "This might be due to permissions or if VPX is running and locking the file." -ForegroundColor Yellow
    exit 1 # Exit on INI error
}

Write-LogOutput "`nScript finished. Please restart Visual Pinball X to apply changes." -ForegroundColor Green
