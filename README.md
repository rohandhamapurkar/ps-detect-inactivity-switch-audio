# PowerShell Audio Device Switcher

An automated PowerShell utility that switches to a specified audio output device after a period of user inactivity.

## Overview

This script monitors your computer for inactivity and automatically switches to a designated audio device when you step away from your computer. When user activity is detected again, the script prepares to monitor for the next inactivity period.

Perfect for users who want to:
- Automatically switch from headphones to speakers when they step away
- Enforce audio routing policies based on activity state
- Switch to different audio devices based on idle time thresholds

## Features

- Automatically detects and installs required dependencies
- Monitors system for user inactivity (keyboard/mouse)
- Switches to a specified audio output device after customizable inactivity period
- Visual progress indicator showing time until audio switch
- Detailed console logging of state changes
- Easily configurable via command-line parameters

## Requirements

- Windows operating system
- PowerShell 5.1 or higher
- Internet connection (for first-time module installation)
- Administrative privileges (may be required for module installation)

## Installation

1. Save the script as `SwitchAudio.ps1` to your preferred location
2. The script will automatically install the required AudioDeviceCmdlets module on first run

## Usage

### Basic Usage

Run the script with default parameters (switches to "Audio Out Rear" after 15 seconds of inactivity):

```powershell
.\SwitchAudio.ps1
```

### Custom Configuration

Specify a different target device and/or inactivity threshold:

```powershell
.\SwitchAudio.ps1 -TargetDeviceName "Speakers" -InactivitySeconds 30
```

### Parameters

- `-TargetDeviceName`: The name of the audio device to switch to (partial matches work)
- `-InactivitySeconds`: Number of seconds of inactivity before switching (default: 15)

## Troubleshooting

If you encounter module installation issues:

1. Run PowerShell as Administrator
2. Execute: `Install-Module -Name AudioDeviceCmdlets -Force`
3. Try running the script again

If your target device isn't being detected:

1. Run the script once to see the list of available devices
2. Use the exact name listed in the output with the `-TargetDeviceName` parameter

## How It Works

The script:
1. Checks for and installs the AudioDeviceCmdlets module if necessary
2. Imports native Windows functions to detect user input activity
3. Identifies available audio playback devices
4. Monitors for user inactivity using a progress bar display
5. Switches to the target device when the inactivity threshold is reached
6. Resets when user activity is detected

## License

This script is provided "as is" with no warranties. Use at your own risk.

## Acknowledgements

- Uses the [AudioDeviceCmdlets](https://www.powershellgallery.com/packages/AudioDeviceCmdlets/) PowerShell module
