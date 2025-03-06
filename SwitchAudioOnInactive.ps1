# PowerShell Audio Device Switcher with Inactivity Detection
# Switches to a specified audio device after a period of inactivity
# Usage: .\SwitchAudio.ps1 -TargetDeviceName "Audio Out Rear" -InactivitySeconds 15
param(
    [Parameter(Mandatory=$false)]
    [string]$TargetDeviceName = "Audio Out Rear (High Definition Audio Device)",
    
    [Parameter(Mandatory=$false)]
    [int]$InactivitySeconds = 60
)

# First, check if the module is installed
if (-not (Get-Module -ListAvailable -Name AudioDeviceCmdlets)) {
    try {
        Write-Host "Installing AudioDeviceCmdlets module..." -ForegroundColor Cyan
        Install-Module -Name AudioDeviceCmdlets -Force -Scope CurrentUser
        Write-Host "Module installed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to install the module: $_" -ForegroundColor Red
        Write-Host "Please run PowerShell as Administrator and try again, or install the module manually." -ForegroundColor Yellow
        Write-Host "You can install it manually by running: Install-Module -Name AudioDeviceCmdlets" -ForegroundColor Yellow
        exit 1
    }
}

# Import the module
Import-Module AudioDeviceCmdlets

# Add the required type for user activity tracking
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class UserInactivity {
    [DllImport("user32.dll")]
    public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

    public static uint GetIdleTime() {
        LASTINPUTINFO lastInput = new LASTINPUTINFO();
        lastInput.cbSize = (uint)Marshal.SizeOf(lastInput);
        
        if (GetLastInputInfo(ref lastInput)) {
            return ((uint)Environment.TickCount - lastInput.dwTime);
        } else {
            return 0;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct LASTINPUTINFO {
        public uint cbSize;
        public uint dwTime;
    }
}
"@

# Function to get all playback devices and find target by name
function Get-AudioDevices {
    try {
        $devices = Get-AudioDevice -List | Where-Object { $_.Type -eq "Playback" }
        return $devices
    }
    catch {
        Write-Host "Error retrieving audio devices: $_" -ForegroundColor Red
        exit 1
    }
}

# Function to set default audio device
function Set-AudioDeviceDefault {
    param (
        [Parameter(Mandatory=$true)]
        [int]$DeviceIndex
    )
    
    try {
        Set-AudioDevice -Index $DeviceIndex
        Write-Host "Successfully switched to device with index: $DeviceIndex" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error switching audio device: $_" -ForegroundColor Red
        return $false
    }
}

# Function to monitor user inactivity and switch device when threshold reached
function Start-InactivityMonitor {
    param (
        [Parameter(Mandatory=$true)]
        [int]$TargetDeviceIndex,
        [Parameter(Mandatory=$false)]
        [int]$InactivityThresholdSeconds = 15
    )
    
    $inactivityThresholdMs = $InactivityThresholdSeconds * 1000
    $devices = Get-AudioDevices
    $targetDevice = $devices | Where-Object { $_.Index -eq $TargetDeviceIndex }
    
    Write-Host "`nInactivity Monitor Started" -ForegroundColor Cyan
    Write-Host "Target Device: $($targetDevice.Name)" -ForegroundColor Cyan
    Write-Host "Inactivity Threshold: $InactivityThresholdSeconds seconds" -ForegroundColor Cyan
    Write-Host "Press Ctrl+C to stop monitoring" -ForegroundColor Yellow
    
    $switchedToTarget = $false
    
    try {
        while ($true) {
            $idleTime = [UserInactivity]::GetIdleTime()
            $currentDevice = Get-AudioDevice -Playback

            # Display current status with a progress bar
            Write-Progress -Activity "Monitoring Inactivity" -Status "Idle Time: $([math]::Round($idleTime/1000, 1)) seconds" -PercentComplete ([math]::Min(100, ($idleTime / $inactivityThresholdMs * 100)))
            
            if ($idleTime -ge $inactivityThresholdMs -and -not $switchedToTarget) {
                # Switch to target device after inactivity threshold
                Write-Host "`n[$(Get-Date -Format 'HH:mm:ss')] Inactivity threshold reached ($InactivityThresholdSeconds seconds)" -ForegroundColor Yellow
                
                if ($currentDevice.Index -ne $TargetDeviceIndex) {
                    $result = Set-AudioDeviceDefault -DeviceIndex $TargetDeviceIndex
                    if ($result) {
                        $switchedToTarget = $true
                    }
                } else {
                    Write-Host "Already using target device" -ForegroundColor Cyan
                    $switchedToTarget = $true
                }
            }
            
            if ($idleTime -lt 1000 -and $switchedToTarget) {
                # User is active again, reset the flag to allow future switching
                $switchedToTarget = $false
                Write-Host "`n[$(Get-Date -Format 'HH:mm:ss')] User activity detected - ready to monitor for next inactivity period" -ForegroundColor Cyan
            }
            
            Start-Sleep -Milliseconds 1000
        }
    }
    catch [System.Management.Automation.PipelineStoppedException] {
        # This exception is thrown when Ctrl+C is pressed
        Write-Host "`nMonitoring stopped by user" -ForegroundColor Yellow
    }
}

# Main script execution
Clear-Host
Write-Host "===== Audio Device Switcher with Inactivity Detection =====" -ForegroundColor Cyan

# Get all audio devices
$allDevices = Get-AudioDevices

# Find the target device by name
$targetDevice = $allDevices | Where-Object { $_.Name -like "*$TargetDeviceName*" }

if ($null -eq $targetDevice) {
    Write-Host "No device found matching name: '$TargetDeviceName'" -ForegroundColor Red
    Write-Host "Available devices:"
    $allDevices | ForEach-Object { Write-Host " - $($_.Name)" }
    exit 1
}

# If multiple matches, take the first one
if ($targetDevice -is [array]) {
    Write-Host "Multiple devices found matching '$TargetDeviceName'. Using the first match: $($targetDevice[0].Name)" -ForegroundColor Yellow
    $targetDevice = $targetDevice[0]
}

Write-Host "Found target device: $($targetDevice.Name) (Index: $($targetDevice.Index))" -ForegroundColor Green

# Start monitoring for inactivity
Start-InactivityMonitor -TargetDeviceIndex $targetDevice.Index -InactivityThresholdSeconds $InactivitySeconds
