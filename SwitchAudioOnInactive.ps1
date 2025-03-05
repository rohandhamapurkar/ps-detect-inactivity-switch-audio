# PowerShell Audio Device Switcher with Inactivity Detection
# Switches to a specified audio device after a period of inactivity

# First, check if the module is installed
if (-not (Get-Module -ListAvailable -Name AudioDeviceCmdlets)) {
    Write-Host "The AudioDeviceCmdlets module is required but not installed." -ForegroundColor Red
    Write-Host "Would you like to install it now? (Y/N)" -ForegroundColor Yellow
    $response = Read-Host
    
    if ($response -eq 'Y' -or $response -eq 'y') {
        try {
            Write-Host "Installing AudioDeviceCmdlets module..." -ForegroundColor Cyan
            Install-Module -Name AudioDeviceCmdlets -Force -Scope CurrentUser
            Write-Host "Module installed successfully!" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install the module: $_" -ForegroundColor Red
            Write-Host "Please run PowerShell as Administrator and try again, or install the module manually." -ForegroundColor Yellow
            Write-Host "You can install it manually by running: Install-Module -Name AudioDeviceCmdlets" -ForegroundColor Yellow
            exit
        }
    }
    else {
        Write-Host "Module installation declined. Exiting script." -ForegroundColor Yellow
        exit
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

# Function to show all playback devices
function Show-AudioDevices {
    Write-Host "`nAvailable Audio Playback Devices:" -ForegroundColor Cyan
    Write-Host "--------------------------------" -ForegroundColor Cyan
    
    try {
        $index = 1
        $devices = Get-AudioDevice -List | Where-Object { $_.Type -eq "Playback" }
        
        foreach ($device in $devices) {
            if ($device.Default) {
                Write-Host "$index. $($device.Name) [CURRENT DEFAULT]" -ForegroundColor Green
            } else {
                Write-Host "$index. $($device.Name)" -ForegroundColor White
            }
            $index++
        }
        
        return $devices
    }
    catch {
        Write-Host "Error retrieving audio devices: $_" -ForegroundColor Red
        exit
    }
}

# Function to set default audio device
function Set-AudioDeviceDefault {
    param (
        [Parameter(Mandatory=$true)]
        [int]$Index,
        [Parameter(Mandatory=$true)]
        [array]$Devices
    )
    
    if ($Index -lt 1 -or $Index -gt $Devices.Count) {
        Write-Host "Invalid device index. Please select a number between 1 and $($Devices.Count)" -ForegroundColor Red
        return $false
    }
    
    $selectedDevice = $Devices[$Index - 1]
    
    Write-Host "Switching to: $($selectedDevice.Name)" -ForegroundColor Yellow
    
    try {
        Set-AudioDevice -Index $selectedDevice.Index
        Write-Host "Successfully switched to: $($selectedDevice.Name)" -ForegroundColor Green
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
        [Parameter(Mandatory=$true)]
        [array]$Devices,
        [Parameter(Mandatory=$false)]
        [int]$InactivityThresholdSeconds = 15
    )
    
    $inactivityThresholdMs = $InactivityThresholdSeconds * 1000
    $defaultDevice = ($Devices | Where-Object { $_.Default }).Index
    
    Write-Host "`nInactivity Monitor Started" -ForegroundColor Cyan
    Write-Host "Target Device: $($Devices[$TargetDeviceIndex - 1].Name)" -ForegroundColor Cyan
    Write-Host "Inactivity Threshold: $InactivityThresholdSeconds seconds" -ForegroundColor Cyan
    Write-Host "Press Ctrl+C to stop monitoring" -ForegroundColor Yellow
    
    $switchedToTarget = $false
    
    try {
        while ($true) {
            $idleTime = [UserInactivity]::GetIdleTime()
            $currentDefaultDeviceIndex = ($Devices | Where-Object { $_.Default }).Index

            # Display current status with a progress bar
            Write-Progress -Activity "Monitoring Inactivity" -Status "Idle Time: $([math]::Round($idleTime/1000, 1)) seconds" -PercentComplete ([math]::Min(100, ($idleTime / $inactivityThresholdMs * 100)))
            
            if ($idleTime -ge $inactivityThresholdMs -and -not $switchedToTarget) {
                # Switch to target device after inactivity threshold
                $currentDevice = Get-AudioDevice -Playback
                Write-Host "`n[$(Get-Date -Format 'HH:mm:ss')] Inactivity threshold reached ($InactivityThresholdSeconds seconds)" -ForegroundColor Yellow
                
                if ($currentDevice.Index -ne $Devices[$TargetDeviceIndex - 1].Index) {
                    $result = Set-AudioDeviceDefault -Index $TargetDeviceIndex -Devices $Devices
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

# Get and display available audio devices
$devices = Show-AudioDevices

# Let user select a target device (to switch to after inactivity)
$targetDeviceSelection = Read-Host "`nEnter the number of the device you want to switch to after inactivity"
$inactivityThreshold = Read-Host "Enter inactivity threshold in seconds (default is 15)"

# Use default if no value entered
if ([string]::IsNullOrWhiteSpace($inactivityThreshold)) {
    $inactivityThreshold = 15
}

# Start monitoring for inactivity
Start-InactivityMonitor -TargetDeviceIndex $targetDeviceSelection -Devices $devices -InactivityThresholdSeconds $inactivityThreshold
