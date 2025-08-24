# USB Drive Enable/Disable Script

This repository provides a PowerShell script to **enable or disable USB drives** on Windows.  
It is mainly designed for external hard drives, allowing them to spin down when disabled in order to reduce heat generation and power consumption.

## Features
- Enable or disable USB storage devices by hardware ID (VID & PID).
- Primarily intended for external hard drives.
- Helps reduce unnecessary heat by spinning down drives when not in use.
- Uses [Sysinternals Handle](https://learn.microsoft.com/en-us/sysinternals/downloads/handle) to check if a drive is currently in use.

## Behavior
- If a drive is disabled while it is in use, it will remain **accessible until the PC is restarted**.
- After reboot, the device will be fully disabled.
- Always make sure the drive is not actively used before disabling it.

## Requirements
- Windows PowerShell (must be run with **Administrator privileges**).
- [Sysinternals Handle](https://learn.microsoft.com/en-us/sysinternals/downloads/handle) (`handle64.exe`) must be available in the script directory or in the system PATH.

## Usage
1. Edit the script to set your target USB device by **Vendor ID (VID)** and **Product ID (PID)**.  
   Example:
   ```powershell
   $targetVid = "067B"
   $targetPid = "2775"

2. Open PowerShell as Administrator and run the script:
.\toggle_usb_drive.ps1

3. Follow the on-screen prompts to enable or disable the USB drive.

## Disclaimer
Use this script at your own risk. Disabling a USB device while in use may cause data corruption.
