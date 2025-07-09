# Automated Office Uninstaller

A PowerShell script to silently uninstall all versions of Microsoft Office using the official Microsoft Support and Recovery Assistant (SaRA).

## How It Works

1.  Downloads the latest SaRA command-line tool from Microsoft.
2.  Extracts the tool to a temporary directory.
3.  Terminates running Office processes.
4.  Executes the SaRA tool to silently remove all Office versions.
5.  Removes all temporary files and folders.

Requires administrator privileges to run. A system restart is recommended after completion.

## Core Command

This script executes the following command to perform the uninstallation:

```powershell
SaRAcmd.exe -S OfficeScrubScenario -AcceptEula -OfficeVersion All