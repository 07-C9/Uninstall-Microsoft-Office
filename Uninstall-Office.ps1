#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Downloads and runs the Microsoft Support and Recovery Assistant (SaRA) to silently uninstall all versions of Microsoft Office.

.DESCRIPTION
    This script automates the process of completely removing Microsoft Office from a machine. It performs the following steps:
    1. Defines a temporary path for downloading and extracting files.
    2. Downloads the official SaRA command-line tool zip file from Microsoft.
    3. Extracts the contents of the zip file.
    4. Terminates any running Office application processes to ensure a clean uninstall.
    5. Runs the SaRA command-line tool with arguments to silently remove all detected Office versions.
    6. Cleans up by deleting the downloaded zip file and the extracted folder.
    7. A restart is recommended after the script completes to finalize the cleanup. [2]

.NOTES
    - This script must be run with Administrator privileges.
    - An active internet connection is required to download the tool.
    - The SaRA tool itself has an expiration date (typically 90 days from its creation) to ensure users have the latest version. [4] This script always downloads the latest version.
#>

# --- Configuration ---
# Direct download link for the SaRA Enterprise version
$downloadUrl = "https://aka.ms/SaRA_EnterpriseVersionFiles"

# Define a temporary working directory in the system's temp folder
$tempPath = Join-Path $env:TEMP "OfficeScrub"

# --- Script Execution ---

try {
    # Create the temporary directory if it doesn't exist
    if (-not (Test-Path -Path $tempPath)) {
        Write-Host "Creating temporary directory at $tempPath..."
        New-Item -Path $tempPath -ItemType Directory -Force | Out-Null
    }

    $zipFilePath = Join-Path $tempPath "SaRACmd.zip"

    # Download the SaRA tool
    Write-Host "Downloading Microsoft Support and Recovery Assistant..."
    Invoke-WebRequest -Uri $downloadUrl -OutFile $zipFilePath -UseBasicParsing

    # Extract the contents of the zip file
    Write-Host "Extracting files..."
    Expand-Archive -Path $zipFilePath -DestinationPath $tempPath -Force

    # Define the path to the SaRA executable
    $saraCmdPath = Join-Path $tempPath "SaRAcmd.exe"

    if (-not (Test-Path -Path $saraCmdPath)) {
        throw "SaRAcmd.exe not found after extraction. The zip file structure may have changed."
    }

    # Close any running Office applications to prevent uninstallation errors
    Write-Host "Closing any open Office applications..."
    $officeProcesses = @("winword", "excel", "powerpnt", "outlook", "onenote", "msaccess", "mspub", "visio", "winproj", "lync", "teams")
    foreach ($process in $officeProcesses) {
        Get-Process $process -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }

    # Run the Office uninstaller silently for all versions
    Write-Host "Starting the silent uninstallation of all Office versions. This may take some time..."
    $arguments = "-S OfficeScrubScenario -AcceptEula -OfficeVersion All"
    
    $processInfo = Start-Process -FilePath $saraCmdPath -ArgumentList $arguments -Wait -PassThru -Verb RunAs
    
    # Check the exit code from SaRAcmd.exe
    if ($processInfo.ExitCode -eq 0) {
        Write-Host "Successfully completed the Office uninstall scenario." -ForegroundColor Green
    } else {
        Write-Warning "The uninstallation script finished with a non-zero exit code: $($processInfo.ExitCode). This may indicate an issue."
        Write-Warning "Refer to SaRA documentation for exit code meanings."
    }

}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
}
finally {
    # Clean up the created directory and downloaded file
    if (Test-Path -Path $tempPath) {
        Write-Host "Cleaning up temporary files..."
        Remove-Item -Path $tempPath -Recurse -Force
    }
    
    Write-Host "Script finished. A computer restart is recommended to complete all cleanup tasks."
}