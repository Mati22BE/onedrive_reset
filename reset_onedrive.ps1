<#
.SYNOPSIS
This script resets Microsoft OneDrive.

.DESCRIPTION
This script stops the Microsoft OneDrive process, resets it, and then closes any Explorer windows and Microsoft Store windows that are open.

.PARAMETER Folders
An array of folder paths where OneDrive might be installed.

.EXAMPLE
.\Reset-OneDrive.ps1
Resets Microsoft OneDrive using default folder paths.

.EXAMPLE
.\Reset-OneDrive.ps1 -Folders "C:\Program Files\Microsoft OneDrive\"
Resets Microsoft OneDrive using the specified folder path.

.NOTES
Version: v2.0
File Name: Reset-OneDrive.ps1
Author: Matisse Vuylsteke
Date Created: 28/03/2024
#>

#region FUNCTIONS
# Check OneDrive locations
function CheckOneDriveLocations {
    param (
        [string[]]$Folders
    )
    $foundLocations = @()

    foreach ($folder in $folders) {
        $OneDriveExe = Join-Path -Path $folder -ChildPath "OneDrive.exe"

        if (Test-Path $OneDriveExe) {
            $foundLocations += $OneDriveExe
        }
    }

    return $foundLocations
}

# Function to reset OneDrive
function ResetOneDrive {
    param (
        [string[]]$OneDriveLocations
    )

    $resetSuccessful = $false  # Flag to track if any reset attempt was successful

    # Check if OneDrive is found in any location
    if (-not ($OneDriveLocations)) {
        Write-Host "OneDrive is not installed in any of the provided locations." -ForegroundColor Red
        return $false
    }

    foreach ($location in $OneDriveLocations) {
        try {
            $oneDriveExe = $location
            # Reset OneDrive
            Write-Host "Resetting OneDrive in $location..."
            Start-Process -FilePath $oneDriveExe -ArgumentList "/reset" -Wait -NoNewWindow
            Start-Sleep -Seconds 5

            # Check if OneDrive process is still running
            if (Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue) {
                Write-Host "OneDrive process is still running after reset." -ForegroundColor Yellow
            } else {
                Write-Host "OneDrive reset completed in $location"
                $resetSuccessful = $true  # Set flag to true if reset is successful
            }
        }
        catch {
            # Catch any errors that occur during the reset process
            Write-Error -Message ("Error resetting OneDrive: {0}" -f $_.Exception.Message)
        }
    }
    
    # Return flag indicating if any reset attempt was successful
    return $resetSuccessful  
}
#endregion

# Variables
$folders = @("$env:LOCALAPPDATA\Microsoft\OneDrive\", "C:\Program Files\Microsoft OneDrive\", "C:\Program Files (x86)\Microsoft OneDrive\")

# Stop OneDrive
Write-Host "Resetting OneDrive..."
wsreset.exe

Start-Sleep -Seconds 5

# Call the function to check OneDrive locations
$OneDriveLocations = CheckOneDriveLocations -Folders $Folders

# Call the function to reset OneDrive
$success = ResetOneDrive -OneDriveLocations $OneDriveLocations

# Check if OneDrive reset was successful
if ($success) {
    # Close Explorer windows
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    # Close Microsoft Store windows
    Get-Process | Where-Object {$_.MainWindowTitle -like "*Microsoft Store*"} | ForEach-Object { Stop-Process -Id $_.Id -Force }
    Start-Sleep -Seconds 2
    Write-Host "OneDrive reset completed successfully." -ForegroundColor Green
} else {
    # Display error message if OneDrive reset encountered errors in all locations or if OneDrive is not installed
    Write-Host "OneDrive reset encountered errors or is not installed in any of the provided locations." -ForegroundColor Red
}
