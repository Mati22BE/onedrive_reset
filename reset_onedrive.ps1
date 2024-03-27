<#
.Synopsis
This script resets Microsoft OneDrive.

.DESCRIPTION
This script resets Microsoft OneDrive by stopping the OneDrive process, resetting it, and then closing any Explorer windows and Microsoft Store windows that are open.

.PARAMETER Folders
An array of folder paths where OneDrive might be installed.

.EXAMPLE
.\Reset-OneDrive.ps1
Resets Microsoft OneDrive using default folder paths.

.EXAMPLE
.\Reset-OneDrive.ps1 -Folders "C:\Program Files\Microsoft OneDrive\"
Resets Microsoft OneDrive using the specified folder path.

.NOTES
Created on:   27/03/2024
Created by:   Matisse Vuylsteke
Filename:     reset_onedrive.ps1
#>

# Variables
$folders = @("$env:LOCALAPPDATA\Microsoft\OneDrive\", "C:\Program Files\Microsoft OneDrive\", "C:\Program Files (x86)\Microsoft OneDrive\")

# Stop OneDrive
Write-Host "Resetting OneDrive..."
wsreset.exe

Start-Sleep -Seconds 5

# Reset OneDrive
function ResetOneDrive {
    param (
        [string[]]$Folders
    )

    foreach ($folder in $Folders){
        if (Test-Path $folder) {
            try {
                $oneDriveExe = Join-Path -Path $folder -ChildPath "OneDrive.exe"
                if (Test-Path $oneDriveExe) {
                    Start-Process -FilePath $oneDriveExe -ArgumentList "/reset" -Wait -NoNewWindow
                    Write-Host "OneDrive reset completed in $folder"
                    return $true  # Return true if no error occurred
                } else {
                    Write-Host "OneDrive executable not found in $folder" -ForegroundColor Red
                    Start-Sleep -Seconds 1
                    Write-Host "Checking next folder..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 5
                    Write-Host "Resetting..."
                }
            }
            catch {
                Write-Error "Error resetting OneDrive: " + $_.Exception.Message
            }
        } else {
            Write-Error "Folder $folder not found"
        }
    }
    return $false  # Return false if an error occurred in all folders
}

# Call the function
$success = ResetOneDrive -Folders $folders

if ($success) {
    # Close Explorer windows
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    # Close Microsoft Store windows
    Get-Process | Where-Object {$_.MainWindowTitle -like "*Microsoft Store*"} | ForEach-Object { Stop-Process -Id $_.Id -Force }
    Start-Sleep -Seconds 2
    Write-Host "OneDrive reset completed successfully." -ForegroundColor Green
} else {
    Write-Host "OneDrive reset encountered errors in all folders." -ForegroundColor Red
}
