# Description: A PowerShell script that creates a new folder at a specified path if it doesn't already exist.
# Author: Ricardo Danganan

# Set the target path for the new folder named variable $folderPath
$folderPath = "C:\Users\RicardoDanganan\Desktop\NewFolder"

# Use Test-Path to verify if a folder already exists at the target path ($folderPath).
if (!(Test-Path $folderPath)) {
    # If the folder doesn't exist (Test-Path returns False), create it using mkdir.
    mkdir -Path $folderPath
    # Output a confirmation message indicating the successful creation of the folder.
    Write-Output "Successfully created a new folder at $folderPath"
} else {
    # If the folder already exists (Test-Path returns True), output a message indicating so.
    Write-Output "The folder already exists at $folderPath"
}

# Use Invoke-Item to open the newly created or existing folder in the default file explorer.
Invoke-Item $folderPath