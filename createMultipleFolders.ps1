# Description: This script creates multiple folders in the specified paths
# Author: Ricardo Danganan

# Define the paths where you want to create the new folders and assign it to an array named variable $folderPaths
$folderPaths = @("C:\Users\RicardoDanganan\Desktop\NewFolder1", "C:\Users\RicardoDanganan\Desktop\NewFolder2", "C:\Users\RicardoDanganan\Desktop\NewFolder3")

# Iterate over each folder path using foreach loop to create the folders and output the results using Write-Output or alias(echo) 
# $folderPath is a placeholder variable that represents each path in the $folderPaths array during each iteration of the loop 
foreach ($folderPath in $folderPaths) {
    # Check if the folder already exists using Test-Path then output the result using Write-Output or alias(echo) 
    if (!(Test-Path $folderPath)) {
        # If the folder doesn't exist, create it using New-Item or alias(md or mkdir) then output it using Write Output or alias(echo)
        mkdir -Path $folderPath
        # Output a confirmation message indicating the successful creation of the folder.
        Write-Output "Folder created at $folderPath"
    } else {
        # If the folder already exists, output a message indicating so.
        Write-Output "Folder already exists at $folderPath"
    }

    # Open the folder using Invoke-Item or alias (ii) 
    Invoke-Item $folderPath
}

