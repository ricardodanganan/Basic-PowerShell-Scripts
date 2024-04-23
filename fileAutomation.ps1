# This script organizes files in a source directory based on their file type (extension) into specific folders in a target directory.
# Created by Ricardo Danganan

# This is the source directory where the files will be organized and the target directory where the files will be moved to.
# They are defined as variables named $sourceDirectory and $targetDirectory, respectively.
$sourceDirectory = "C:\Users\RicardoDanganan\Downloads"
$targetDirectory = "C:\Users\RicardoDanganan\Desktop\FileAutomationFolder"

# Check if the source directory exists and if it contains any files to organize.
# Test-path is used to check if the directory exists
if (-Not (Test-Path -Path $targetDirectory)) {
    # If the target directory does not exist, using new-item to create a new directory with the specified path to the variable $targetDirectory
    New-Item -ItemType Directory -Path $targetDirectory
}

# Create an array of file types and corresponding folder names where the files will be moved to based on their extension.
# I named the array $fileTypeFolders 
$fileTypeFolders = @{
# Appropriate files are moved to the corresponding folders based on their extension.
# Example: (PDF files will be moved to the 'PDF-Files' folder)
    'pdf'  = 'PDF-Files' 
    'docx' = 'WordDocuments-Files' 
    'doc'  = 'WordDocuments-Files' 
    'pptx' = 'PowerpointPresentations-Files' 
    'ppt'  = 'PowerpointPresentations-Files' 
    'txt'  = 'Text-Files' 
    'xlsx' = 'Excel-Files' 
    'xls'  = 'Excel-Files'
}

# Get all the files in the source directory using Get-ChildItem and pipe the output to ForEach-Object to process each file.
Get-ChildItem -Path $sourceDirectory -File | ForEach-Object {
    # This variable is a placeholder variable that represents each file in the source directory during each iteration of the loop.
    # Assign the current file to a variable named $file.
    $file = $_
    # Get the extension of the file and remove the leading dot. 
    # Assign the extension to a variable named $extension.
    $extension = $file.Extension.TrimStart('.')
    # Check if the file extension exists in the $fileTypeFolders array.
    # ContainsKey is used to check if the key exists in the hashtable.
    if ($fileTypeFolders.ContainsKey($extension)) {
        # If the file extension exists in the array, create the target folder path using Join-Path and assign it to a variable named $targetFolder.
        # Join-Path is used to combine the $targetDirectory and the corresponding folder name based on the file extension.
        $targetFolder = Join-Path -Path $targetDirectory -ChildPath $fileTypeFolders[$extension]

        # Check if the target folder exists, if not, create the folder using New-Item.
        if (-Not (Test-Path -Path $targetFolder)) {
            New-Item -ItemType Directory -Path $targetFolder
        }

        # This variable named $targetFile is used to define the path where the file will be moved to.
        $targetFile = Join-Path -Path $targetFolder -ChildPath $file.Name
        Move-Item -Path $file.FullName -Destination $targetFile -Force
        # Output a message using Write-Host to indicate the successful move of the file.
        Write-Host "Moved '$($file.FullName)' to '$targetFile'"
    }
}

# Output a message using Write-Host to indicate the completion of the file organization process.
Write-Host "File organization complete."
