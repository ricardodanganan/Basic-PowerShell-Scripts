# Description: This script demonstrates how to create a new PowerPoint presentation using PowerShell and save it to a specified directory with a unique filename.
# The script utilizes the PowerPoint.Application COM object to interact with Microsoft PowerPoint and create a new presentation with custom content.
# The script also includes logic to find the next available filename by checking existing files in the specified directory.
# Author: Ricardo Danganan

# Create a new PowerPoint application object using the COM object, New-Object cmdlet.
# COM objects are used to interact with applications that support COM automation, such as Microsoft PowerPoint.
# I named the variable $powerpoint to represent the PowerPoint application object.
$powerpoint = New-Object -ComObject PowerPoint.Application

# Make the PowerPoint application visible (This is optional)
$powerpoint.Visible = $true

# Create a new presentation using the Presentations.Add() method of the PowerPoint application object, and assign it to the $presentation variable.
# The $presentation variable represents the newly created presentation.
$presentation = $powerpoint.Presentations.Add()

# Define the base directory and filename for the PowerPoint presentation, which will be used to save the presentation with a unique filename.
# The $baseDirectory variable represents the base directory where the presentation will be saved.
# The $baseFilename variable represents the base filename for the presentation.
$baseDirectory = "C:\Users\RicardoDanganan\Desktop"
$baseFilename = "PowerPointPresentationTest"

# Using a while loop to find the next available number for the filename by checking existing files in the specified directory.
# The $index variable is initialized to 1, and it will be incremented until a unique filename is found.
# The Test-Path cmdlet is used to check if a file with the specified filename exists in the directory.
$index = 1
while (Test-Path -Path "$baseDirectory\$baseFilename$index.pptx") {
    $index++
}

# Define the full path for the presentation with the next available number in the filename.
# The $filePath variable represents the full path where the presentation will be saved with the unique filename.
$filePath = "$baseDirectory\$baseFilename$index.pptx"

# Save the presentation to the specified directory using the SaveAs() method of the $presentation object.
$presentation.SaveAs($filePath)

# Close the presentation using the Close() method of the $presentation object.
$presentation.Close()

# Quit the PowerPoint application using the Quit() method of the $powerpoint object to close the PowerPoint application.
$powerpoint.Quit()