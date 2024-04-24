# Description: This script demonstrates how to create a new Excel file using PowerShell and save it to a specified directory with a unique filename.
# The script utilizes the Excel.Application COM object to interact with Microsoft Excel and create a new workbook with custom content.
# The script also includes logic to find the next available filename by checking existing files in the specified directory.
# Author: Ricardo Danganan

# Create a new Excel application object using the COM object, New-Object cmdlet.
# COM objects are used to interact with applications that support COM automation, such as Microsoft Excel.
# I named the variable $excel to represent the Excel application object.
$excel = New-Object -ComObject Excel.Application

# Make the Excel application visible (This is optional)
$excel.Visible = $true

# Create a new workbook using the Workbooks.Add() method of the Excel application object, and assign it to the $workbook variable.
# The $workbook variable represents the newly created workbook.
$workbook = $excel.Workbooks.Add()

# Define the base directory and filename for the Excel file, which will be used to save the workbook with a unique filename.
# The $baseDirectory variable represents the base directory where the workbook will be saved.
# The $baseFilename variable represents the base filename for the workbook.
$baseDirectory = "C:\Users\RicardoDanganan\Desktop"
$baseFilename = "ExcelWorkbookTest"

# Using a while loop to find the next available number for the filename by checking existing files in the specified directory.
# The $index variable is initialized to 1, and it will be incremented until a unique filename is found.
# The Test-Path cmdlet is used to check if a file with the specified filename exists in the directory.
$index = 1
while (Test-Path -Path "$baseDirectory\$baseFilename$index.xlsx") {
    $index++
}

# Define the full path for the workbook with the next available number in the filename.
# The $filePath variable represents the full path where the workbook will be saved with the unique filename.
$filePath = "$baseDirectory\$baseFilename$index.xlsx"

# Save the workbook to the specified directory using the SaveAs() method of the $workbook object.
$workbook.SaveAs($filePath)

# Close the workbook using the Close() method of the $workbook object.
$workbook.Close()

# Quit the Excel application using the Quit() method of the $excel object to close the Excel application.
$excel.Quit()