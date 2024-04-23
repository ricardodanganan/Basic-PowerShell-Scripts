# Description: This script demonstrates how to create a new Word document using PowerShell and save it to a specified directory with a unique filename. 
# The script utilizes the Word.Application COM object to interact with Microsoft Word and create a new document with custom content. 
# The script also includes logic to find the next available filename by checking existing files in the specified directory.
# Author: Ricardo Danganan

# Create a new Word application object using the COM object, New-Object cmdlet.
# I named the variable $word to represent the Word application object.
$word = New-Object -ComObject Word.Application

# Make the Word application visible (This is optional)
$word.Visible = $true

# Create a new document using the Documents.Add() method of the Word application object, and assign it to the $document variable.
# The $document variable represents the newly created document.
$document = $word.Documents.Add()

# Add content to the document by creating a new paragraph using the Content.Paragraphs.Add() method and setting the text using the Range.Text property.
# The text "This is a test document created with PowerShell." is added to the document when the script is executed.
# The $paragraph variable represents the newly added paragraph.
$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "This is a test document created with PowerShell."

# Define the base directory and filename for the Word document, which will be used to save the document with a unique filename.
# The $baseDirectory variable represents the base directory where the document will be saved.
# The $baseFilename variable represents the base filename for the document.
$baseDirectory = "C:\Users\RicardoDanganan\Desktop"
$baseFilename = "WordDocumentTest"

# Using a while loop to find the next available number for the filename by checking existing files in the specified directory.
# The $index variable is initialized to 1, and it will be incremented until a unique filename is found.
# The Test-Path cmdlet is used to check if a file with the specified filename exists in the directory.
$index = 1
while (Test-Path -Path "$baseDirectory\$baseFilename$index.docx") {
    $index++
}

# Define the full path for the document with the next available number in the filename.
# The $filePath variable represents the full path where the document will be saved with the unique filename.
$filePath = "$baseDirectory\$baseFilename$index.docx"

# Save the document to the specified directory using the SaveAs() method of the $document object.
$document.SaveAs($filePath)

# Close the document using the Close() method of the $document object.
$document.Close()
