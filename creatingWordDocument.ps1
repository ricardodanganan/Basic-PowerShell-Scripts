# Description: This script demonstrates how to create a Word document using PowerShell.
# Author: Ricardo Danganan

# Create a new Word application object using COM (Component Object Model) named variable $word
# COM is a technology that allows inter-process communication and dynamic object creation in a large range of programming languages.
$word = New-Object -ComObject Word.Application

# Make the Word application visible to the user (optional)
$word.Visible = $true

# Create a new document in Word application object named variable $document
$document = $word.Documents.Add()

# Add content to the document (in this case, a simple paragraph) using the Range object within a Paragraph object within the Content object of the document object 
# I named the variable $paragraph to represent the paragraph object that will be added to the document
# The Range object represents a contiguous area in the document, and the Text property is used to set the text content of the range
$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "This is a test document created with PowerShell."

# Save the document to a specified path using the SaveAs method of the document object
# You can adjust the path to save the document to a different location
$document.SaveAs("C:\Users\RicardoDanganan\Desktop\Ticket-Tutorials\TestDocument")

# Close the document without saving changes using the Close method of the document object
$document.Close()

# Quit the Word application using the Quit method of the Word application object
$word.Quit()


