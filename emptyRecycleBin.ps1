# Description: This script demonstrates how to empty the Recycle Bin using PowerShell.
# The script utilizes the Shell COM object to interact with the Recycle Bin and empty its contents.
# Author: Ricardo Danganan

# Create a new Shell application object using the COM object, New-Object cmdlet.
# I named the variable $shell to represent the Shell application object.
# The Shell.Application COM object is used to interact with the Windows Shell, including the Recycle Bin.
# The Shell object provides access to various system folders and functionalities.
$shell = New-Object -ComObject Shell.Application

# Get the Recycle Bin folder using the NameSpace() method of the Shell object.
# The (0xa) parameter represents the Recycle Bin folder in the Shell namespace.
# The $recycleBin variable represents the Recycle Bin folder.
$recycleBin = $shell.NameSpace(0xa)

# Check if the Recycle Bin is not empty by counting the items in the folder.
# If the Recycle Bin is not empty, use the Items() method of the $recycleBin object to get all items in the folder.
if ($recycleBin.Items().Count -gt 0) {
    # Iterate through each item in the Recycle Bin and delete it using the InvokeVerb() method with the "delete" verb.
    # The "delete" verb is used to permanently delete an item from the Recycle Bin.
    # The $item variable represents each item in the Recycle Bin during the iteration.
    foreach ($item in $recycleBin.Items()) {
        $item.InvokeVerb("delete")
        # Output a message using Write-Host to indicate the deletion of the item.
        # The $item.Name property is used to display the name of the deleted item.
        Write-Host "Deleted item: $($item.Name)"
    }
    # Output a message using Write-Host to indicate the successful emptying of the Recycle Bin.
    Write-Host "Recycle Bin emptied successfully."
} else {
    # Output a message using Write-Host to indicate that the Recycle Bin is already empty.
    Write-Host "Recycle Bin is already empty."
}

# Release the Shell COM object to free up resources.
# The [System.Runtime.Interopservices.Marshal]::ReleaseComObject() method is used to release the COM object.
# Pipe the output to Out-Null to suppress any output.
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
Remove-Variable shell

# Output a message using Write-Host to indicate the completion of the Recycle Bin cleanup process.
Write-Host "Recycle Bin cleanup process completed."

