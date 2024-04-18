# This script opens multiple URLs in Google Chrome using PowerShell. 
# It demonstrates how to use a foreach loop to iterate over a list of URLs and open them in separate Chrome windows. 
# The script uses the Start-Process cmdlet to launch Chrome with the specified URL as an argument.
# Author: Ricardo Danganan

# Set the path to my Google Chrome shortcut on the desktop.
# in programming this is called setting a variable.
$chromePath = 'C:\Users\RicardoDanganan\Desktop\Google Chrome.lnk'

# List of my essential work URLs at Typetec in an array.
$urls = @("https://outlook.office.com/","https://eu.myconnectwise.net", "https://typetec.hostedrmm.com:8040/Login", "https://typetec.eu.itglue.com/")

# Go through each URL in my list.
# foreach is a keyword that loops through a collection of items.
foreach ($url in $urls) {
    # Open each URL in a new Google Chrome window.
    # start-process is a cmdlet that starts one or more processes on the local computer.
    Start-Process -FilePath $chromePath -ArgumentList $url
}