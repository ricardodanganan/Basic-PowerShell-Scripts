# Description: This script opens Google Chrome and navigates to the Outlook web app.
# Author: Ricardo Danganan

# Set the path to the Google Chrome shortcut on the desktop named variable $chromePath
$chromePath = 'C:\Users\RicardoDanganan\Desktop\Google Chrome.lnk'

# I made a variable called $url and set it to the URL of the Outlook web app.
$url = "https://outlook.office.com/"

# Use the Start-Process cmdlet to open Google Chrome and navigate to the URL.
# -ArgumentList is used to pass the URL as an argument to Google Chrome.
Start-Process -FilePath $chromePath -ArgumentList $url