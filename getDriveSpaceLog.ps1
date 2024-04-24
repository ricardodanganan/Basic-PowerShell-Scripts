# Description: This script will get the disk space of the drives and log it to a text file along with the current date and time.
# Author: Ricardo Danganan

# Get the current date and time using the Get-Date cmdlet and pipe it to the Out-File cmdlet to append it to the log file called "Disk-Space-log.txt".
# Out-file is used to write the output of a command to a file.
Get-Date | Out-File -FilePath C:\Users\RicardoDanganan\Desktop\Disk-Space-log.txt -Append
# Get the disk space of all drives using the Get-PSDrive cmdlet with the -PSProvider parameter set to FileSystem.
# Pipe the output to the Out-File cmdlet to append it to the log file.
Get-PSDrive -PSProvider FileSystem | Out-file -FilePath C:\Users\RicardoDanganan\Desktop\Disk-Space-log.txt -Append