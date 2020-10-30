# get-emlProperties
Pull email to/from/cc/bcc/replyto addresses from a directory of eml files.
A lot of the work was done by the excellent script: https://github.com/PsCustomObject/PowerShell-Functions/blob/master/Convert-EmlFile.ps1

I merely expanded upon it to:
loop through a directory of email files
pull the properties I am interested in - email addresses
use regex to pull the addresses 
export the results into a CSV file
