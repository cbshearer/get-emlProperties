# get-emlProperties
- Pull email to/from/cc/bcc/replyto addresses from a directory of eml files.
- A lot of the work was done by the excellent [script](https://github.com/PsCustomObject/PowerShell-Functions/blob/master/Convert-EmlFile.ps1) by [PsCustomObject](https://github.com/PsCustomObject)

## Operations
- Loop through a directory of email files
- Pull the properties I am interested in - email addresses
- Use regex to pull the addresses 
- Export the results into a CSV file

## Variables to Modify
- Line 50: $emlFiles - the directory containing your .eml files
- Line 51: $extractedAddresses = the csv file to put the email addresses into

## Output
- File: The name of the file bring processed
- Subject: The subject of the email being processed
- Addresses: The email addresses extracted and put into the CSV file.