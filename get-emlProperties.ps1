## This function is from "https://github.com/PsCustomObject/PowerShell-Functions/blob/master/Convert-EmlFile.ps1"

function Convert-EmlFile
{
<#
    .SYNOPSIS
        Function will parse an eml files.

    .DESCRIPTION
        Function will parse eml file and return a normalized object that can be used to extract infromation from the encoded file.

    .PARAMETER EmlFileName
        A string representing the eml file to parse.

    .EXAMPLE
        PS C:\> Convert-EmlFile -EmlFileName 'C:\Test\test.eml'

    .OUTPUTS
        System.Object
#>

    [CmdletBinding()]
    [OutputType([object])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$EmlFileName
    )

    # Instantiate new ADODB Stream object
    $adoStream = New-Object -ComObject 'ADODB.Stream'

    # Open stream
    $adoStream.Open()

    # Load file
    $adoStream.LoadFromFile($EmlFileName)

    # Instantiate new CDO Message Object
    $cdoMessageObject = New-Object -ComObject 'CDO.Message'

    # Open object and pass stream
    $cdoMessageObject.DataSource.OpenObject($adoStream, '_Stream')

    return $cdoMessageObject
}

$emlFiles = get-childitem "C:\temp\cbshearer\*.eml"
$extractedAddresses = "c:\temp\cbshearer\emails.csv"

foreach ($file in $emlfiles)

    {
        ## Null out variables
            $addresses = $null
            $emailData = $null
            
        Write-Host "`n################"
        Write-Host "File     :" $file.Name

        $emailData = convert-EmlFile -emlfilename $file

            ## Put address fields into the $addresses variable
                $addresses += $emailData.To
                $addresses += $emailData.CC
                $addresses += $emailData.BCC
                $addresses += $emailData.From
                $addresses += $emailData.ReplyTo
                $addresses += $emailData.Sender

        Write-Host "Subject  :" $emailData.Subject

        ## Extract email addresses
            $pattern = "([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})"
            $results = ($addresses | Select-String $pattern -AllMatches).Matches  

        ## export each address to a line of a CSV file.   
        Write-Host "Addresses: " 
        foreach ($item in ($results)) 
            { 
                write-host "  -" $item.value 
                $item | select-object value | export-csv $extractedAddresses -NoTypeInformation -NoClobber -Append
            }
    }