<#
.SYNOPSIS
    Uses Outlook Com.Object to directly import contacts with birthday. Outlook must be installed.
.DESCRIPTION
    Uses Outlook Com.Object to directly import contacts with birthday. Outlook must be installed.
.EXAMPLE
    PS C:\> .\Import-ContactsToOutlook.ps1 -SourceCsv "C:\Example.csv" -Encoding UTF8
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>

[CmdletBinding()]
param (
    # Specifies a path to one or more locations.
    [Parameter(Mandatory=$true,
               Position=0,
               HelpMessage="Path to import file - Example: C:\Source.csv ")]
    [ValidateNotNullOrEmpty()]
    [string]
    $SourceCsv,

    [Parameter(HelpMessage="Delimiter for Import-Csv. Default: , ")]
    [System.Char]
    $Delimiter = ",",

    [Parameter(HelpMessage="Encoding for Import-Csv. Default: Default ")]
    [string]
    $Encoding = "Default",

    [Parameter(HelpMessage="Headername for column containing FirstName")]
    [string]
    $HeaderFirstName = "FirstName",

    [Parameter(HelpMessage="Headername for column containing LastName")]
    [string]
    $HeaderLastName = "LastName",

    [Parameter(HelpMessage="Headername for column containing Birthday")]
    [string]
    $HeaderBirthday = "Birthday",

    [Parameter(HelpMessage="Headername for column containing first E-MailAddress")]
    [string]
    $HeaderMail = "Email1Address"
)

if (!(Test-Path -Path $SourceCsv)) {
    Write-Host -ForegroundColor Red "File not found ..."
    return
}

# read csv to object
$csv_read = Import-Csv -Path $SourceCsv -Delimiter $Delimiter -Encoding $Encoding -ErrorAction Stop

# Convert DataTypes
foreach ($entry in $csv_read) {
    [string]$entry.$HeaderFirstName = $entry.$HeaderFirstName
    [string]$entry.$HeaderLastName = $entry.$HeaderLastName
    [datetime]$entry.$HeaderBirthday = Get-Date -Date $entry.$HeaderBirthday
    [string]$entry.$HeaderMail = $entry.$HeaderMail
}

# Show GridView for contact selection
$csv_selected = $csv_read | Out-GridView -Title "Welche Kontakte sollen importiert werden?" -PassThru

if (!$csv_selected) {
    Write-Host -ForegroundColor Yellow "Keine Kontakte ausgewählt ..."
    return
}

# Create ComObject for Outlook
$objOutlook = New-Object –ComObject Outlook.Application  

# read contacts from AddressBook
$OutlookContacts = $objOutlook.Session.GetDefaultFolder(10).items

# start
foreach ($csv_contact in $csv_selected) {
    
    # search existing entry by mail
    $search = $OutlookContacts | Where-Object {$_.Email1Address -eq $csv_contact.$HeaderMail}

    if ($search) {
        if ($search.count -gt 1) {
            # if found one or more matching, show GridView for selection
            $selected = $search | Out-GridView -PassThru -Title "In welchem Kontakt soll der Geburtstag von $($csv_contact.$HeaderMail) gespeichert werden?"
        } else {
            # if found only one
            $selected = $search
        }
        Write-Host "Verwende vorhandenen Kontakt für $($csv_contact.$HeaderMail) ... " -NoNewline

        # check if not equal
        if ($csv_contact.$HeaderBirthday -ne $selected.Birthday) {
            $selected.Birthday = $csv_contact..$HeaderBirthday    
            Write-Host -ForegroundColor Green "OK."
        } else {
            Write-Host -ForegroundColor Blue "Allready there."
        }

    } else {
        # create new contact entry
        # https://docs.microsoft.com/de-de/office/vba/api/outlook.olitemtype

        Write-Host "Erstelle neuen Kontakt $($csv_contact.$HeaderFirstName) $($csv_contact.$HeaderLastName) $($csv_contact.$HeaderMail) ... " -NoNewline

        $NewContact = $objOutlook.CreateItem(2)
        $NewContact.Birthday = $csv_contact.$HeaderBirthday
        $NewContact.FirstName = $csv_contact.$HeaderFirstName
        $NewContact.LastName =  $csv_contact.$HeaderLastName
        $NewContact.Email1Address = $csv_contact.$HeaderMail
        $NewContact.Close(0)

        Write-Host -ForegroundColor Green "OK."

    }
}
#$objOutlook.Session.GetDefaultFolder(10).items | Where-Object {$_.Birthday –ne ([datetime]”1/1/4501”)}| Format-Table -AutoSize Fullname, Firstname, Lastname, Birthday, Email1Address
$objOutlook.Quit()
return