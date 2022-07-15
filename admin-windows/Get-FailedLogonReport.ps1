#Script Functions
#- Reads Security Event log of all DCs
#- known IDs have his own Filter
#- Links to Google and ultimatewindowssecurity.com
#- Makes a NS Lookup

#Works only with ADModule
Import-Module ActiveDirectory -ErrorAction Stop
$DC = (Get-ADDomainController -Filter * | Select-Object Hostname).Hostname

#Alternative setting up DC list by your self
#$DC = dc01, dc02, dc03


#Grouping Users and Devices with multiple failed attemps.
#this vars filters events with less ...
$min_failed_logins = 3 

$day_start = Get-Date "$(Get-Date -Format "dd.MM.yyyy")"
$day_stop = (Get-Date "$(Get-Date -Format "dd.MM.yyyy")").AddDays(1)   
 
#Filename and Path
$Report = Join-Path -Path $env:USERPROFILE -ChildPath "$(Get-Date -Format yyyyMMdd_HHmmss)-AD_Failed_Logon_Report.html"

#old vars
#$Date = Get-date 
#$days = -1
#$DC = (Get-DomainController | select -First 1).Name
#$domain = Get-DomainController

 
$HTML=@" 
<title>Failed Login Report from $(Get-Date)</title> 
<style>
BODY{background-color :#FFFFFF;text-align: center;} 
TABLE{Border-width:thin;border-style: solid;border-color:#7d7d7d;border-collapse: collapse;align: center;float:inherit;} 
TH{border-width: 1px;padding: 10px;border-style: solid;border-color:#7d7d7d;background-color: #CCCCCC;text-align: center;} 
TD{border-width: 1px;padding: 5px;border-style: solid;border-color:#7d7d7d;background-color: Transparent} 
</style>
"@ 
 

$eventsdc_read = Get-Eventlog Security -Computer $DC -After $day_start -Before $day_stop -EntryType FailureAudit

#$eventsdc_read = Get-Eventlog security -Computer $DC -InstanceId 4625 -After (Get-Date).AddDays($days)

$eventsDC= $eventsdc_read | 
   Select MachineName,EventID,TimeGenerated,ReplacementStrings | 
   % { 
        IF ($_.EventID -eq "4776") { 
       #     $ip_address = ($_.ReplacementStrings[6])
       #     IF ($ip_address.Remove(7,$ip_address.Length-7) -eq "::ffff:") {$ip_address = $ip_address.remove(0,7)}
       #
       #     try {
       #     $dns_lookup = (Resolve-DnsName $ip_address -ErrorAction Stop).NameHost
       #     }
       #     catch {
       #     $dns_lookup = "not found"
       #     }

            New-Object PSObject -Property @{ 
            Source_Computer = $_.ReplacementStrings[2].ToLower() 
            EventID = $_.EventID
            UserName = $_.ReplacementStrings[1].ToLower() 
            IP_Address = ""
            DNS_Lookup = ""
            Date = $_.TimeGenerated
            DebugResult = '<a href="http://www.google.de/#hl=de&output=search&sclient=psy-ab&q=Windows Eventlog ID ' + $_.EventID + ' Result ' + $_.ReplacementStrings[3] + '" target="window">' + $_.ReplacementStrings[3] + '</a>'
            DebugUrl = '<a href="https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=' + $_.EventID + '" target="window">Info</a>' 
            }
        } ELSEIF ($_.EventID -eq "4768"){
            $ip_address = ($_.ReplacementStrings[9])
            IF ($ip_address.Remove(7,$ip_address.Length-7) -eq "::ffff:") {$ip_address = $ip_address.remove(0,7)}

            try {
            $dns_lookup = (Resolve-DnsName $ip_address -ErrorAction Stop).NameHost
            }
            catch {
            $dns_lookup = "not found"
            }

            $username = $_.ReplacementStrings[0] + " \ " + $_.ReplacementStrings[1]
            IF ($_.ReplacementStrings[0] -eq "" -and $_.ReplacementStrings[1] -eq "") {$username = ""}

            New-Object PSObject -Property @{ 
            Source_Computer = $_.MachineName.ToLower() 
            EventID = $_.EventID
            UserName = $username.ToLower()
            IP_Address = $ip_address
            DNS_Lookup = $dns_lookup.ToLower()
            Date = $_.TimeGenerated 
            DebugResult = '<a href="http://www.google.de/#hl=de&output=search&sclient=psy-ab&q=Windows Eventlog ID ' + $_.EventID + ' Result ' + $_.ReplacementStrings[6] + '" target="window">' + $_.ReplacementStrings[6] + '</a>'
            DebugUrl = '<a href="https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=' + $_.EventID + '" target="window">Info</a>' 
            } 
        } ELSEIF ($_.EventID -eq "4769"){
            $ip_address = ($_.ReplacementStrings[6])
            IF ($ip_address.Remove(7,$ip_address.Length-7) -eq "::ffff:") {$ip_address = $ip_address.remove(0,7)}

            try {
            $dns_lookup = (Resolve-DnsName $ip_address -ErrorAction Stop).NameHost
            }
            catch {
            $dns_lookup = "not found"
            }

            $username = $_.ReplacementStrings[0] + " \ " + $_.ReplacementStrings[1]
            IF ($_.ReplacementStrings[0] -eq "" -and $_.ReplacementStrings[1] -eq "") {$username = ""}

            New-Object PSObject -Property @{ 
            Source_Computer = $_.MachineName.ToLower()
            EventID = $_.EventID
            UserName = $username.ToLower()
            IP_Address = $ip_address
            DNS_Lookup = $dns_lookup.ToLower()
            Date = $_.TimeGenerated 
            DebugResult = '<a href="http://www.google.de/#hl=de&output=search&sclient=psy-ab&q=Windows Eventlog ID ' + $_.EventID + ' Result ' + $_.ReplacementStrings[8] + '" target="window">' + $_.ReplacementStrings[8] + '</a>'
            DebugUrl = '<a href="https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=' + $_.EventID + '" target="window">Info</a>' 
            }
        } ELSEIF ($_.EventID -eq "4771"){

            $ip_address = ($_.ReplacementStrings[6])
            IF ($ip_address.Remove(7,$ip_address.Length-7) -eq "::ffff:") {$ip_address = $ip_address.remove(0,7)}

            try {
            $dns_lookup = (Resolve-DnsName $ip_address -ErrorAction Stop).NameHost
            }
            catch {
            $dns_lookup = "not found"
            }

            New-Object PSObject -Property @{ 
            Source_Computer = $_.MachineName.ToLower() 
            EventID = $_.EventID
            UserName = $_.ReplacementStrings[0].ToLower() 
            IP_Address = $ip_address
            DNS_Lookup = $dns_lookup.ToLower()
            Date = $_.TimeGenerated 
            DebugResult = '<a href="http://www.google.de/#hl=de&output=search&sclient=psy-ab&q=Windows Eventlog ID ' + $_.EventID + ' Result ' + $_.ReplacementStrings[4] + '" target="window">' + $_.ReplacementStrings[4] + '</a>'
            DebugUrl = '<a href="https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=' + $_.EventID + '" target="window">Info</a>' 
            } 
        } ELSE {
            New-Object PSObject -Property @{ 
            Source_Computer = "" 
            EventID = $_.EventID
            UserName = ""
            IP_Address = ""
            DNS_Lookup = ""
            Date = $_.TimeGenerated
            DebugResult = '<a href="http://www.google.de/#hl=de&output=search&sclient=psy-ab&q=Windows Eventlog ID ' + $_.EventID + ' Result ' + $_.ReplacementStrings + '" target="window">' + $_.ReplacementStrings.ToLower() + '</a>'
            DebugUrl = '<a href="https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=' + $_.EventID + '" target="window">Info</a>' 
             }
        }
   } 
  

$Logon_errors = $eventsDC | Group-Object -Property Username | Where {$_.count -gt 3} | Sort-Object -Property Count -Descending | select -Property Count, Name

#Remove-Item $Report -ErrorAction SilentlyContinue
    


$html_before = '<div align="center"><h3><a align="center">List of all failed Login attempts between ' + $day_start + ' and ' + $day_stop +' ... Total: ' + $eventsdc_read.Count + '</a></h3><br>'
$html_after = '<br></div>'

$html_before2 = '<div align="center"><h3><a align="center">List of Users and Devices with more then ' + $min_failed_logins + ' failed login attemps ... Total Groups: ' + $Logon_errors.Count + '</a></h3><br>'
$html_after2 = '<br></div>'


$html_out = $eventsDC | Sort-Object -Property Date -Descending | ConvertTo-Html -Property Date,EventID,Source_Computer,UserName,IP_Address,DNS_Lookup,DebugResult,DebugUrl -head $HTML -body "<H1>Generated on $Date</H1>" -PreContent $html_before -PostContent $html_after

$html_out2 =  $Logon_errors | ConvertTo-Html -Property Count,Name -PreContent $html_before2 -PostContent $html_after2


$html_out_finish = $html_out + $html_out2

#Need this to get links working . . . 
Add-Type -AssemblyName System.Web
[System.Web.HttpUtility]::HtmlDecode($html_out_finish) | Out-File $Report -Append

Start-Process chrome.exe $Report