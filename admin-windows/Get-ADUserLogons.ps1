Import-Module ActiveDirectory

$alluserhistory = @()
$startDate = (get-date).AddDays(-1)
$startDate = (get-date).AddHours(-1)
$DCs = Get-ADDomainController -Filter *
foreach ($DC in $DCs){
$logonevents = Get-Eventlog -LogName Security -InstanceID 4768 -after $startDate -ComputerName $dc.HostName
foreach ($event in $logonevents){
if ($event.ReplacementStrings[0] -notlike '*$') {
$userhistory = New-Object PSObject -Property @{
UserName = $event.ReplacementStrings[0]
IPAddress = $event.ReplacementStrings[9]
Date = $event.TimeGenerated
DC = $dc.Name
}
$alluserhistory += $userhistory
}
}
}
$alluserhistory