

$Connection = Connect-SPOService -Url https://grafenbergschule-admin.sharepoint.com

$DeletedMC365Sites = Get-SPODeletedSite | ? {$_.URL -like "*/fa-*" -or $_.URL -like "*/ab-*" -or $_.URL -like "*/th-*" -or $_.URL -like "*/pj-*" -or $_.URL -like "*/kl-*" -or $_.URL -like "*/gr-*"}

$DeletedMC365Sites = Get-SPODeletedSite

if (!$DeletedMC365Sites) {
    Write-Host -ForegroundColor Yellow "Keine gelöschten Seiten vorhanden die den Kriterien entsprechen."
    Start-Sleep -Seconds 10
    Break

}

$SelectedSite = $DeletedMC365Sites | Out-GridView -PassThru -Title "Bitte Seiten zum entgültigen löschen auswählen:"

if ($SelectedSite) {
    $SelectedSite | Remove-SPODeletedSite
} else {
    Write-Host -ForegroundColor Yellow "Keine Seiten ausgewählt."
    Start-Sleep -Seconds 10
    Break
}


Start-Sleep -Seconds 10