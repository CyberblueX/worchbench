function Get-ActInfo {
    $slmgr_vbs = Join-Path -Path $env:SystemRoot -ChildPath "System32\slmgr.vbs"
    cscript //Nologo $slmgr_vbs /dli        
}
function Set-ActKey {
    param (
        $Key
    )
    $slmgr_vbs = Join-Path -Path $env:SystemRoot -ChildPath "System32\slmgr.vbs"
    cscript //Nologo $slmgr_vbs /ipk $Key        
}
function Start-ActOnline {
    $slmgr_vbs = Join-Path -Path $env:SystemRoot -ChildPath "System32\slmgr.vbs"
    cscript //Nologo $slmgr_vbs /ato                
}

$ProductKey = (Get-WmiObject -query "select * from SoftwareLicensingService").OA3xOriginalProductKey

if ($productkey) {

    Write-Host "Found Key: $productkey" -ForegroundColor Green

    Get-ActInfo

    Set-ActKey -Key $ProductKey
    #Start-ActOnline

    Get-ActInfo

} else {

    Write-Host "No Key found :(" -ForegroundColor Magenta

}

Start-Sleep -Seconds 5