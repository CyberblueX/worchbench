
$apps = get-wmiobject Win32_Product

$apps | where {$_.Name -like "*paint*"}