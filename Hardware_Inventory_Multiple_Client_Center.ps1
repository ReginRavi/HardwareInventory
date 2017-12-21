Import-Module "C:\SCCM Client Center 2 NEW\smsclictr.automation.DLL"

$computers= get-content "C:\comp.txt"
Function HWInv_Full($comp){
$SCCMClient = New-Object -TypeName smsclictr.automation.SMSClient($comp)
$SCCMClient.schedules.HardwareInventory($true) 
}
foreach($comp in $computers){
    if ((Test-Connection $comp -Count 1 -Quiet) -eq $true){     
        HWInv_Full($comp)
}
}
