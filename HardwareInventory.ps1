[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $initialDirectory
$OpenFileDialog.filter = "txt (*.txt)| *.txt"
$OpenFileDialog.ShowDialog() | Out-Null
$getcomputers=$OpenFileDialog.filename
$ErrorActionPreference = "Continue"
Function Policy($comp) {
try {
    $ScheduleID = "{00000000-0000-0000-0000-000000000001}"
    $SmsClient = [wmiclass]"\\$comp\root\ccm:SMS_Client" 
    $SmsClient.TriggerSchedule($ScheduleID)   }
catch {
    $erroractionpreference = 'SilentlyContinue'
    write-host "Caught an exception:" -ForegroundColor blue
    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor black
    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
continue
}
}

function Ping-computer ($computer) {
    trap {$false; continue}
    $timeout = 1000
    $object = New-Object system.Net.NetworkInformation.Ping
    $ping_status= (($object.Send($computer, $timeout)).Status -eq 'Success')
    if ($ping_status -eq $true)
    {
        policy $computer
     }
    else 
    {
     write-host "$computer is offline"
    }   
}

foreach ($computer in $getcomputers)
{
    Ping-computer $computer

}


