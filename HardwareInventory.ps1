$ErrorActionPreference = "Continue"
$computers= get-content "C:\Computers.txt"
foreach ($comp in $computers){

try {

$ScheduleID = "{00000000-0000-0000-0000-000000000001}"

$SmsClient = [wmiclass]”\\$comp\root\ccm:SMS_Client” 

$SmsClient.TriggerSchedule($ScheduleID)   }

catch {
$erroractionpreference = 'SilentlyContinue'
write-host "Caught an exception:" -ForegroundColor blue
write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor black
write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
continue
}

}