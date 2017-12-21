
Function HWInv_Full($comp){
    
    $HardwareInventoryID = '{00000000-0000-0000-0000-000000000001}'
    Get-WmiObject -ComputerName $comp -Namespace   'Root\CCM\INVAGT' -Class 'InventoryActionStatus' -Filter "InventoryActionID='$HardwareInventoryID'" | Remove-WmiObject
    Start-Sleep -s 5
    Invoke-WmiMethod -computername $comp -Namespace root\CCM -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000001}"
    
    }
    Function Check_Drive_Space ($comp){
    
    $disk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" |
    Foreach-Object {$_.Size,$_.FreeSpace}
    
    $disk = ([wmi]"\\$comp\root\cimv2:Win32_logicalDisk.DeviceID='c:'")
    $free_Space= "{0:#.0}" -f ($disk.FreeSpace/1GB)
    $free_Space
    if ($free_space -le 2) {
    write-output "Cleaning"
    sleep -s 20
      remove-item -force -Recurse "\\$comp\c$\windows\temp\*" -Verbose
    }
    
    }
    
    $computers=  get-content "c:\computers.txt"
    
    foreach($comp in $computers){
    
        if ((Test-Connection $comp -Count 1 -Quiet) -eq $true){
            Check_Drive_Space ($comp)
            HWInv_Full($comp)
    }
    }

    