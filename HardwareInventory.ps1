$comp="localhost"
$HardwareInventoryID = '{00000000-0000-0000-0000-000000000001}'
Get-CimInstance -Namespace 'Root\CCM\INVAGT' -ClassName 'InventoryActionStatus' -Filter "InventoryActionID='$HardwareInventoryID'" | Remove-CimInstance
