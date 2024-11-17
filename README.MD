


```powershell
$MyData = Get-Disk
@{
    MyData  = $MyData
    MyData2 = $MyData | Where-Object { $_ }
    MyData3 = $MyData | Where-Object { $_.OperationalStatus -eq 'Online' }
    MyData4 = $MyData | Where-Object { $_.OperationalStatus -eq 'Offline' }
}
```