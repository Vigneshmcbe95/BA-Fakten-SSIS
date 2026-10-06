Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime='2026-10-06 01:30'; EndTime='2026-10-06 04:00'} |
  Where-Object { $_.ProviderName -like 'MSSQL*' -and $_.Id -in 1101,1105,5144,5145,9002 } |
  Select-Object TimeCreated, ProviderName, Id, Message | Format-Table -Wrap
