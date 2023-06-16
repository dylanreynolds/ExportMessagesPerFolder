Get-OutlookFolder -Recurse | Select Name,FullFolderPath,@{Name = "Count"; Expression = {$_.Items.Count}} | ? Count -gt 0 | Sort Count -Descending | FL