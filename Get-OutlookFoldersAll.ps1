$getContent = Read-Host "Please provide the location of file you want to read: (E.g C:\temp\MailboxItems.csv)"
$content = Import-CSV $getContent
$count = 0
$totalFolders = 0

Write-Host "Reading file. Please wait..."
foreach ($row in $content) {$folders = $row.Name; $totalFolders++}

Write-Host "Querying Outlook on folders and dates..." -ForegroundColor Yellow
foreach ($folder in $folders) {
    Get-OutlookFolder -Recurse -MainOnly | Where-Object {$_.Name -Match $folder} | ForEach-Object {
        Write-Progress -Activity "Retrieving messages..." -Status "Processing folder $count of $totalFolders" -PercentComplete (($count / $totalFolders) * 100)
        Get-OutlookFolder -Recurse -MainOnly | Where-Object {$_.Name -Match $folder} | Export-OutlookFolder -FileNameFormat "DATE= %ReceivedTime% FROM= %SenderName% SUBJECT= %Subject% " -OutputFolder "C:\temp\"
        $count++
    }
}