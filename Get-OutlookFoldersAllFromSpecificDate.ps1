# Import the Outlook Connector module
Import-Module -Name OutlookConnector

# Connect to Outlook
Connect-Outlook | Out-Null

$count = 0
$readInput = Read-Host "Please provide the location of file you want to read: (E.g C:\temp\MailboxItems.csv)"
$inputFile = Import-CSV $readInput
$readOutput = Read-Host "Please provide the output location for operation: E.g. C:\temp"

Write-Host "Reading file. Please wait..."
foreach ($row in $inputFile) {
    $folders = $row.Name
    $totalFolders = $inputFile.count
    
    foreach ($folder in $folders) {
        Write-Progress -Activity "Retrieving messages..." -Status "Processing folder $count of $totalFolders" -PercentComplete (($count / $totalFolders) * 100)
        Get-OutlookFolder -MainOnly -Recurse | Where-Object {$_.Name -Match $folder} | Export-OutlookFolder -OutputFolder "$($readOutput)\$($folder.owner)\$($folder.Name)\" -FileNameFormat "DATE %ReceivedTime% FROM %SenderName% SUBJECT %Subject%"
        
    }
}

