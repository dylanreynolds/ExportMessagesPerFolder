$Date = Get-Date -Format "dd-MMM-yyyy-hhmm"
$LogPath = "C:\ProgramData\EmailExport\"
$LogFileName = $Date + ".log"

if(Test-Path $LogPath -ErrorAction SilentlyContinue) {
	Write-Host "Logging directory exists - $($LogPath)"
} else{
	New-Item -ItemType Directory -Path $LogPath -Force
}
#$Host.UI.RawUI.BufferSize = New-Object Management.Automation.Host.Size (500, 9999)

$LogDestination = $LogPath+"\"+$LogFileName
$FileCheck = Test-Path $LogDestination

If (-not($FileCheck)) {
	"This file logs all changes made to move requests in this script" | Out-File $LogDestination
    "---------------------------------------------------------------" | Out-File $LogDestination -Append
	" " | Out-File $LogDestination -Append
}
' ' | Out-File $LogDestination -Append
"### START @ $($Date) ###" | Out-File $LogDestination -Append
' ' | Out-File $LogDestination -Append



# Import the Outlook Connector module
Import-Module -Name OutlookConnector

# Connect to Outlook
Connect-Outlook | Out-Null
    
$readInput = Read-Host "Please provide the location of file you want to read: (E.g C:\temp\MailboxItems.csv)"
$inputFile = Import-CSV $readInput
$readOutput = Read-Host "Please provide the output location for operation: E.g. C:\temp"
#$uriLocation = "$($readOutput)\"
$count = 0
Write-Host "Reading file. Please wait..."
foreach ($row in $inputFile) {
    $targetFolders = $row.Name
    $folders = Get-OutlookFolder -MainOnly -Recurse | Where-Object {$_.Name -Match $targetFolders}
    foreach ($folder in $folders) {
        try {
                        
        #$output = Get-OutlookFolder -MainOnly -Recurse | Where-Object {$_.Name -Match $folder}                   
        $items = $folder.Items
        $totalItems = $items.Count
        #Write-Host "Owner" $folder.parent.name
        #$parentFolder = $folder.parent.name
        #Write-Host "Folder" $folder.name
        #$folderName = $folder.name
        
        
        
        foreach ($item in $items) {
            
            
            Write-Progress -PercentComplete (($count*100)/$totalItems) -Activity "Saving $($folder.Name) - $($item.name) " -Status "Please wait...  $(([math]::Round((($count)/$totalItems * 100),0)))%"
            $item | Export-OutlookMessage -OutputFolder "$($readOutput)" -FileNameFormat "DATE %ReceivedTime% FROM %SenderName% SUBJECT %Subject%"
            #Write-Host  $uriLocation
            $count++
        }
        
        
        #$output | Export-OutlookFolder -OutputFolder "$($readOutput)\$($folder.owner)\$($folder.Name)\" -FileNameFormat "DATE %ReceivedTime% FROM %SenderName% SUBJECT %Subject%"
        #$output.Items | Export-OutlookMessage -OutputFolder "$($readOutput)\$($folder.owner)\$($folder.Name)\" -FileNameFormat "DATE %ReceivedTime% FROM %SenderName% SUBJECT %Subject%"
        } catch {
            "$Date, Error message - $_.Exception.Message" | Out-File $LogDestination -Append
               
        }
    }
}

    



