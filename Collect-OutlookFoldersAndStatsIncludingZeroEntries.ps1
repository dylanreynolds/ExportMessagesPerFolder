<#
.SYNOPSIS
  Collects all folders including folders with zero items from Outlook client
.DESCRIPTION
  This script will collect all folders even if they zero items and send the info about the folder name
  and statistics like total number of items into a CSV file.  From there the user can edit the CSV file
  and change what folders are exported by Export-MessagePerFolderFromOutlook
  save them into folders with dates they were received.

  Its also recommended that you setup your outlook client to download all mailbox
  items - I.e. change offline cached mode from download only 1 year, change to download
  all items.  This will take time for larger mailboxes, also be aware that you might
  reach the max allowable PST file size. (50GB)
.OUTPUTS
  .EML files including attachments for that file, into respective dated folders and parent folder
.NOTES
  Version:        1.7
  Author:         Dylan Reynolds
  Creation Date:  27/2/23 
  Purpose/Change: Better logging and error handling
.DEPENDENCIES
  Requires that you run "Collect-OutlookFoldersAndStatsIncludingZeroEntries" or "Collect-OutlookFoldersAndStats"
  prior to running this script as it will list the folders in a csv file.  Edit this CSV file to exclude
  folders you don't want to download.

  Requires Install-Script OutlookConnector

  Require outlook profile be created in Outlook(32bit) and uses Powershell (x86) 

  Requires the REG Key Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE
  Ensure the Reg String for (Default) and Path both point to C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE
  Or your Outlook.exe instance.  If neither keys exists, go ahead and create them.

  Please check OutlookConnector is available and you have elevated priviledges to install and run
  See author for details: https://github.com/iricigor/OutlookConnector
.EXAMPLE
  Run the script with .\ExportSpecifiedFoldersFromOutlook and select either Y to continue or N to end.
#>

$Date = Get-Date -Format "dd-MMM-yyyy-hhmm"
$LogPath = "C:\ProgramData\EmailExport\"
$LogFileName = "CollectOutlook-" + $Date + ".log"

if(Test-Path $LogPath -ErrorAction SilentlyContinue) {
	Write-Host -ForegroundColor "Yellow" "Log directory exists - $($LogPath + $LogFileName) - Skipping!"
} else{
	New-Item -ItemType Directory -Path $LogPath -Force
}
#$Host.UI.RawUI.BufferSize = New-Object Management.Automation.Host.Size (500, 9999)

$LogDestination = $LogPath + $LogFileName
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

$csvLocation = "C:\temp\MailboxItems.csv"

Write-Host "Writing to $csvLocation, please wait..."

try {
    Get-OutlookFolder -Recurse | Select-Object Name,FullFolderPath,@{Name = "Count"; Expression = {$_.Items.Count}}`
    | Sort-Object Count -Descending | Export-CSV $csvLocation
    Write-Host -ForegroundColor "Green" "Completed! Check $csvLocation for your file."
} catch {
    $Date = Get-Date -Format "dd-MMM-yyyy-hhmm"
    "$Date - ERR - " + $_.Exception.Message | Out-File $LogDestination -Append
    Write-Host -ForegroundColor "Red" "Failed! Please check logs located in $LogDestination for further information."
    Write-Host $_.Exception.Message
    Start-Sleep 5
}
