<#
.SYNOPSIS
  Export email messages with attachments from Outlook desktop client
.DESCRIPTION
  This script will export specified folders from a CSV file from Outlook and
  save them into folders with dates they were received.
.OUTPUTS
  .EML files including attachments for that file, into respective dated folders and parent folder
.NOTES
  Version:        1.7
  Author:         Dylan Reynolds
  Creation Date:  27/2/23 
  Purpose/Change: Better error handling and logging specific messages that failed to collect
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

$Date = Get-Date -Format "dd-MMM-yyyy"
$Hours = Get-Date -Format "dd-MMM-yyyy-hhmm"
$LogPath = "C:\ProgramData\EmailExport\"
$LogFileName = "ExportOutlook-" + $Date + ".log"

if(Test-Path $LogPath -ErrorAction SilentlyContinue) {
	Write-Host "Logging directory exists - $($LogPath)"
} else{
	New-Item -ItemType Directory -Path $LogPath -Force
}

$LogDestination = $LogPath+"\"+$LogFileName
$FileCheck = Test-Path $LogDestination

If (-not($FileCheck)) {
	"This file logs all changes made to move requests in this script" | Out-File $LogDestination
    "---------------------------------------------------------------" | Out-File $LogDestination -Append
	" " | Out-File $LogDestination -Append
}
' ' | Out-File $LogDestination -Append
"### START @ $($Hours) ###" | Out-File $LogDestination -Append
' ' | Out-File $LogDestination -Append

$Countdown = 5

$executeScript = Read-Host "Do you want continue with script execution (Y/N)?" 
$executeScript = $executeScript.ToUpper()

if ($executeScript -eq "Y") {
  try {
        $count = 0
        $folders = $null
        # change the following if you require a different save path
        $dirPath = 'C:\temp'

        # start up for Powershell module OutlookConnector.psm1 (Start-Outlook and Connect-Outlook)
        Import-Module OutlookConnector
        Start-Outlook 
        
        Connect-Outlook | Out-Null 
        
        Write-Host "Output location $dirPath" -ForegroundColor Yellow
        Write-Host "Connecting to Outlook. Please wait..." -ForegroundColor Yellow
        $getContent = Read-Host "Please provide the location of file you want to read: (E.g C:\temp\MailboxItems.csv)"
        try {
            $content = Import-CSV $getContent
        } catch {
            Write-Host "Unable to read file, ensure you file selected is the correct format and try again" -ForegroundColor Red
            "$Date - ERR - " + $_.Exception.Message | Out-File $LogDestination -Append
        }
        Write-Host "Reading file. Please wait..."
        foreach ($row in $content) {
            $count = 0
            $totalEmailsInFolder = $null
            $folders = $row.FullFolderPath
            Write-Host "Querying Outlook folder: $folders..." -ForegroundColor Yellow
            
            foreach ($folder in $folders) {
                "$Date MSG -  Started $folder "  | Out-File $LogDestination -Append
                # Outlook Connector Get-OutlookFolder looks at all folders recursively and matches the ones we want
                $totalEmailsInFolder = Get-OutlookFolder -Recurse -MainOnly | Where-Object {$_.FolderPath -like $folder}
                $messages = $totalEmailsInFolder.Items
                $totalEmailsInFolder = $messages.Count
                $messages | ForEach-Object {
                    Write-Progress -Activity "Processing message from $folder" -Status "Processing email $count of $totalEmailsInFolder" -PercentComplete (($count / $totalEmailsInFolder) * 100)
                    $count++
                }
                $count = 0
                if ($messages.Count -ne 0) {    
                    foreach ($message in $messages) {
                        $totalMessages = $messages.Count
                        Write-Progress -Activity "Saving messages from $folder" -Status "Saving email $count of $totalMessages" -PercentComplete (($count / $totalMessages) * 100)
                        # try catch to ensure any messages that do not get a received time or has bad date gets caught
                        # any failed / skipped messages go to catch and recorded on why they were skipped, displayed in the log file
                        try {
                            $messageDate = [DateTime]$message.ReceivedTime
                            $year = $messageDate.year
                            $dateStr = '{0:MM-dd}' -f $message.ReceivedTime
                            # trim the leading \\ to just be \ for our full folder path
                            $folder = $folder.trimStart("\")
                            $Path = "$dirPath\$folder\$year\$dateStr"
                            New-Item -ItemType Directory -Path $Path -Force | Out-Null
                        }catch {
                            $Date = Get-Date -Format "dd-MMM-yyyy-hhmm"
                            "$Date - SKP " + $folder +", " + $message.Subject +", # $count" + ", " + $_.Exception.Message +" "+ $_ +",This item was not exported." | Out-File $LogDestination -Append
                             $Path = "$dirPath\$folder\BADYEAR\BADDATE"                          
                            New-Item -ItemType Directory -Path $Path -Force | Out-Null
                          } finally {
                            # Export-OutlookMessage we pipe current message, creates eml file with FROM SUBJECT DATE to the folder created
                            $message| Export-OutlookMessage -FileNameFormat "FROM %SenderName% SUBJECT %Subject% DATE %ReceivedTime%" -OutputFolder $Path -ErrorAction SilentlyContinue
                            #Write-Host $message.Subject $Path 
                            $count++
                        }                      
                    }
                    $Date = Get-Date -Format "dd-MMM-yyyy-hhmm"
                    "$Date MSG - Finished $folder - processed $count messages"  | Out-File $LogDestination -Append
                } else {
                  $Date = Get-Date -Format "dd-MMM-yyyy-hhmm"  
                  Write-Host "No items in Outlook folder: $folder...  Skipping!" -ForegroundColor Red
                  "$Date - INF - No items in Outlook folder: $folder...  Skipping!" | Out-File $LogDestination -Append
                }
              }           
        }
        # append completion
        $Date = Get-Date -Format "dd-MMM-yyyy-hhmm"
        Write-Host -ForegroundColor "Green" "Completed! Check $dirPath for your file."
        "$Date - !-------------------------------------------------------- All folders completed! Check $dirPath for your file for further details ------------------------------------------------------!" | Out-File $LogDestination -Append
        Start-Sleep -Seconds $Countdown
    } catch{
        # append errors
        $Date = Get-Date -Format "dd-MMM-yyyy-hhmm"
        "$Date - ERR - " + $_.Exception.Message +" "+ $_ | Out-File $LogDestination -Append
        Write-Host -ForegroundColor "Red" "Failed! Please check logs located in $LogPath $LogFileName  for further information."
        Start-Sleep -Seconds $Countdown
    }
}

