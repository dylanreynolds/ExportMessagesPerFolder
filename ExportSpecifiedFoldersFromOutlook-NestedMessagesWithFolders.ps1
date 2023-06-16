<#
.SYNOPSIS
  Export email messages with attachments from Outlook desktop client
.DESCRIPTION
  This script will export specified folders from a CSV file from Outlook and
  save them into folders with dates they were received
.OUTPUTS
  .EML files including attachments for that file, into respective dated folders and parent folder
.NOTES
  Version:        1.5
  Author:         Dylan Reynolds
  Creation Date:  27/2/23 
  Purpose/Change: Added Logging
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

$Countdown = 5
$executeScript = Read-Host "Do you want continue with script execution (Y/N)?" 
$executeScript = $executeScript.ToUpper()

if ($executeScript -eq "Y"){
    try{
        
        Start-Outlook 
        Connect-Outlook | Out-Null 

        $dirPath = 'C:\temp'
        Write-Host "Output location $dirPath" -ForegroundColor Yellow
        Write-Host "Connecting to Outlook. Please wait..." -ForegroundColor Yellow

        $count = 0
        $folders = $null
        #$readDate = Read-Host "Please provide the 'TO' date in the format 30/12/2022: (E.g. I want ever messages up TO the 1st December 2022, type 30/12/2022)"
        #$date = Get-Date $readDate

        $getContent = Read-Host "Please provide the location of file you want to read: (E.g C:\temp\MailboxItems.csv)"
        try {
            $content = Import-CSV $getContent
        } catch {
            Write-Host "Unable to read file, ensure you file selected is the correct format and try again" -ForegroundColor Red
            $_
        }
        
        Write-Host "Reading file. Please wait..."
        foreach ($row in $content) {
            $folders = $row.Name
            $count = 0
            $totalEmailsInFolder = $null
            Write-Host "Querying Outlook folder: $folders..." -ForegroundColor Yellow
            foreach ($folder in $folders) {
                $totalEmailsInFolder = Get-OutlookFolder -Recurse -MainOnly | Where-Object {$_.Name -Match $folder}
                $messages = $totalEmailsInFolder.Items
                $totalEmailsInFolder = $messages.Count
                
                $messages | ForEach-Object {
                    Write-Progress -Activity "Processing message from $folder" -Status "Processing email $count of $totalEmailsInFolder" -PercentComplete (($count / $totalEmailsInFolder) * 100)
                    $count++
                    #$_
                }
                $count = 0
                try {
                    foreach ($message in $messages) {
                        $totalMessages = $messages.Count
                        Write-Progress -Activity "Saving messages from $folder" -Status "Saving email $count of $totalMessages" -PercentComplete (($count / $totalEmailsInFolder) * 100)
                        try {
                            $messageDate = [DateTime]$message.ReceivedTime
                            $year = $messageDate.year
                            $dateStr = '{0:dd-MM}' -f $message.ReceivedTime
                        } catch {
                            $messageDate = "BadDate"
                            $year = "BadYear"
                            $dateStr = "BadStr"
                        }
                        
                        $parentFolder = $message.Parent.Name
                        
                        New-Item -ItemType Directory -Path $dirPath\$parentFolder\$year\$dateStr -Force | Out-Null
                        $message | Export-OutlookMessage -FileNameFormat "FROM= %SenderName% SUBJECT= %Subject% DATE= %ReceivedTime%" -OutputFolder $dirPath\$parentFolder\$year\$dateStr
                        $count++
                    }
                } catch {
                    Write-Host "Error with message:" $_

                }
                
            }
        }



        Write-Host "Completed!"
        Start-Sleep -Seconds $Countdown
    } catch{
        Write-Output "Output failed. See following error message:"
        Write-Output $_
        Start-Sleep -Seconds $Countdown
    }
    
}

