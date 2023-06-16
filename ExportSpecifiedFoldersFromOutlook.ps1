<#
.SYNOPSIS
  
.DESCRIPTION
  This script will
.OUTPUTS
  none
.NOTES
  Version:        1.0
  Author:         Dylan Reynolds
  Creation Date:  date 
  Purpose/Change: Initial script creation
.EXAMPLE
  Run the script with .\... and select either Y to continue or N to end.
#>

$Countdown = 5
$executeScript = Read-Host "Do you want do the thing (Y/N)?" 
$executeScript = $executeScript.ToUpper()

if ($executeScript -eq "Y"){
    try{
        
        Start-Outlook 
        Connect-Outlook | Out-Null 

        $dirPath = 'C:\temp\folders'

        Write-Host "Connecting to Outlook. Please wait..." -ForegroundColor Yellow


        $count = 0
        #$readDate = Read-Host "Please provide the 'TO' date in the format 30/12/2022: (E.g. I want ever messages up TO the 1st December 2022, type 30/12/2022)"
        #$date = Get-Date $readDate

        $getContent = Read-Host "Please provide the location of file you want to read: (E.g C:\temp\MailboxItems.csv)"
        $content = Import-CSV $getContent
        
        Write-Host "Reading file. Please wait..."
        foreach ($row in $content) {$folders = $row.Name}

        Write-Host "Querying Outlook on folders and dates..." -ForegroundColor Yellow
        foreach ($folder in $folders) {
            $totalEmailsInFolder = Get-OutlookFolder -Recurse -MainOnly | Where-Object {$_.Name -Match $folder}
            $totalEmailsInFolder = $totalEmailsInFolder.Items.Count
            $messages = Get-OutlookFolder -Recurse -MainOnly | Where-Object {$_.Name -Match $folder} | ForEach-Object {
                Write-Progress -Activity "Retrieving emails" -Status "Processing email $count of $totalEmailsInFolder" -PercentComplete (($count / $totalEmailsInFolder) * 100)
                $count++
                $_
            }
        }


        foreach ($message in $messages) {
            #Write-Progress -PercentComplete (($i*100)/$totalEmails) -Activity "Saving messages." -Status "Please wait...  $(([math]::Round((($i)/$totalEmails * 100),0)))%"
            $messageDate = [DateTime]$message.ReceivedTime
            $year = $messageDate.year
            $dateStr = '{0:MM-dd}' -f $message.ReceivedTime
            
            New-Item -ItemType Directory -Path $dirPath\$year\$dateStr -Force | Out-Null
            $message | Export-OutlookMessage -FileNameFormat "FROM= %SenderName% SUBJECT= %Subject% DATE= %ReceivedTime%" -OutputFolder $dirPath\$year\$dateStr -ErrorAction SilentlyContinue
            $i++
        }
        Write-Host "Completed!"
        Start-Sleep -Seconds $Countdown
    } catch{
        Write-Output "Output failed. See following error message:"
        Write-Output $_
        Start-Sleep -Seconds $Countdown
    }
    
}

