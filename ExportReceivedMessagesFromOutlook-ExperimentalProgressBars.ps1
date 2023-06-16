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

        $dirPath = 'C:\temp\received'

        Write-Host "Connecting to Outlook. Please wait..." -ForegroundColor Yellow

        $count = 0
        $readDate = Read-Host "Please provide the 'FROM' date (E.g. I want ever message FROM 1st December up until Today) in the format 30/12/2022:"
        $date = Get-Date $readDate

        Write-Host "Querying Outlook on date provided..." -ForegroundColor Yellow
        $totalEmails = (Get-OutlookInbox | Where-Object {$_.ReceivedTime -GE $date}).Count
        Write-Host $totalEmails "received messages found in Inbox folder. From" $date"."
        $receivedMail = Get-OutlookInbox | Where-Object {$_.ReceivedTime -GE $date} | ForEach-Object {
            Write-Progress -Activity "Retrieving emails" -Status "Processing email $count of $totalEmails" -PercentComplete (($count / $totalEmails) * 100)
            $count++
            $_
        }

        foreach ($message in $receivedMail) {
            Write-Progress -PercentComplete (($i*100)/$totalEmails) -Activity "Saving messages." -Status "Please wait...  $(([math]::Round((($i)/$totalEmails * 100),0)))%";
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


