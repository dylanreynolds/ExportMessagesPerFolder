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
        $date = [DateTime]"1-DEC-2022"
        Write-Host "Retrieving mail from Outlook. Please wait..." -ForegroundColor Yellow
        $receivedMail = Get-OutlookInbox | Where-Object {$_.ReceivedTime -GE $date}
        $countMessages = $receivedMail.count

        foreach ($message in $receivedMail) {
            Write-Progress -PercentComplete (($i*100)/$countMessages) -Activity "Saving received messages." -Status "Please wait...  $(([math]::Round((($i)/$countMessages * 100),0)))%";
            $dateStr = '{0:yyyyMMdd}' -f $message.ReceivedTime
            New-Item -ItemType Directory -Path $dirPath\$dateStr -Force | Out-Null
            $message | Export-OutlookMessage -FileNameFormat "FROM= %SenderName% SUBJECT= %Subject% DATE= %ReceivedTime%" -OutputFolder $dirPath\$dateStr -ErrorAction SilentlyContinue
            $i++
        }

        Start-Sleep -Seconds $Countdown
    } catch{
        Write-Output "Message output failed."
        Write-Output $_
        Start-Sleep -Seconds $Countdown
    }
    
}