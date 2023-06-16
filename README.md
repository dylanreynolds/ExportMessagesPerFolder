# ExportMessagesPerFolder
 This script will export specified folders from a CSV file from Outlook and   save them into folders with dates they were received.

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
