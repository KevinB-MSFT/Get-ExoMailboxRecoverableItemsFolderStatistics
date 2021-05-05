#########################################################################################
# LEGAL DISCLAIMER
# This Sample Code is provided for the purpose of illustration only and is not
# intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
# RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
# EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
# MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
# nonexclusive, royalty-free right to use and modify the Sample Code and to
# reproduce and distribute the object code form of the Sample Code, provided
# that You agree: (i) to not use Our name, logo, or trademarks to market Your
# software product in which the Sample Code is embedded; (ii) to include a valid
# copyright notice on Your software product in which the Sample Code is embedded;
# and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
# against any claims or lawsuits, including attorneys’ fees, that arise or result
# from the use or distribution of the Sample Code.
# 
# This posting is provided "AS IS" with no warranties, and confers no rights. Use
# of included script samples are subject to the terms specified at 
# https://www.microsoft.com/en-us/legal/intellectualproperty/copyright/default.aspx.
#
# Returns all EXO mailboxes and their Recoverable Items folder quota, size, and other data relevant to troubleshooting Recoverable Items folder size issues
# Get-ExoMailboxRecoverableItemsFolderStatistics.ps1
#  
# Created by: Kevin Bloom Kevin.Bloom@Microsoft.com  5/3/2021
#
#########################################################################################
#
#########################################################################################

##Define variables and constants
#Creates a hash table to collect and gather all of the results
$ColRecords = @()
#Gets all Mailboxes and returns relevant properties 
$Mbxs = Get-EXOMailbox -ResultSize unlimited -Properties primarysmtpaddress,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,RecoverableItemsQuota,RecoverableItemsWarningQuota,CalendarLoggingQuota,ArchiveQuota,ArchiveWarningQuota,ArchiveStatus,ArchiveState,AutoExpandingArchiveEnabled | select primarysmtpaddress,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,RecoverableItemsQuota,RecoverableItemsWarningQuota,CalendarLoggingQuota,ArchiveQuota,ArchiveWarningQuota,ArchiveStatus,ArchiveState,AutoExpandingArchiveEnabled

#Loops through all of the mailboxes and gets their Recoverable Items Folder size
foreach ($Mbx in $Mbxs)
{
    #Gets the Mailbox Recoverable Items Folder and sub-folders sizes
    $MbxRIFStats = Get-EXOMailboxFolderStatistics -Identity $Mbx.PrimarySmtpAddress -Folderscope recoverableitems | select name,Foldersize,FolderAndSubfolderSize
    #Loops through the folders and only includes the Recoverable Items folder
    $Rif = Foreach ($Folder in $MbxRifStats){if ($Folder.name -eq 'Recoverable Items') {$Folder | select name,Foldersize,FolderAndSubfolderSize}}
    $Record = "" | select PrimarySmtpAddress,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,RecoverableItemsQuota,RecoverableItemsWarningQuota,CalendarLoggingQuota,ArchiveQuota,ArchiveWarningQuota,ArchiveStatus,ArchiveState,AutoExpandingArchiveEnabled,RecoverableItemsFolderSize,RecoverableItemsFolderAndSubfolderSize
    $Record.PrimarySmtpAddress = $Mbx.PrimarySmtpAddress
    $Record.IssueWarningQuota = $Mbx.IssueWarningQuota
    $Record.ProhibitSendQuota = $Mbx.ProhibitSendQuota
    $Record.ProhibitSendReceiveQuota = $Mbx.ProhibitSendReceiveQuota
    $Record.CalendarLoggingQuota = $Mbx.CalendarLoggingQuota
    $Record.ArchiveQuota = $Mbx.ArchiveQuota
    $Record.ArchiveWarningQuota = $Mbx.ArchiveWarningQuota
    $Record.ArchiveStatus = $Mbx.ArchiveStatus
    $Record.ArchiveState = $Mbx.ArchiveState
    $Record.AutoExpandingArchiveEnabled = $Mbx.AutoExpandingArchiveEnabled
    $Record.RecoverableItemsWarningQuota = $Mbx.RecoverableItemsWarningQuota
    $Record.RecoverableItemsQuota = $Mbx.RecoverableItemsQuota
    $Record.RecoverableItemsFolderSize = $Rif.Foldersize
    $Record.RecoverableItemsFolderAndSubfolderSize = $Rif.FolderAndSubFolderSize
    #Adds the record to the hash table    
    $ColRecords += $Record    
}

#Exports the hash table to Csv
$ColRecords | Export-Csv .\AllMailboxesRecoverableItemsFolders.csv -NoTypeInformation