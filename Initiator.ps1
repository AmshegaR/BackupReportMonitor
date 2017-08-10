#Requires -Version 5.0
#Requires -RunAsAdministrator
if (-Not $env:PROCESSOR_ARCHITECTURE -eq "AMD64") { Write-Warning -Message "Program not compatible with 32 bit architecture"; Break}

<#
--Create EventLog on local machine
$LogName = "BackupReportMonitor"
$EventSource = @("EWS", "HAP", "MailItem", "Initiator", "Public", "Private")
New-EventLog -LogName $LogName -Source $EventSource

--Credetial Manager https://www.powershellgallery.com/packages/CredentialManager/2.0
Install-Module -Name CredentialManager
Import-Module CredentialManager
#>

Import-Module $PSScriptRoot\BackupReportMonitor -Force

#region EWS
$EWS = Connect-EWS -ExchVer "Exchange2010_SP1" `
                   -AccountWithImpersonationRights "user@mail.com" `
                   -ImpersonationAccountPassword "password" `
                   -MailboxToImpersonate "report@mail.com"

#Bind Inbox folder
if([bool]$EWS.BindInboxFolder()){
    if (-Not ($EWS.FindItems())) {
        $MSG = "-- Inbox folder is unbelievable or empty --"
        Write-EventLog -LogName "BackupReportMonitor" -Source "Initiator" -EventId 303 -Category 3 -EntryType "Information" -Message $MSG
        break;
    }
} else {
    $ErrorMSG = "-- Cannot Bind Inbox Folder --"
    Write-EventLog -LogName "BackupReportMonitor" -Source "Initiator" -EventId 303 -Category 3 -EntryType "Error" -Message $ErrorMSG
    break;
}

<#
foreach ($MailItem in $EWS.InboxItems){
    $MailItem.Load()
}
#>

try {
    $EWS.InboxItems[0].Load()
}
catch {
    $ErrorMSG = "{0}`n{1}" -f $_.Exception.Message,$_.InvocationInfo.PositionMessage; [int]$EventID = 303
    Write-EventLog -LogName "BackupReportMonitor" -Source "Initiator" -EventId $EventID -Category 3 -EntryType "Error" -Message $ErrorMSG
    break;
}
#endregion

#region HAP + MailItem
$MailItem = $EWS.InboxItems[0]

$HAP = New-HAP -MailItem $MailItem
#$HAP.DeleteMail('MoveToDeletedItems', 'E:\trash\testArch', $EWS.ExchService)
#endregion
