Add-Type -Path "C:\Users\grish\OneDrive\GitRepos\BackupReportMonitor\BackupReportMonitor\Lib\Microsoft.Exchange.WebServices.dll"

enum ExchVer {
    Exchange2007
    Exchange2007_SP1
    Exchange2010
    Exchange2010_SP1
    Exchange2010_SP2
    Exchange2013
    Exchange2013_SP1
}

$AccountWithImpersonationRights = "alexl@bynetos.co.il"
$ImpersonationAccountPassword = "Lsd2dmt@23"
$MailboxToImpersonate = "backup@bynetos.co.il"
[ExchVer]$ExchVer = "Exchange2010_SP1"

$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchVer
$ExchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$ExchService.Credentials = New-Object Net.NetworkCredential($AccountWithImpersonationRights, $ImpersonationAccountPassword)
$ExchService.AutodiscoverUrl($AccountWithImpersonationRights ,{$true})
$ExchService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxToImpersonate)

$InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
[array]$InboxItems = $InboxFolder.FindItems($InboxFolder.TotalCount)