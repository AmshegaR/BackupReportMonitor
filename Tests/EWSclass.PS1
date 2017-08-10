Add-Type -Path "C:\Users\grish\OneDrive\GitRepos\BackupReportMonitor\BackupReportMonitor\Lib\Microsoft.Exchange.WebServices.dll"

. ..\BackupReportMonitor\Enums\ExchVer.ps1
. ..\BackupReportMonitor\Enums\EventType.ps1

class EWS{

    hidden [ExchVer]$ExchangeVersion
    hidden [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchService
    hidden [Microsoft.Exchange.WebServices.Data.ServiceObject]$InboxFolder
    hidden [System.Object]$MailItems

    hidden [string]$EventCategory = 1
    hidden [string]$EventSource = "EWS"

    EWS([ExchVer]$ExchVer,
        [string]$AccountWithImpersonationRights,
        [string]$ImpersonationAccountPassword,
        [string]$MailboxToImpersonate){
        try {
            #Exchange Version
            $this.ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchVer

            #Exchange service obj
            $this.ExchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($this.ExchangeVersion)
            $this.ExchService.Credentials = New-Object Net.NetworkCredential($AccountWithImpersonationRights, $ImpersonationAccountPassword)
            $this.ExchService.AutodiscoverUrl($AccountWithImpersonationRights ,{$true})
            $this.ExchService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxToImpersonate)
        }
        catch {
            # Write Error to Event Log and terminate the program
            $ErrorMSG = "{0}`r`n{1}" -f $_.Exception.Message,$_.InvocationInfo.PositionMessage; [EventType]$EntryType = "Error"; [int]$EventID = 101
            $this.WriteLog($EntryType, $ErrorMSG, $EventID)
            Break
        }
    }

    hidden [void]WriteLog([EventType]$EntryType, $ErrorMSG, [int]$EventID){
        Write-EventLog -LogName "BackupReportMonitor" -Source $this.EventSource -EventId $EventID -Category $this.EventCategory -EntryType $EntryType -Message $ErrorMSG
    }

    [bool] BindInboxFolder(){
        $this.InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($this.exchService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
        # If Inbox folder is not empty
        if ($this.InboxFolder.TotalCount -ne 0) {
            $this.MailItems = $this.InboxFolder.FindItems($this.InboxFolder.TotalCount)
            return $true
        } else {
            return $false
        }
    }
}

$EWS = [EWS]::new("Exchange2010_SP1", "alexl@bynetos.co.il", "Lsd25gdmt@23", "backup@bynetos.co.il")
if(-Not ($EWS.BindInboxFolder())){
    $MSG = "--Inbox folder is empty--"
    Write-EventLog -LogName "BackupReportMonitor" -Source "Script" -EventId 303 -Category 3 -EntryType "Information" -Message $MSG
    break;
}