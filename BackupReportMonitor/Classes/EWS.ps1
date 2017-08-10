class EWS{

    hidden [ExchVer]$ExchangeVersion
    hidden [Microsoft.Exchange.WebServices.Data.ServiceObject]$InboxFolder
    [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchService
    [array]$InboxItems

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
            $ErrorMSG = "{0}`n{1}" -f $_.Exception.Message,$_.InvocationInfo.PositionMessage; [EventType]$EntryType = "Error"; [int]$EventID = 101
            Write-EventLog -LogName "BackupReportMonitor" -Source "EWS" -EventId $EventID -Category 1 -EntryType $EntryType -Message $ErrorMSG
        }
    }

    [bool] BindInboxFolder(){
        try {
            $this.InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($this.exchService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
            return $true
        }
        catch {
            $ErrorMSG = "{0}`n{1}" -f $_.Exception.Message,$_.InvocationInfo.PositionMessage; [EventType]$EntryType = "Error"; [int]$EventID = 102
            Write-EventLog -LogName "BackupReportMonitor" -Source "EWS" -EventId $EventID -Category 1 -EntryType $EntryType -Message $ErrorMSG
            return $false
        }
    }

    [bool] FindItems(){
        # If Inbox folder is not empty
        if ($this.InboxFolder.TotalCount -ne 0) {
            $this.InboxItems = $this.InboxFolder.FindItems($this.InboxFolder.TotalCount)
            return $true
        } else {
            return $false
        }
    }
}