function New-HAP{
    param(
        [Parameter(Mandatory=$true)][Microsoft.Exchange.WebServices.Data.Item]$MailItem
    )

    try {
        $HAP = [HAP]::new($MailItem)
        return $HAP
    }
    catch {
        $MSG = "{0}`n{1}" -f $_.Exception.Message,$_.InvocationInfo.PositionMessage
        Write-EventLog -LogName "BackupReportMonitor" -Source "Public" -EventId 201 -Category 2 -EntryType "Error" -Message $MSG
        break
    }
}