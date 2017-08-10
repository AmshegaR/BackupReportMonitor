class MailItem {
    [string]$ID
    [string]$from
    [DateTime]$sent
    [DateTime]$received
    [string]$subject
    [string]$body

    hidden [datetime]$Now
    hidden [string]$FileName

    #Format OBJ
    MailItem([Microsoft.Exchange.WebServices.Data.Item]$MailItem){
        $this.ID = $MailItem.id.uniqueid
        $this.from = $MailItem.from.address
        $this.sent = $MailItem.datetimesent
        $this.received = $MailItem.DateTimeReceived
        $this.subject = $MailItem.subject
        $this.body = $MailItem.body.text

        $this.Now = Get-Date
        $this.FileName = "$($this.Now.day)-$($this.Now.Month)-$($this.Now.Year)-$($this.Now.Hour)-$($this.Now.Minute)-$($this.Now.Second)-$($this.Now.Millisecond).html"
    }

    #Save msg as HTML file
    [bool]hidden ArchiveMail([string]$archPath){
        if (Test-Path -Path $archPath) {
            $path = $archPath + "\" + $this.FileName
            $content = $this.ToString()
            try {
                Set-Content -Value $content -Path $path
                return $true
            }
            catch {
                $ErrorMSG = "{0}`n{1}" -f $_.Exception.Message,$_.InvocationInfo.PositionMessage; $EntryType = "Error"
                Write-EventLog -LogName "BackupReportMonitor" -Source "MailItem" -EventId 102 -Category 1 -EntryType $EntryType -Message $ErrorMSG
                return $false
            }
        } else {
            $ErrorMSG = "Archive path not found"; $EntryType = "Error"
            Write-EventLog -LogName "BackupReportMonitor" -Source "MailItem" -EventId 102 -Category 1 -EntryType $EntryType -Message $ErrorMSG
            return $false
        }
    }

    #Delete mail by its ID ($this.MailItem.id)
    [void]DeleteMail([DeleteMode]$DeleteMode, [string]$archPath, [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchService){
        if($this.ArchiveMail($archPath)){
            try {
                $PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $Mode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::$DeleteMode
                $Item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($ExchService, $this.ID, $PropertySet)
                $Item.Delete($Mode)                
            }
            catch {
                $ErrorMSG = "{0}`n{1}" -f $_.Exception.Message,$_.InvocationInfo.PositionMessage; $EntryType = "Error"
                Write-EventLog -LogName "BackupReportMonitor" -Source "MailItem" -EventId 102 -Category 1 -EntryType $EntryType -Message $ErrorMSG
            }

        }
    }

    #Convert Obj properties to HTML format
    [string]ToString(){
        $string = "ID: " + $this.ID + "<br />" + `
                  "From: " + $this.from + "<br />" + `
                  "Sent: " + $this.sent + "<br />" + `
                  "Recived: " + $this.received + "<br />" + `
                  "Subject: " + $this.subject + "<br />" + `
                  "Body: " + $this.body + "<br />" + `
                  "Runtime: " + $this.Now
        return $string
    }
}