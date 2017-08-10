class HAP: MailItem {
    hidden [System.Object]$HTMLDocument
    
    [string]$ReportVer = "N/A"
    [string]$FullReportVer = "N/A"
    [string]$ReportType = "N/A"
    [string]$JobName = "N/A"

    #Call MailItem constractor before HAP
    HAP([Microsoft.Exchange.WebServices.Data.Item]$MailItem) : base([Microsoft.Exchange.WebServices.Data.Item]$MailItem){
        #Load Mail Body
        try {
            $this.HTMLDocument = New-Object HtmlAgilityPack.HtmlDocument
            $this.HTMLDocument.LoadHtml($this.body)
            $this.ReportGeneralData()
        }
        catch {
            $ErrorMSG = "{0}`n{1}" -f $_.Exception.Message,$_.InvocationInfo.PositionMessage; $EntryType = "Error"
            Write-EventLog -LogName "BackupReportMonitor" -Source "HAP" -EventId 101 -Category 1 -EntryType $EntryType -Message $ErrorMSG
            break;
        }
    }

    hidden [void]ReportGeneralData(){
        switch -Wildcard ($this.body) {
            "*Veeam*" {
                $FullVer = $this.HtmlDocument.DocumentNode.SelectSingleNode("/html/body/table/tbody/tr[3]/td").InnerText.trim().replace("&amp;", "&")
                $Title = $this.HtmlDocument.DocumentNode.SelectSingleNode("/html/body/table/tbody/tr[1]/td/table/tbody/tr[1]/td[1]").InnerText.trim().split([environment]::NewLine)[0]
                if ($FullVer -and $Title) {
                    $this.ReportVer = "Veeam"; $this.FullReportVer = $FullVer;
                    if($Title -match "Backup" -or $Title -match "Replication"){
                        $TypeANDname = $Title.split(":").trim()
                        $this.ReportType = $TypeANDname[0]; $this.JobName = $TypeANDname[1]
                    }
                    elseif ($Title -match "Configuration") {
                        $TypeANDname = $Title.split("for").trim()
                        $this.ReportType = $TypeANDname[0]; $this.JobName = $TypeANDname[1]
                    }
                } 
                break;
            }
            "*BackupExec*" {
                $this.ReportVer = "BackupExec"
                break;
            }
            Default { 
                $this.ReportVer =  "Undefined"
                $this.FullReportVer =  "Undefined"
                $this.ReportType = "Undefined"
                $this.JobName = "Undefined"
                $this.ReportType = "Undefined"
            }
        }
    }

    hidden [void]ParseReport(){

    }
}