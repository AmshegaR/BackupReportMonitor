Add-Type -Path "C:\Users\grish\OneDrive\GitRepos\BackupReportMonitor\BackupReportMonitor\Lib\HtmlAgilityPack.dll"
$HTMLDocument = New-Object HtmlAgilityPack.HtmlDocument
$doc = Get-Content "c:\Users\grish\OneDrive\GitRepos\BackupReportMonitor\Tests\simple2.html"
$HTMLDocument.LoadHtml("$doc")
if ($TRcoll = $HTMLDocument.DocumentNode.SelectNodes("//html/body/table/tr")) {
    $trOne = $TRcoll[0].SelectNodes(".//tr") #Bad resolve from the absolute path
}
