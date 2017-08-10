Add-Type -Path "C:\Users\grish\OneDrive\GitRepos\BackupReportMonitor\BackupReportMonitor\Lib\HtmlAgilityPack.dll"
$HTMLDocument = New-Object HtmlAgilityPack.HtmlDocument

$Backup9_5_0_8 = Get-Content "C:\Users\grish\OneDrive\temp\VeeamRepo\9.5.0.8\Backup9.5.0.8.html"
$HTMLDocument.LoadHtml($Backup9_5_0_8)
$TRcoll = $HTMLDocument.DocumentNode.SelectNodes("//html/body/table/tbody/tr")

if($TRcoll){
    $BackupFullVer = $TRcoll[-1].InnerText.Replace("&amp;", "&").trim()
    $Node = $TRcoll[0].OuterHtml
    $Table = [HtmlAgilityPack.HtmlNode]::createnode("$Node")
    $test = $Table.SelectNodes("/tr")
}





#$split = $TRcoll[0].SelectNodes(".//html/body/table/tbody/tr")
#html/body/table/tbody/tr
#https://stackoverflow.com/questions/10583926/html-agility-pack-selectnodes-from-a-node

<#
$tr = $TRcoll[0].OuterHtml
$test = [HtmlAgilityPack.HtmlNode]::createnode("$tr")
#$test.selectnodes(xpath to select first two)
$trcol = $test.SelectNodes("//tr")
#>






