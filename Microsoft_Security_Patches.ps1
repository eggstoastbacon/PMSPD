[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$patchRepo = "D:\Installers\Downloads"

$headers= @{ 
            "accept" = "application/json;odata=verbose" 
            } 
$OSList = "Server+2012+R2","Server+2008+R2","Server+2016"
#$OSLIst = "Server+2012+R2"
foreach($OS in $OSlist){
$Year = get-date -format yyyy
$Month = get-date -format MM
$currentMonthDate = get-date -format yyyy-MM

$URL = "https://www.catalog.update.microsoft.com/Search.aspx?q=$currentMonthDate+$OS"

$Page = Invoke-WebRequest $URL
$pagelinks = $Page.links

$CurrentResults = $Pagelinks | where-object{$_.outerText -like "*$currentMonthDate*"}
    
$list = [System.Collections.Generic.List[psobject]]::new()
$i=0

foreach($currentResult in $currentResults){
$currentResult
$post = @{ size = 0; updateID = $currentResult.ID.replace('_link',''); uidInfo = $currentResult.ID.replace('_link','')} | ConvertTo-Json -Compress
$postBody = @{ updateIDs = "[$post]" } 
$i = $i+1
$list.AddRange(@(
  [pscustomobject]@{id = $i;Patch = ([regex]'KB\d{7}').match($CurrentResult.outerText).Value;Notes=$CurrentResult.outerText;KBID=$CurrentResult.ID.replace('_link','');`
  URL=(Invoke-WebRequest -Uri 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx' -Method Post -Body $postBody |
                Select-Object -ExpandProperty Content |
                Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" | 
                ForEach-Object { $_.matches.value })}
) -as [psobject[]])
}

$systemType = $OS.Replace("+","_")

ForEach ( $item in $list ) {
if($item.URL -like "*download.windowsupdate.com*" -and $item.Notes -notlike "*Itanium*" -and $item.Notes -notlike "*(1803)*" -and $item.Notes -notlike "*Preview*" -and $item.Notes -notlike "*Adobe Flash*" ){
$itemURL = $item.URL | out-string
if ($itemURL.trim() -like "*.exe*"){$ext = ".exe"}
if ($itemURL.trim() -like "*.msu*"){$ext = ".msu"}
$onlinefilesize = (Invoke-WebRequest -uri $itemURL.trim() -Method Head).Headers.'Content-Length'
$onlinefilesizeMB = [math]::round($onlinefilesize / 1024000 )
$descfilepath = ("$patchRepo\" + $item.Patch + "-" + $systemType + $ext + ".txt")
$item.Notes + " " + $onlinefilesizeMB + "MB" | out-file $descfilepath -ErrorAction SilentlyContinue

Invoke-WebRequest -Uri $itemURL.trim() -OutFile ("$patchRepo\" + $item.Patch + "-" + $systemType + $ext) -TimeoutSec 600 -ErrorAction SilentlyContinue
$offlinefilesize = Get-ChildItem ("$patchRepo\" + $item.Patch + "-" + $systemType + $ext) | % {[math]::ceiling($_.length)}
$Validate = (Get-AppLockerFileInformation -path ("$patchRepo\" + $item.Patch + "-" + $systemType + $ext)).Publisher.PublisherName

while ($onlinefilesize -notlike $offlinefilesize -or $validate -notlike "*Microsoft*"){
Invoke-WebRequest -Uri $itemURL.trim() -OutFile ("$patchRepo\" + $item.Patch + "-" + $systemType + $ext) -TimeoutSec 600 -ErrorAction SilentlyContinue
$Validate = (Get-AppLockerFileInformation -path ("$patchRepo\" + $item.Patch + "-" + $systemType + $ext)).Publisher.PublisherName
$offlinefilesize = Get-ChildItem ("$patchRepo\" + $item.Patch + "-" + $systemType + $ext) | % {[math]::ceiling($_.length)}
}

}
}
}





