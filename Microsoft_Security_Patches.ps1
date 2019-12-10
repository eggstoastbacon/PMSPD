[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#Downloads the currents month's security updates and servicing stack updates. 
#Includes, IE 11, .NET, Servicing Stack, Security Only and Monthly Rollup.
#Compares file size headers and re-downloads corrupt files.
#Downloads the patch description to an individual text file.

#12/10/2019 Some KB Aritcles contain several sub-patches or may reference a different KB installer. Added alias detection. Alias will be written to description file.

#OS to download, Variables..
$OSList = "Server 2012 R2","Server 2008 R2","Server 2016"
$Year = get-date -format yyyy
$Month = get-date -format MM
$currentMonthDate = get-date -format yyyy-MM
$patchRepo = "D:\Installers\Downloads"

$OSList = $OSList.Replace(" ","+")
$headers= @{ 
"accept" = "application/json;odata=verbose" 
} 

foreach($OS in $OSlist){

#Microsoft catalog URL
$URL = "https://www.catalog.update.microsoft.com/Search.aspx?q=$currentMonthDate+$OS"

$Page = Invoke-WebRequest $URL
$pagelinks = $Page.links

$CurrentResults = $Pagelinks | where-object{$_.outerText -like "*$currentMonthDate*"}
    
$list = [System.Collections.Generic.List[psobject]]::new()
$i=0

#Create a list of results to parse
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
clear-variable alias -erroraction silentlycontinue

if($item.URL -like "*download.windowsupdate.com*" -and $item.Notes -notlike "*Itanium*" -and $item.Notes -notlike "*(1803)*" -and $item.Notes -notlike "*Preview*" -and $item.Notes -notlike "*Adobe Flash*" ){
$itemURL = $item.URL
if ($itemurl.Count -gt 1){
$itemURL = $item.URL[0]
}
if ($itemURL.trim() -like "*.exe*"){$ext = ".exe"}
if ($itemURL.trim() -like "*.msu*"){$ext = ".msu"}
$itemSplit = $itemURL.Split("-")
$truekb = $itemSplit[1]

if ($truekb -notlike $item.Patch -and $truekb -like "*kb*"){
$itemkb = $truekb
$alias = "true"
} else {$itemkb = $item.Patch}


$onlinefilesize = (Invoke-WebRequest -uri $itemURL -Method Head).Headers.'Content-Length'
$onlinefilesizeMB = [math]::round($onlinefilesize / 1024000 )
$descfilepath = ("D:\Installers\Downloads\Description\" + $itemkb + "-" + $systemType + $ext + ".txt")
if($alias -like "true"){
$item.Notes + ", Alias KB: " + $item.Patch + ", " + $onlinefilesizeMB + "MB" | out-file $descfilepath -ErrorAction SilentlyContinue
} else {
$item.Notes + " " + $onlinefilesizeMB + "MB" | out-file $descfilepath -ErrorAction SilentlyContinue
}

Invoke-WebRequest -Uri $itemURL -OutFile ("$patchRepo\" + $itemkb + "-" + $systemType + $ext) -TimeoutSec 600 -ErrorAction SilentlyContinue
$offlinefilesize = Get-ChildItem ("$patchRepo\" + $itemkb + "-" + $systemType + $ext) | % {[math]::ceiling($_.length)}
$Validate = (Get-AppLockerFileInformation -path ("$patchRepo\" + $itemkb + "-" + $systemType + $ext)).Publisher.PublisherName

while ($onlinefilesize -notlike $offlinefilesize -or $validate -notlike "*Microsoft*"){
Invoke-WebRequest -Uri $itemURL -OutFile ("$patchRepo\" + $itemkb + "-" + $systemType + $ext) -TimeoutSec 600 -ErrorAction SilentlyContinue
$Validate = (Get-AppLockerFileInformation -path ("$patchRepo\" + $itemkb + "-" + $systemType + $ext)).Publisher.PublisherName
$offlinefilesize = Get-ChildItem ("$patchRepo\" + $itemkb + "-" + $systemType + $ext) | % {[math]::ceiling($_.length)}
}
}
}
}







