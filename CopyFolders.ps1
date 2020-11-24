#Connect to PnP Online
$SiteURL = "https://tam3.sharepoint.com/"
Connect-PnPOnline -Url $SiteURL -UseWebLogin

$csv = Import-Csv C:\Test.csv 
$csv | ForEach-Object {
    $source = $_.source
    $Dest = $_.Dest
   }

$i = 0
$csv | ForEach-Object {
$source = $_.source
$Dest = $_.Dest 

$SourceFolderURL = $siteurl + "Shared Documents/" + $source
$folderurl = $Dest + "/" + $source

$Check = "/" + "Shared Documents/" + $source

####Creating folder
Try {
 $folder = Add-PnPFolder -Name $source -Folder $Dest -ErrorAction Stop
    Write-host -f Green "New Folder '$dest' Added!" 
	$TargetFolderURL ="/" + $Dest + "/" + $source

Write-Host "****************************************"
Write-Host  "Source Location : $SourceFolderURL"
Write-Host  "Destination Location $TargetFolderURL"
Write-Host "****************************************"

Copy-PnPFile -SourceUrl $SourceFolderURL -TargetUrl $TargetFolderURL -force -ErrorAction SilentlyContinue -ErrorVariable error

$D = Resolve-PnPFolder -SiteRelativePath $TargetFolderURL 
$s = Resolve-PnPFolder -SiteRelativePath $Check 
}

catch {
 $2 = Resolve-PnPFolder -SiteRelativePath $FolderURL
 Write-host "Folder $Dest Already Exist"  
 $TargetFolderURL ="/" + $Dest + "/" + $source


Write-Host "****************************************"
Write-Host  "Source Location : $SourceFolderURL"
Write-Host  "Destination Location $TargetFolderURL"
Write-Host "****************************************"

Copy-PnPFile -SourceUrl $SourceFolderURL -TargetUrl $TargetFolderURL -force -ErrorAction SilentlyContinue -ErrorVariable error


$D = Resolve-PnPFolder -SiteRelativePath $TargetFolderURL 
$s = Resolve-PnPFolder -SiteRelativePath $Check 
}
$count

$D , $S |out-gridview
$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("Click OK to continue, if you Count equals.", 0, "Check your count!!", 0)
}


