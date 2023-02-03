# Connect to PnP Online
$SiteURL = "https://tam3.sharepoint.com/"
Connect-PnPOnline -Url $SiteURL -UseWebLogin

# Import CSV file
$csv = Import-Csv C:\Test.csv 

# Loop through each row of the CSV file
$csv | ForEach-Object {
  $source = $_.source
  $Dest = $_.Dest
  
  # Create URLs for the source and destination folders
  $SourceFolderURL = "$SiteURL" + "Shared Documents/" + "$source"
  $DestFolderURL = "$SiteURL" + "$Dest" + "/" + "$source"

  # Check if the destination folder already exists
  $DestFolder = Get-PnPFolder -Url $DestFolderURL -ErrorAction SilentlyContinue

  # If the destination folder does not exist, create it
  if (!$DestFolder) {
    $folder = Add-PnPFolder -Name $source -Folder $Dest -ErrorAction Stop
    Write-host -f Green "New Folder '$dest' Added!"
  } else {
    Write-host "Folder $Dest Already Exist"  
  }

  # Copy the source folder to the destination folder
  Write-Host "****************************************"
  Write-Host  "Source Location : $SourceFolderURL"
  Write-Host  "Destination Location : $DestFolderURL"
  Write-Host "****************************************"
  
  Copy-PnPFile -SourceUrl $SourceFolderURL -TargetUrl $DestFolderURL -force -ErrorAction SilentlyContinue

  # Get the number of items in the destination folder
  $destinationItemCount = (Get-PnPFolderItem -FolderSiteRelativeUrl $DestFolderURL).Count
  Write-Host "Destination Folder Item Count: $destinationItemCount"
}

# Show a message box to indicate that the script has completed
$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("Script Completed. Check the destination folder item count.", 0, "Check your count!!", 0)
