# Select module installation folder between PowerShell module paths
$itemsListArray = $env:PSModulePath -split ';'

$itemsList = @{}
$i = 0
$itemsListArray | %{
    $key = $i++
    $itemsList[$key] =  $_
}

$itemsList.Keys | Sort-Object | %{ Write-Host $_":" $itemsList[$_] }

$itemIndex = $itemsListArray.Count
do {
    Write-Host -NoNewline -ForegroundColor green "Please select destination path: "
    [int]$itemIndex = Read-Host
} until ($itemsList.ContainsKey($itemIndex)) 

$Dest = Join-Path -Path $itemsList[[int]$itemIndex] -ChildPath 'Search-VMwareKB'

if (-not (Test-Path $Dest)) {
    Write-Host "Creating destination folder `"$Dest`".."
    New-Item -ItemType Directory -Path $Dest | Out-Null
    Write-Host "Copying files.."
    Copy-Item Search-VMwareKB.* $Dest
    Write-Host "Running `"Unblock-File`" Cmdlet.."
    Unblock-File (Join-Path -Path $Dest -ChildPath 'Search-VMwareKB.psm1')
    Unblock-File (Join-Path -Path $Dest -ChildPath 'Search-VMwareKB.ps1xml')
    Write-Host "Module installed to `"$Dest`" successfully."
    Write-Host "Importing module `"Search-VMwareKB`".."
    Import-Module Search-VMwareKB
}
else {
    Write-Host -ForegroundColor red "Destination folder `"$Dest`" exists. Please remove it and try again."
}
