# Define script parameters
param(
    [Parameter(Mandatory)]
    [string]$Share,
    [Parameter(Mandatory)]
    [string]$outputPath,
    [Parameter(Mandatory)]
    [string]$DepartmentName,
    [Parameter(Mandatory)]
    [int]$threads 
)

### do not edit belows
[array]$Folders = (Get-ChildItem -Depth 4 -Path $share -Directory)
$Output = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()
$folders | ForEach-Object -ThrottleLimit $threads -Parallel {
    Write-Progress -Activity "Processing $($psitem.fullname)"

    $path = ($using:share)
    $level = $($Psitem.FullName).substring($path.length + 1).split("\").count

    $item = [PSCUstomObject]@{
        Folder = $Psitem.Fullname
        Level = $level

    }
    ($using:Output).add($item)
}

$Output | Sort-Object -Property "Folder" | Export-Excel -WorksheetName "FolderList" -Path "$Outputpath\$(Get-Date -Format "ddMMyyyy")-$DepartmentName-FolderIndex.csv" -AutoSize -AutoFilter