param (
    [string]$ScanPath,
    [string]$ExportPath
)

if (!($ScanPath)) {
    Write-Host "[ERROR] " -ForegroundColor Red -NoNewline
    Write-Host "NO PATH WAS PROVIDED TO SCAN" -ForegroundColor Red -BackgroundColor Yellow
    break
}

$Test = Test-Path $ScanPath
if (!($Test)) {
    Write-Host "[ERROR] " -ForegroundColor RED -NoNewline
    Write-Host "Cannot find $($Scanpath)" -ForegroundColor red -BackgroundColor yellow
    break
}
if ($ScanPath.EndsWith("\")){
    $scanpath = $scanpath.Substring(0,$ScanPath.Length-1)
}

$path = $Scanpath
$number = ($Path.split("\")).count - 1
$name = $path.split("\")[$number]

function Add-Item {
    param(
        [string]$folder,
        [string]$level,
        [string]$SizeMb,
        [string]$sizeGb,
        [int]$itemCount
    )

    $item = [PSCustomObject]@{
        Folder    = $folder
        Level     = $level
        SizeMB    = $SizeMb
        SizeGB    = $sizeGb
        ItemCount = $itemCount
    }

    return $item
}
function Add-Error {
    param(
        [string]$folder,
        [string]$Fault
    )

    $item = [PSCustomObject]@{
        Folder = $folder
        Error  = $Fault
    }

    return $item
}
function Add-Files {
    param(
        [string]$folder,
        [string]$file,
        [string]$extension
    )

    $item = [PSCustomObject]@{
        Folder    = $folder
        File      = $file
        Extention = $extension
    }

    return $item
}

$Summary = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()
$DirStats = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()
$ErrorFolders = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()
$FilesExt = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()
$exeArray = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()
$MediaArray = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()

$funcdef = ${function:Add-Item}.tostring()
$errorFunction = ${function:Add-Error}.tostring()
$FileExtentions = ${function:Add-Files}.tostring()

$Folders = Get-ChildItem -Path $Path -Directory -Recurse -Depth 4 -errorAction SilentlyContinue

$Folders | Foreach-Object -ThrottleLimit 30 -Parallel {
    Write-progress -Activity "Processing $($psItem.Fullname)"
    #Action that will run in Parallel. Reference the current object via $PSItem and bring in outside variables with $USING:varname
    ${function:Add-item} = $using:funcDef
    ${function:Add-Error} = $using:ErrorFunction
    ${function:Add-Files} = $using:FileExtentions

    try {
        $files = Get-ChildItem -path $($PSItem.Fullname) -file -Recurse -ErrorAction Stop
        #$files = Get-ChildItem -Path $psitem.fullname -File -Recurse -ErrorAction Stop
        $Size = $files | Measure-Object -Property Length -sum
    }
    catch {
        $Fault = $_
    }
    if (!($Files)) {
        if ($Files.Count -eq 0){
            $FileError = "Empty Folder"
        }
        else {
            $FileError = $Fault.exception.message
        }

        $errorfolder = Add-Error -folder $Psitem.fullname -Fault $fileError
        ($using:ErrorFolders).add($errorfolder)
    }
    else {
    
        $Mb = "{0:N2}" -f ($size.sum / 1Mb)
        $Gb = "{0:N2}" -f ($size.sum / 1Gb)
        $path = ($using:path)
        $level = $($Psitem.FullName).substring($path.length + 1).split("\").count
        $Stats = Add-Item -folder $psitem.Fullname -level $level -SizeMb $mb -sizeGb $gb -itemCount $Size.Count
        ($using:DirStats).add($stats)

        $extention = $files | Where-Object { $_.extension -eq ".pst" -or $_.extension -eq ".one" }
        $executables = $files | Where-Object {$_.extension -eq ".exe" -or $_.extension -eq ".bat" -or $_.extension -eq ".vbs"}
        $mediaFiles = $files | Where-Object {$_.extension -eq ".jpg" -or $_.extention -eq ".jpeg" -or $_.extention -eq ".bmp" -or $_.extention -eq ".png" -or $_.extention -eq ".gif" -or $_.extention -eq ".mp4" -or $_.extention -eq ".mov" -or $_.extention -eq ".avi" -or $_.extention -eq ".flv"}

        if ($Extention) {
            forEach ($item in $extention){
            $ext = Add-Files -folder $item.DirectoryName -File $item.name -extension $Item.Extension
            ($using:FilesExt).add($ext)
            }
        }
        if ($executables) {
            forEach ($item in $executables){
            $exe = Add-Files -folder $item.DirectoryName -File $item.name -extension $Item.Extension
            ($using:exeArray).add($exe)
            }
        }
        if ($mediaFiles) {
            forEach ($item in $mediaFiles){
            $media = Add-Files -folder $item.DirectoryName -File $item.name -extension $Item.Extension
            ($using:mediaArray).add($media)
            }
        }
        
    }
    Write-Progress -Activity "Processing $($psitem.fullname)" -completed


}

# Counts processing 
$Level1Count = $dirstats | Where-Object {$_.level -eq 1}
$TotalSizeMB = [PSCustomObject]@{
    Subject = "TotalSizeMB"
    Size    = "{0:N2}" -f ($Level1Count | Measure-Object -Property SizeMB -Sum).sum | out-String
    Number = ""
}
$Summary += $TotalSizeMB

$TotalSizeGB = [PSCustomObject]@{
    Subject = "TotalSizeGB"
    Size    = "{0:N2}" -f ($Level1Count | Measure-Object -Property SizeGB -Sum).sum | out-String
    Number = ""
}
$Summary += $TotalSizeGB

$TotalSizeTB = [PSCustomObject]@{
    Subject = "TotalSizeTB"
    Size    = "{0:N2}" -f ($Level1Count | Measure-Object -Property sizeGB -Sum).sum /1024 | out-String
    Number = ""
}
$Summary += $TotalSizeTB

$ItemCount = [PSCustomObject]@{
    Subject = "itemCount"
    Size = ""
    Number = ($Dirstats | Measure-Object -Property itemcount -Sum).sum | out-String
}
$Summary += $ItemCount

$accessDenied = $errorfolders |Where-Object {$_.Error -like "*Access*denied*"}
$Denied = [PSCustomObject]@{
    Subject = "Error - Access Denied"
    Size = ""
    Number = $accessDenied.count
}
$Summary += $Denied

$Media = [PSCustomObject]@{
    Subject = "Photo or Video"
    Size = ""
    Number = $MediaArray.count
}
$Summary += $media

$AppData = [PSCustomObject]@{
    Subject = "Potential Application Data"
    Size = ""
    Number = $exeArray.count
}
$Summary += $Appdata

$pst = $filesExt | Where-Object {$_.extention -like "*.pst"}
$PSTFiles = [PSCustomObject]@{
    Subject = "PSTFiles"
    Size = ""
    Number = $pst.count
}
$Summary += $PSTFiles

$OneNote = $filesExt | Where-Object {$_.extention -like "*.one"}
$OneNoteFiles = [PSCustomObject]@{
    Subject = "OneNote Files"
    Size = ""
    Number = $OneNote.count
}
$Summary += $OneNoteFiles


$summary | Export-Excel -WorksheetName "Summary" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
$Dirstats | Export-Excel -WorksheetName "All Folders" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
$Dirstats | Where-Object {$_.Level -eq 1}  | Export-Excel -WorksheetName "Level 1" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
$Dirstats | Where-Object {$_.Level -eq 2}  | Export-Excel -WorksheetName "Level 2" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
$Dirstats | Where-Object {$_.Level -eq 3}  | Export-Excel -WorksheetName "Level 3" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
$Dirstats | Where-Object {$_.Level -eq 4}  | Export-Excel -WorksheetName "Level 4" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
$ErrorFolders | Export-Excel -WorksheetName "Errors" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
$FilesExt | Export-Excel -Worksheetname "PST - OneNote" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
$ExeArray | Export-Excel -Worksheetname "Executables" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
$mediaArray | Export-Excel -Worksheetname "MediaFiles" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow


