<#
.SYNOPSIS
    Generating an Excel report with the depth of 4 levels of a file share

.PARAMETER ScanPath (Required)
Add the UNC path to scan11

.PARAMETER ExportPath (Required)
Add the path to export the Excel

.PARAMETER threads
provide the amout of threads at the same time to run. if not provided, the default 10 will be set

.NOTES
  Version:        2.0
  Author:         Alex ter Neuzen
  Creation Date:  01-08-2023
  Purpose/Change: Version 2.0 release

#>

param (
    [Parameter(Mandatory)]
    [array]$ScanPath,
    [Parameter(Mandatory)]
    [string]$ExportPath,
    [int]$threads
)

If ($PSVersiontable.PSVersion.major -ne 7){
    Write-Host "ERROR: Powershell 7.x is required" -ForegroundColor Red
    break
}

if (!(Test-Path $ExportPath)){
    Write-host "ERROR: Cannot find $($Export)" -ForegroundColor Red
    break
}

Try {
    Import-Module ImportExcel -Force -ErrorAction Stop
}
Catch {
    Write-Host "Installing IMPORTEXCEL module"
    Install-Module ImportExcel -Force
}

if (!($ScanPath)) {
    Write-Host "[ERROR] " -ForegroundColor Red -NoNewline
    Write-Host "NO PATH WAS PROVIDED TO SCAN" -ForegroundColor Red -BackgroundColor Yellow
    break
}

if (!($threads)){
    $threads = 10
}

function Add-Item {
    param(
        [string]$folder,
        [string]$level,
        [string]$SizeMb,
        [string]$sizeGb,
        [int]$itemCount,
        [string]$DateCreated,
        [string]$LastWriteTime,
        [string]$LastAccess
    )

    $item = [PSCustomObject]@{
        Folder    = $folder
        Level     = $level
        SizeMB    = $SizeMb
        SizeGB    = $sizeGb
        ItemCount = $itemCount
        CreationTime = $DateCreated
        LastWriteTime = $LastWriteTime
        LastAccessTime = $LastAccess
    }

    return $item
}
function Add-Error {
    param(
        [string]$folder,
        [string]$file,
        [string]$Fault
    )

    $item = [PSCustomObject]@{
        Folder = $folder
        File = $file
        Error  = $Fault
    }

    return $item
}
function Add-Files {
    param(
        [string]$folder,
        [string]$file,
        [string]$extension,
        [string]$size,
        [string]$DateCreated,
        [string]$LastAccess
    )

    $item = [PSCustomObject]@{
        Folder    = $folder
        File      = $file
        Extention = $extension
        Size      = $size
        CreationTime = $DateCreated
        LastAccessed = $LastAccess
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

Write-Host "-------------------------------------"
Write-host "Start of fileshare Assessment"
Write-Host "-------------------------------------"

ForEach ($path in $Scanpath) {
    $folder = $path
    if (!(Test-Path $folder)) {
        Write-Host "[ERROR] " -ForegroundColor RED -NoNewline
        Write-Host "Cannot find $($folder)" -ForegroundColor red -BackgroundColor yellow
        break
    }
    if ($folder.EndsWith("\")) {
        $folder = $folder.Substring(0, $folder.Length - 1)
    }

    $Folders = Get-ChildItem -Path $Folder -Directory -Recurse -Depth 4 -errorAction SilentlyContinue

    $folders | Foreach-Object -ThrottleLimit $threads -Parallel {
        #Write-progress -Activity "Processing $($psItem.Fullname)"

        $characters = '<','?','>','*','|',':','\','"','/'

        Write-Host "Processing $($PSItem.Fullname)..." -ForegroundColor Yellow

        ${function:Add-item} = $using:funcDef
        ${function:Add-Error} = $using:ErrorFunction
        ${function:Add-Files} = $using:FileExtentions
    
        if ($PSItem.Fullname -contains "_vti_"){
            $errorfolder = Add-Error -folder $Psitem.fullname -File "" -Fault "Folder contains _vti_ which is not allowed"
            ($using:ErrorFolders).add($errorfolder)
        }

        try {
            $files = Get-ChildItem -path $($PSItem.Fullname) -file -Recurse -ErrorAction Stop
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
    
            $errorfolder = Add-Error -folder $Psitem.fullname -File "" -Fault $fileError
            ($using:ErrorFolders).add($errorfolder)
        }
        else {
    
            foreach ($File in $files){
                if ($File.fullname.length -gt 250){
                    $errorfolder = Add-Error -folder $Psitem.fullname -file $file.fullname -Fault "Folder or file is longer than 250 characters"
                    ($using:ErrorFolders).add($errorfolder)
                }
                if ($File.name -contains $characters){
                    $ErrorName = Add-Error -folder $PSItem.Fullname -file $file.Fullname -Fault "File contains invalid characters"
                    ($using:ErrorFolders).add($ErrorName)
                }


            }
        
            $Mb = "{0:N2}" -f ($size.sum / 1Mb)
            $Gb = "{0:N2}" -f ($size.sum / 1Gb)
    
            $path = ($using:path)
            $level = $($Psitem.FullName).substring($path.length + 1).split("\").count
            $Stats = Add-Item -folder $psitem.Fullname -level $level -SizeMb $mb -sizeGb $gb -itemCount $Size.Count -DateCreated $psitem.CreationTime -LastWriteTime $psitem.LastWriteTime -LastAccess $psItem.LastAccessTime
            ($using:DirStats).add($stats)
    
            $extention = $files | Where-Object { $_.extension -eq ".pst" -or $_.extension -eq ".one" }
            $executables = $files | Where-Object {$_.extension -eq ".exe" -or $_.extension -eq ".bat" -or $_.extension -eq ".vbs"}
            $mediaFiles = $files | Where-Object {$_.extension -eq ".jpg" -or $_.extention -eq ".jpeg" -or $_.extention -eq ".bmp" -or $_.extention -eq ".png" -or $_.extention -eq ".gif" -or $_.extention -eq ".mp4" -or $_.extention -eq ".mov" -or $_.extention -eq ".avi" -or $_.extention -eq ".flv"}
    
            if ($Extention) {
                forEach ($item in $extention){
                $ext = Add-Files -folder $item.DirectoryName -File $item.name -extension $Item.Extension -size $("{0:n2}" -f (($item).length /1Mb )) -DateCreated $item.CreationTime -LastAccess $item.LastAccessTime
                ($using:FilesExt).add($ext)
                }
            }
            if ($executables) {
                forEach ($item in $executables){
                $exe = Add-Files -folder $item.DirectoryName -File $item.name -extension $Item.Extension -size $("{0:n2}" -f (($item).length /1Mb )) -DateCreated $item.CreationTime -LastAccess $item.LastAccessTime
                ($using:exeArray).add($exe)
                }
            }
            if ($mediaFiles) {
                forEach ($item in $mediaFiles){
                $media = Add-Files -folder $item.DirectoryName -File $item.name -extension $Item.Extension -size $("{0:n2}" -f (($item).length /1Mb )) -DateCreated $item.CreationTime -LastAccess $item.LastAccessTime
                ($using:mediaArray).add($media)
                }
            }
            
        }
           
        #Write-Progress -Activity "Processing $($psitem.fullname)" -completed
    
    
    }
}

Write-Host "-------------------------------------"
Write-host "Done of fileshare Assessment"
Write-Host "-------------------------------------"
Write-Host "Exporting all to XLSX"

# Counts processing 
$Level1Count = $dirstats | Where-Object {$_.level -eq 1}
$TotalSizeMB = [PSCustomObject]@{
    Subject = "TotalSizeMB"
    Size = ""
    #Size    = "{0:N2}" -f ($Level1Count | Measure-Object -Property SizeMB -Sum).sum | out-String
    Number = ""
}
$Summary += $TotalSizeMB

$TotalSizeGB = [PSCustomObject]@{
    Subject = "TotalSizeGB"
    Size = ""
    #Size    = "{0:N2}" -f ($Level1Count | Measure-Object -Property SizeGB -Sum).sum | out-String
    Number = ""
}
$Summary += $TotalSizeGB

$TotalSizeTB = [PSCustomObject]@{
    Subject = "TotalSizeTB"    
    Size = ""
    #Size    = "{0:N2}" -f ($Level1Count | Measure-Object -Property sizeGB -Sum).sum /1024 | out-String
    Number = ""
}
$Summary += $TotalSizeTB

$ItemCount = [PSCustomObject]@{
    Subject = "itemCount"
    Size = ""
    Number = $Dirstats.count
}
$Summary += $ItemCount

$accessDenied = $errorfolders |Where-Object {$_.Error -like "*Access*denied*"}
$Denied = [PSCustomObject]@{
    Subject = "Error - Access Denied"
    Size = ""
    Number = $accessDenied.count
}
$Summary += $Denied

$toolong = $errorfolders |Where-Object {$_.Error -like "*longer*"}
$long = [PSCustomObject]@{
    Subject = "Error - Possible too long"
    Size = ""
    Number = $long.count
}
$Summary += $long

$mediaArray = $mediaArray | Sort-Object -Property Folder,File,Extention -Unique
$Media = [PSCustomObject]@{
    Subject = "Photo or Video"
    Size = ""
    Number = $MediaArray.count
}
$Summary += $media

$exeArray = $exeArray | Sort-Object -Property Folder,File,Extention -Unique
$AppData = [PSCustomObject]@{
    Subject = "Potential Application Data"
    Size = ""
    Number = $exeArray.count
}
$Summary += $Appdata

$filesExt = $filesExt | Sort-Object -Property Folder,File,Extention -Unique
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

if ($ScanPath.count -gt 1) {
    $name = "CUSTOM"

    $summary | Export-Excel -WorksheetName "Summary" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $Dirstats | Sort-Object -Property "Folder" | Export-Excel -WorksheetName "All Folders" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $Dirstats | Sort-Object -Property "Folder" | Where-Object { $_.Level -eq 1 }  | Export-Excel -WorksheetName "Level 1" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $Dirstats | Sort-Object -Property "Folder" | Where-Object { $_.Level -eq 2 }  | Export-Excel -WorksheetName "Level 2" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $Dirstats | Sort-Object -Property "Folder" | Where-Object { $_.Level -eq 3 }  | Export-Excel -WorksheetName "Level 3" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $Dirstats | Sort-Object -Property "Folder" | Where-Object { $_.Level -eq 4 }  | Export-Excel -WorksheetName "Level 4" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $ErrorFolders | Sort-Object -Property "Folder" | Export-Excel -WorksheetName "Errors" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $FilesExt | Sort-Object -Property "Folder" | Export-Excel -Worksheetname "PST - OneNote" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $ExeArray | Sort-Object -Property "Folder" | Export-Excel -Worksheetname "Executables" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $mediaArray | Sort-Object -Property "Folder" | Export-Excel -Worksheetname "MediaFiles" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow

}
else {
    if ($scanpath.EndsWith("\")) {
        $scanpath = $scanpath.Substring(0, $scanpath.Length - 1)
    }
    $path = $Scanpath
    $number = ($Path.split("\")).count - 1
    $name = $path.split("\")[$number]


    $summary | Export-Excel -WorksheetName "Summary" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $Dirstats | Sort-Object -Property "Folder" | Export-Excel -WorksheetName "All Folders" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $Dirstats | Sort-Object -Property "Folder" | Where-Object { $_.Level -eq 1 }  | Export-Excel -WorksheetName "Level 1" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $Dirstats | Sort-Object -Property "Folder" | Where-Object { $_.Level -eq 2 }  | Export-Excel -WorksheetName "Level 2" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $Dirstats | Sort-Object -Property "Folder" | Where-Object { $_.Level -eq 3 }  | Export-Excel -WorksheetName "Level 3" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $Dirstats | Sort-Object -Property "Folder" | Where-Object { $_.Level -eq 4 }  | Export-Excel -WorksheetName "Level 4" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $ErrorFolders | Sort-Object -Property "Folder" | Export-Excel -WorksheetName "Errors" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $FilesExt | Sort-Object -Property "Folder" | Export-Excel -Worksheetname "PST - OneNote" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $ExeArray | Sort-Object -Property "Folder" | Export-Excel -Worksheetname "Executables" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow
    $mediaArray | Sort-Object -Property "Folder" | Export-Excel -Worksheetname "MediaFiles" -path "$ExportPath\$NAME-Export-FileAnalyse-$(Get-Date -Format "ddMMyyyy").xlsx" -FreezeTopRow -AutoSize -AutoFilter -BoldTopRow


}
Write-Host "-------------------------------------"
Write-host "Exported of fileshare Assessment"
Write-Host "-------------------------------------"