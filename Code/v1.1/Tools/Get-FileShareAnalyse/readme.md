# File Analysis and Export Tool
[![powershell][powershell]][powershell-url] <br>

## Introduction
This PowerShell script is designed to generate an Excel report that analyzes the content of a file share up to four levels deep. The script scans the specified file share path(s), collects information about folders and files, and generates various statistics, including folder size, item count, and identification of specific file types. The report is then exported to an Excel file for further analysis and review.

## Script Parameters
### -ScanPath (Required)
Description: An array of UNC paths representing the file share(s) to scan. Multiple paths can be provided for scanning multiple file shares simultaneously.
Data Type: Array (String[])
Example Usage:

```powershell
$ScanPath = "\\server\share1", "\\server\share2"
```

### -ExportPath (Required)
Description: The path to export the generated Excel report.
Data Type: String
Example Usage:

```powershell
$ExportPath = "C:\Reports\"
```

### -Threads
Description: The number of threads to use for parallel processing while scanning the file share. If not provided, the default value of 10 will be used.
Data Type: Integer
Example Usage:

```powershell
$threads = 15
```

## Required Module
This script uses the ImportExcel module to handle Excel export functionalities. If the module is not installed, the script will attempt to install it using Install-Module.

## Functions
The script defines three functions to collect and process information about folders, files, and errors encountered during the scan:

* Add-Item: Creates a custom object to store folder statistics, such as size, item count, and creation time.
* Add-Error: Creates a custom object to store information about folders that encountered errors during the scan.
* Add-Files: Creates a custom object to store information about individual files, including size and extension.

## Main Processing
The script utilizes parallel processing to scan the specified file share(s) and collect relevant data. It iterates through the provided ScanPath array, performs the following tasks for each file share:

1 Checks if the specified file share path exists. If not, an error is displayed, and the script proceeds to the next file share.
2 Scans the file share up to four levels deep, collecting folder statistics, file information, and potential errors.
3 Segregates files into different categories based on file extensions (e.g., executable files, media files, PST files, etc.).

## Output
After scanning all specified file shares and collecting the data, the script generates several worksheets in the Excel report:

* "Summary": Provides an overview of total sizes, item counts, and identified errors.
* "All Folders": Lists folder statistics for all levels.
* "Level 1" to "Level 4": Lists folder statistics for each level of the file share hierarchy.
* "Errors": Lists folders that encountered errors during the scan, including access denied and possible filename length issues.
* "PST - OneNote": Lists information about PST and OneNote files.
* "Executables": Lists information about executable files (e.g., .exe, .bat, .vbs).
* "MediaFiles": Lists information about photo and video files (e.g., .jpg, .jpeg, .bmp, .png, .gif, .mp4, .mov, .avi, .flv).

If multiple file shares were scanned, additional worksheets will be created for each individual file share under the appropriate headings.

# Example Usage
```powershell
$ScanPath = "\\server1\share", "\\server2\share"
$ExportPath = "C:\Reports\"
$threads = 15

.\Get-FileShareAnalyse.ps1 -ScanPath $ScanPath -ExportPath $ExportPath -Threads $threads
```
[powershell]: https://img.shields.io/badge/script-Powershell-blue?style=for-the-badge&logo=PowerShell&logoColor=4FC08D
[powershell-url]: https://learn.microsoft.com/en-gb/powershell/scripting/overview?view=powershell-7.3