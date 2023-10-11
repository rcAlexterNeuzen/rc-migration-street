# Folder Index Generator
[![powershell][powershell]][powershell-url] <br>

The "Folder Index Generator" PowerShell script is designed to scan a specified network share, gather information about its subfolders, and export the collected data to an Excel spreadsheet. The script utilizes multi-threading for efficient processing, allowing it to handle large directory structures with improved performance.

### Script Parameters
The script requires the following parameters to be provided when executing it:

- -Share (Mandatory): The network share path from which the script will scan the subfolders.
- -outputPath (Mandatory): The directory where the Excel output file will be saved.
- -DepartmentName (Mandatory): The name of the department associated with the generated folder index.
- -threads (Mandatory): The number of threads to be used for multi-threading processing. It determines how many subfolders are processed simultaneously.

### Script Execution
To run the script, open a PowerShell window or script editor, and execute the script with the required parameters. For example:

```powershell
.\Get-FolderIndex.ps1 -Share "\\server\share" -outputPath "C:\Output" -DepartmentName "Finance" -threads 4
```

### Script Functionality
- Scanning Folders: The script starts by scanning all the subfolders within the specified network share ($Share). It uses the Get-ChildItem cmdlet with the -Directory switch to retrieve only directories and not files.

- Multi-Threading: The script utilizes multi-threading for faster processing of subfolders. The number of threads is determined by the $threads parameter, allowing several subfolders to be processed simultaneously.

- Gathering Folder Information: For each subfolder found, the script collects the following information:

- - Folder: The full path of the subfolder.
- - Level: The depth level of the subfolder within the directory structure. The root folder is considered level 1, its direct subfolders are level 2, and so on.
- Progress Monitoring: While processing subfolders, the script displays a progress bar using the Write-Progress cmdlet, indicating which folder is currently being processed.

- Output File Creation: After processing all subfolders, the collected data is sorted alphabetically based on the Folder property. The sorted information is then exported to an Excel spreadsheet using the Export-Excel cmdlet from the PSExcel module. The Excel file is saved in the directory specified by the $outputPath parameter with a filename in the format of ddMMyyyy-DepartmentName-FolderIndex.csv.

### Script Notes
The script utilizes the Sort-Object cmdlet to ensure that the output in the Excel file is sorted alphabetically based on the Folder property.
The Export-Excel cmdlet is used to export data to an Excel file. Make sure you have the PSExcel module installed before running the script. You can install it using the following command:

``` powershell
Install-Module -Name ImportExcel
```

[powershell]: https://img.shields.io/badge/script-Powershell-blue?style=for-the-badge&logo=PowerShell&logoColor=4FC08D
[powershell-url]: https://learn.microsoft.com/en-gb/powershell/scripting/overview?view=powershell-7.3