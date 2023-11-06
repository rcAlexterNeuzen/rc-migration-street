# Rapid Circle Migration Street
[![powershell][powershell]][powershell-url] ![GitHub release (latest by date)](https://img.shields.io/github/v/release/rcalexterneuzen/rc-migration-street?style=for-the-badge) <br>

## Introduction
This documentation explains the functionality and usage of a PowerShell script designed for installing and updating "Rapid Circle Migration Street." This script is responsible for checking for updates and installing the latest version from a GitHub repository. It also handles creating backups and dependencies installation. The script provides status messages throughout its execution to keep users informed.

## Download
The file can be download from github or through this command: 

```powershell
(Invoke-WebRequest -Uri "https://raw.githubusercontent.com/rcAlexterNeuzen/rc-migration-street/main/install-rcmigrationstreet.ps1" -outFile Install-rcMigrationStreet.ps1)
```

## Usage
Before running the script, ensure that PowerShell is opened with administrator privileges. The script can be executed with the following command:

```powershell
.\Install-RCMigrationStreet.ps1 -InstallFolder "C:\Path\To\Installation"
```
Replace "C:\Path\To\Installation" with the desired installation directory.

## Function: Show-Message
The Show-Message function is used to display informative messages during script execution. It can display messages with different statuses, such as INFO, WARNING, ERROR, DONE, etc., in different colors. The function enhances the user experience by providing detailed progress updates.

## Permission Check
The script begins by checking if it has been executed with administrator privileges. If not, it displays an error message and terminates, as administrator privileges are required to perform certain actions.

## Script Workflow
The script follows the following workflow:

* It defines the installation folder, GitHub repository details, and other variables.
* If the $InstallFolder parameter is not provided, it uses a default folder path, which is "C:\RCScripts."
* It checks if the installation folder exists. If not, it attempts to create it and displays relevant status messages.
* It retrieves the latest release tag from the specified GitHub repository. If an error occurs during this step, it displays an error message.
* If a new installation is being performed, it downloads and extracts the latest release, removes unnecessary files, and provides update progress messages.
* If the script is in "in-place" mode, it checks if the installed version matches the latest release. If they match, it displays a "FINISHED" message, indicating that no update is needed.
* If an update is required for an existing installation, it creates a backup of the current version, downloads and extracts the new release, and removes unnecessary files. It also provides progress messages for each step.
* After the update is completed, the script asks the user if they want to start the installation process for Rapid Circle Migration Street. If confirmed, it installs dependencies and other components.

## Requirements
- The script requires Windows PowerShell with administrator privileges to execute successfully.
- It assumes that the required PowerShell script files for installation and app registration are available in the specified installation folder.
- It relies on GitHub to fetch the latest release information and the installation package.

## Summary
This PowerShell script simplifies the process of installing and updating "Rapid Circle Migration Street" from a GitHub repository. It provides informative messages throughout the process to keep the user informed about the progress. Users are expected to execute the script with administrator privileges to ensure smooth operation.



[powershell]: https://img.shields.io/badge/script-Powershell-blue?style=for-the-badge&logo=PowerShell&logoColor=4FC08D
[powershell-url]: https://learn.microsoft.com/en-gb/powershell/scripting/overview?view=powershell-7.3
