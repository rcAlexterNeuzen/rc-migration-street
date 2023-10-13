# Rapid Circle Migration Street
[![powershell][powershell]][powershell-url] <br>

Doing file migrations can be a pain in the ass. As Rapid Circle we developed a Migration Street working with ShareGate and simple Excel sheets as input.

<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#Requirements">Requirements</a>
    </li>
    <li>
      <a href="#Install-Requirements.ps1">Install-Requirements.ps1</a>
    </li>
    <li><a href="#Install-app-registration.ps1">Install-App-Registration.ps1</a></li>
    <li><a href="#Start-ProvisioningTeams.ps1">Start-ProvisioningTeams.ps1</a></li>
    <li><a href="#Start-FileShareMigration.ps1">Start-FileShareMigration.ps1</a></li>
    <li><a href="#Start-OneDriveMigration.ps1">Start-OneDriveMigration.ps1</a></li>
  </ol>
</details>


## Requirements
For the fileshare assessments is PowerShell 7.x required which can be downloaded here: https://github.com/PowerShell/PowerShell/releases
The street need some requirements also and they can be installed running 

```powershell
.\Install-requirements.ps1
```
Also is an dedicated Migration Account needed. This account needs to have an M365 license (including mailbox) for the migrations and notifications.


# Install-Requirements.ps1
1. Logging function (Log-Message)<br>
The script begins with a custom Log-Message function, which appends log entries with timestamps to a specified log file.
2. Environment Check<br>
The script checks whether it's running on macOS ($isMacOs). If not, it assumes a Windows environment.
3. Powershell version check<br>
It verifies that the script is running in Windows PowerShell version 5.1, displaying an error message if not.
4. Elevation check<br>
The script checks if it's running with administrator privileges and displays an error message if not.
5. Module import<br>
Custom PowerShell modules are imported from required-modules.psd1 and rc-required-modules.psm1.
6. Install-requirements <br>
The main part of the script calls the Install-requirements function, which installs the required components based on the imported modules
7. Download folder check<br>
The script checks if the download folder exists. If not, it creates it
8. Sharepoint Migration Tool check<br>
The script checks for the presence of the SharePoint Migration Tool (spmtsetup.exe) and downloads it if necessary.
9. Visual C++ Redistributable check<br>
It checks for the presence of Microsoft Visual C++ Redistributable (vcredist_x64.exe) and downloads it if necessary.
10. .Net Framework Version Check<br>
The script checks if the required .NET Framework version (4.6 or higher) is installed. If not, it downloads and installs it.
11. Reporting<br>
The script reports the installation status of each component, including success and any errors encountered.
12. Logging<br>
Log entries are added for each step of the installation process

# Install-App-Registration.ps1
This PowerShell script is designed to create an application registration for Teams provisioning. It involves importing required modules, creating the app registration, setting permissions, obtaining admin consent, and generating secrets.

## Script Structure

The script is divided into two main sections:

1. **App Registration Setup**
   - Importing necessary modules.
   - Creating an Azure AD application registration.
   - Adding required permissions.
   - Waiting for the app to be available.
   - Obtaining admin consent.
   - Creating an application secret.
   - Exporting connection details.

2. **SharePoint Information**
   - Checking for a SharePoint information file.
   - Collecting SharePoint information if not found or requested.

## App Registration Setup

### Module Import
- The script begins by importing required PowerShell modules for Azure and custom modules for migration.

### App Registration Creation
- An Azure AD application registration is created using information imported from a PowerShell data file.

### Adding Permissions
- The script iterates through permissions and adds them to the app registration.

### Waiting for App Availability
- The script waits for the app to be available in Azure.

### Admin Consent
- Admin consent is requested for the app registration.

### Application Secret
- An application secret is generated for the app registration and exported.

### Certificate Generation
- A self-signed certificate is created for the app registration and exported as both .pfx and .cer files.

### Exporting Connection Details
- Connection details, including tenant ID, app ID, client secret, and more, are exported to an XML file.

## SharePoint Information

### SharePoint File Handling
- The script checks for an existing SharePoint information file and asks if the user wants to replace it.

### SharePoint Information Collection
- If needed, the script collects SharePoint admin credentials, company name, Teams configuration, and other details.

### Exporting SharePoint Information
- The collected SharePoint information is exported to a PowerShell data file.

## Summary

This PowerShell script automates the creation of an Azure AD app registration for Teams provisioning and collects SharePoint information for migration purposes. Ensure that you have the required modules and permissions before running the script. Follow the prompts to provide necessary information when prompted.

For more details on each section and specific commands, please refer to the script's source code.


# Start-ProvisioningTeams.ps1
The script to create Microsoft Teams and Channels including adding Owners and Members.
The script is organized into several sections, each performing specific tasks related to the provisioning process.

1. Default Settings: Sets up default settings, file paths, and log messages.
2. Checking Prerequisites: Checks for required modules and data files.
3. Importing Data: Imports customer data from an Excel spreadsheet.
4. Provisioning Teams: Creates Microsoft Teams and assigns owners and members. Creating also a folder to provision the drive and site. After creating it will be deleted.
5. Adding Channels: Adds channels to the created teams.
6. Exporting Updated Data: Updates and exports data back to the Excel spreadsheet.
7. Logging and Error Handling: Logs activities and errors to a log file.


# Start-FileShareMigration.ps1
This PowerShell script is designed for migrating files from a file share to Microsoft Teams using the SharePoint Migration Tool (SPMT). It automates the migration process and provides detailed logging for tracking the migration progress.

## Parameters
The script accepts the following parameters:
- `$CreatedAfter`: Specifies the creation date after which to migrate files (optional).
- `$ModifiedAfter`: Specifies the modification date after which to migrate files (optional).

## Prerequisites
- Ensure that PowerShell version 5.1 is installed.
- The SharePoint Migration Tool (SPMT) must be installed on the system.
- The `ImportExcel` module should be available.
- Custom PowerShell modules are required and should be placed in the 'modules' directory.
- SharePoint credentials and connection details should be available in the 'import-data' directory.
- A customer data file (Excel) containing migration information should be present.

## Execution Flow
1. Imports the required PowerShell modules, both custom and those from SPMT.
2. Reads customer SharePoint credentials and other connection details.
3. Checks if the 'ImportExcel' module is available; if not, it installs the module.
4. Imports customer data and SharePoint information from an Excel file.
5. Checks for the existence of specified paths on the file share.
6. Establishes a connection to SharePoint using provided credentials.
7. Creates migration tasks for file shares to Teams and logs the progress.
8. Initiates the file share migration process.
9. Monitors the migration process, updating the status of tasks.
10. Generates detailed reports and logs the outcome of the migration.
11. Exports reports to an Excel file and zips the migration results.
12. Optionally uploads the logs to Microsoft Teams.

## Logging and Reports
- A log file is created for tracking the entire migration process.
- Detailed reports are generated in Excel format for both successful and failed tasks.
- The final results are exported to Excel files and archived.

## Teams Integration
- The script can upload logs and reports to Microsoft Teams for better visibility and accessibility.

## Retry Mechanism
- In case of failed tasks, the script provides an option to retry the migration of those tasks.

## Additional Notes
- The script is designed for migrating files from a file share to Microsoft Teams using the SharePoint Migration Tool.
- Detailed error handling and logging are implemented for tracking the migration process.
- Custom modules and credentials are expected to be available in specific directories.

## Execution
1. Open PowerShell.
2. Navigate to the directory containing this script.
3. Run the script, providing the optional `CreatedAfter` and `ModifiedAfter` parameters if needed.

Example:
```powershell
.\Start-FileShareMigration.ps1 -CreatedAfter "2023-01-01" -ModifiedAfter "2023-09-30"
```


# Start-OneDriveMigration.ps1
This documentation explains the functionality and usage of the PowerShell script designed for migrating OneDrive data using the SharePoint Migration Tool (SPMT).

## Prerequisites

Before using this script, make sure you meet the following prerequisites:

- Windows PowerShell version 5.1 or later is required.
- The SharePoint Migration Tool (SPMT) should be installed on your system.
- Required PowerShell modules and custom modules should be available.
- Customer credentials for SharePoint should be configured in `sharepoint-information.psd1`.

## Script Parameters

The script accepts the following parameters:

- `-CreatedAfter`: Specifies the date after which content was created for migration.
- `-ModifiedAfter`: Specifies the date after which content was modified for migration.

## Default Module Import

- The script attempts to import the `rc-migration-module.psm1` module, necessary for logging and connections.
- If the module import fails, an error message is displayed.

## Module Imports

- The script checks and imports the necessary modules, including `ImportExcel`, custom modules, and the SharePoint Migration Tool module.
- If any module import fails, appropriate error messages are displayed.

## Data Files

- The script searches for PowerShell data files in the `import-data` directory with the extension `.psd1`.
- It imports customer credentials for SharePoint from `sharepoint-information.psd1`.

## OneDrive Data Retrieval

- The script imports OneDrive data from an Excel sheet, and the data is stored in the `$ToMigrate` variable.

## Checking Folder Paths

- The script checks if all folder paths exist for migration.
- It separates folders with correct paths and those with missing paths.

## Connection to SharePoint Online

- The script connects to SharePoint Online using customer credentials.
- Site collection administrators are assigned to the respective OneDrive sites.

## Adding Migration Tasks

- Migration tasks are created for each folder to be migrated to OneDrive.
- A task is created in SPMT for each folder with its corresponding source and target information.

## Starting Migration

- The script starts the migration using SPMT with the `Start-SPMTMigration` cmdlet.
- Migration progress is monitored while waiting for the completion of the migration tasks.

## Handling Failed Migrations

- If any migration tasks fail, they are restarted and monitored for completion.

## Generating Reports

- Once the migration is completed, a summary report is generated with details of the migration.
- Individual reports are generated for failed files and folders.
- The reports are exported to Excel and CSV files.

## Uploading to Microsoft Teams (Optional)

- If configured, the script uploads migration logs and reports to Microsoft Teams channels.

## Sending Email Notification

- The script sends an email notification containing the migration summary to the specified recipient.

## Cleanup

- Site collection administrators are removed from OneDrive sites.
- The final logs and reports are compressed into a ZIP archive.

## Summary

This script provides a comprehensive solution for migrating OneDrive data using the SharePoint Migration Tool, monitoring the progress, and generating reports. It can be configured to send notifications and upload logs to Microsoft Teams if required.

For any issues, error messages will be logged and displayed for debugging.

Please ensure all prerequisites and configuration files are correctly set up before running the script.

## Execution
1. Open PowerShell.
2. Navigate to the directory containing this script.
3. Run the script, providing the optional `CreatedAfter` and `ModifiedAfter` parameters if needed.

Example:
```powershell
.\Start-OneDriveMigration.ps1 -CreatedAfter "2023-01-01" -ModifiedAfter "2023-09-30"
```










[powershell]: https://img.shields.io/badge/script-Powershell-blue?style=for-the-badge&logo=PowerShell&logoColor=4FC08D
[powershell-url]: https://learn.microsoft.com/en-gb/powershell/scripting/overview?view=powershell-7.3