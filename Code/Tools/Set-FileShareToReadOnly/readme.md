# Set Fileshare to Read Only
[![powershell][powershell]][powershell-url] <br>

The "User Mapping Script for Homedrive Analysis" is a PowerShell script designed to collect information about user shares located within a specified path. It retrieves data from Active Directory to create a user mapping report, including user display names, user principal names (usually email addresses), and the size of their respective homedrive directories. The collected data is then exported to two separate Excel worksheets within a single Excel file for further analysis and reporting.

### Script Parameters
The script expects the following parameters to be provided when executing it:

- -Path (Mandatory): The path to the root directory where user shares (homedrive directories) are located.
- -OutputPath (Mandatory): The directory where the output Excel file will be saved.

## Prerequisites
Before running the script, ensure that the required PowerShell module "rc-required-modules.psm1" is available in the same directory as the script. Additionally, the script calls a function "Install-Requirements" from the module, which installs the "ImportExcel" module. Make sure the script has the necessary permissions to install modules and access Active Directory information. Run PowerShell in an administrative mode and use PowerShell 7.x.

### Script Execution
To execute the script, open a PowerShell window or script editor, and run the script with the required parameters. For example:

```Powershell
.\Get-UserMapping.ps1 -Path "C:\UserShares" -outputPath "C:\Output"
```

### Script Functionality
1 Module Import: The script starts by importing the custom PowerShell module "rc-required-modules.psm1," which contains the "Install-Requirements" function.

2 Module Installation: The script calls the "Install-Requirements" function to ensure the required PowerShell module "ImportExcel" is installed.

3 Retrieving User Shares: The script gets a list of all users from Active Directory using the Get-ADUser cmdlet with no filter. This enables it to later match user shares with user information.

4 Processing User Shares: The script iterates through each subfolder found in the $Path directory (user shares) using Get-ChildItem with the -Directory switch.

5 User Mapping Export: For each user share, the script attempts to find the corresponding user in Active Directory based on the folder's name (assuming it matches the SamAccountName of the user). If a matching user is found, it creates an export item containing the following properties:

- -DisplayName: The user's display name.
- -userPrincipalName: The user's principal name, typically representing their email address.
- -FullPath: The full path of the user share directory.

6 User Share Size Calculation: Additionally, the script calculates the size of each user share (homedrive). It creates a size item with the following properties:

- -userPrincipalName: The user's principal name (email address).
- -FullPath: The full path of the user share directory.
- -Size: The size of the user share in megabytes, calculated by summing the sizes of all files and folders within the user share.

7 Excel Export: After processing all user shares, the script generates the filename for the output Excel file by appending the current date in "ddMMyyyy" format to the specified $outputPath.

8 Export to Excel Worksheets: The script exports the collected user mapping data to an Excel worksheet named "Homedrive Mapping" and the user share sizes to an Excel worksheet named "Size" within the same Excel file. The data is sorted alphabetically based on the DisplayName for the "Homedrive Mapping" worksheet and userPrincipalName for the "Size" worksheet. Additionally, the script applies formatting options such as freezing the top row, adding an autofilter, and making the top row bold for both worksheets.

[powershell]: https://img.shields.io/badge/script-Powershell-blue?style=for-the-badge&logo=PowerShell&logoColor=4FC08D
[powershell-url]: https://learn.microsoft.com/en-gb/powershell/scripting/overview?view=powershell-7.3