# Module to import
[![powershell][powershell]][powershell-url] <br>

For the migration street there are some custom module files created with some PowerShell functions

<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#rc-azuread-module.psm1">rc-azuread-module.psm1</a>
    </li>
    <li>
      <a href="#rc-graph-api-module.psm1">rc-graph-api-module.psm1</a>
    </li>
    <li><a href="#rc-migration-module.psm1">rc-migration-module.psm1</a></li>
    <li><a href="#rc-required-modules.psm1">rc-required-modules.psm1</a></li>
    <li><a href="#rc-sharepoint-module.psm1">rc-sharepoint-module.psm1</a></li>
    <li><a href="#rc-teams-module.psm1">rc-teams-module.psm1</a></li>
  </ol>
</details>

# RC-AZUREAD-MODULE.PSM1
work in progress

# RC-GRAPH-API-MODULE.PSM1
## Get-TokenForGraphAPI
This function is used to obtain an access token for the Microsoft Graph API using client credentials.

### Parameters
- appid (Mandatory): The application ID (client ID) for your registered Azure AD application.
- clientsecret (Mandatory): The client secret (application password) for your Azure AD application.
- tenantid (Mandatory): The Azure AD tenant ID.

### Return Value
The function returns an access token for the Microsoft Graph API.

## Get-TokenForGraphAPIWithCertificate
This function is used to obtain an access token for the Microsoft Graph API using a certificate.

### Parameters
- appid (Mandatory): The application ID (client ID) for your registered Azure AD application.
- tenantname (Mandatory): The name of the Azure AD tenant.
- Thumbprint (Mandatory): The thumbprint of the certificate.

### Return Value
The function returns an access token for the Microsoft Graph API.

## RunQueryandEnumerateResults
This function is used to run a query against the Microsoft Graph API and enumerate the results.

### Parameters
- apiUri (Mandatory): The URI of the API endpoint you want to query.

### Return Value
The function returns an array of results from the Microsoft Graph API.

### Functionality
The RunQueryandEnumerateResults function first attempts to retrieve an access token using the Get-TokenForGraphAPIWithCertificate function.
If the token has expired or is not yet valid, it will attempt to obtain a new token using the Get-TokenForGraphAPI function.
Once a valid access token is obtained, the function sends an HTTP GET request to the specified apiUri with the access token in the request headers.
It then handles pagination, collecting and concatenating results from multiple pages if available.

### Additional Notes
This module assumes that you have the necessary security details, such as the certificate or client secret, stored in appropriate files or variables.
It handles token expiration and renewal automatically, ensuring that the API query is executed with a valid access token.

# RC-MIGRATION-MODULE.PSM1
## Send-MailToInform
This function sends an email notification to inform recipients about a specific event or message.

### Parameters
- to (Mandatory): The recipient's email address.
- from (Mandatory): The sender's email address.
- Subject (Mandatory): The email subject.
- Message (Mandatory): The email content message.

### Return Value
This function returns a message indicating the status of the email sending process.

## Log-Message
This function logs messages to the console and a specified log file with different statuses such as "INFO," "WARNING," "ERROR," and more.

### Parameters
- Message: The message to log.
- file: The log file where the message is logged.
- Status: The status of the log message (e.g., "INFO," "ERROR," "DONE").

## Connect-SPMT
This function is used for connecting to SharePoint Migration Tool (SPMT) for migration tasks. It supports various migration configurations.

### Parameters
- ScanOnly: A boolean indicating whether the migration is for scanning only.
- CreatedAfter: The date from which to migrate files created.
- ModifiedAfter: The date from which to migrate files modified.
- LoginWithWeb: A boolean indicating whether to log in using web credentials.

## Get-RandomCharacters
This function generates a random string of characters based on the provided length and character set.

Parameters
- length: The desired length of the random string.
- characters: The set of characters to choose from for generating the random string.

### Return Value
The function returns a random string of characters.

## Scramble-String
This function scrambles a given string by rearranging its characters randomly.

### Parameters
- inputString: The input string to be scrambled.

### Return Value
The function returns the input string with its characters scrambled.

## Usage
These PowerShell functions can be used to send emails, log messages, connect to SharePoint Migration Tool, generate random characters, and scramble strings. They are designed to streamline SharePoint migration and related tasks.

Make sure to refer to individual function documentation for specific usage instructions and parameter details.


# RC-REQUIRED-MODULES.PSM1
## Install-Requirements
This function checks for and manages the installation of required PowerShell modules based on a list of module names. It verifies the installed module's version, prompts to install missing modules, and upgrades existing modules if newer versions are available.

### Parameters
- modules (Mandatory): An array of module names to be checked and installed.
- file (Mandatory): The log file to record installation and update actions.

## ConvertTo-Psd
This function serializes input objects into PowerShell Data (PSD) format, providing a textual representation of the data in a PSD string.

### Parameters
- InputObject: The object or data to be serialized.
- Depth: An optional parameter that limits the depth of serialization.
- Indent: An optional parameter to specify the indentation character(s).

### Return Value
The function returns a PSD formatted string.

## Convert-Indent
This function generates indentation based on the provided value for indenting PSD output.

### Parameters
- Indent: The desired indentation value, specifying the number of spaces or tabs for each level of indentation.

### Return Value
The function returns the indentation characters for PSD formatting.

## Write-Psd
This function writes serialized data to the PSD format, providing textual representation suitable for creating PSD strings. It supports a variety of data types and handles complex objects, collections, and custom objects.

### Parameters
- Object: The data to be serialized.
- Depth: An optional parameter that limits the depth of serialization.
- NoIndent: An optional switch to disable indentation.

### Return Value
The function serializes and writes the data into the PSD format.

## Usage
These PowerShell functions can be used to install and manage required modules, convert data to the PSD format, and serialize complex objects for easy representation in PowerShell scripts.

For specific usage details, please refer to the individual function documentation provided above.

# RC-SHAREPOINT-MODULE.PSM1
work in progress

# RC-TEAMS-MODULE.PSM1
## Get-TeamsChannels
This function retrieves a list of channels for a specified Microsoft Teams group using the Microsoft Graph API.

### Parameters:
- $teamsid (Mandatory): The ID of the Microsoft Teams group for which you want to retrieve channels.

## Create-Teams
This function creates a new Microsoft Teams group using the Microsoft Graph API. It allows you to specify the group's name, owner, and description.

### Parameters:
- $Teamsname (Mandatory): The name of the new Microsoft Teams group.
- $Owner (Mandatory): The user principal name (UPN) of the owner of the new Teams group.
- $Description (Mandatory): The description of the new Teams group.

## Create-Channel
This function creates a new channel within a Microsoft Teams group using the Microsoft Graph API. It enables you to specify the channel's name, description, owner, and membership type (public or private).

### Parameters:
- $TeamsId (Mandatory): The ID of the Microsoft Teams group in which you want to create a channel.
- $ChannelName (Mandatory): The name of the new channel.
- $Description (Mandatory): The description of the new channel.
- $owner (Mandatory): The UPN of the owner of the new channel.
- $Type (Mandatory): The membership type of the channel (public or private).

## Add-MemberToTeams
This function adds a member to a Microsoft Teams group using the Microsoft Graph API.

### Parameters:
- $TeamsId (Mandatory): The ID of the Microsoft Teams group to which you want to add a member.
- $Upn (Mandatory): The UPN of the user you want to add as a member.

## Add-MemberToTeamsChannel
This function adds one or more members to a channel within a Microsoft Teams group using the Microsoft Graph API.

### Parameters:
- $TeamsId (Mandatory): The ID of the Microsoft Teams group.
- $ChannelId (Mandatory): The ID of the channel to which you want to add members.
- $Upn (Mandatory): An array of UPNs for users you want to add to the channel.

## Add-OwnerToTeams
This function adds an owner to a Microsoft Teams group using the Microsoft Graph API.

### Parameters:
- $TeamsName (Mandatory): The name of the Microsoft Teams group.
- $Upn (Mandatory): The UPN of the user you want to add as an owner.
- $TeamsId (Mandatory): The ID of the Microsoft Teams group.

## Add-OwnerToTeamsChannel
This function adds one or more owners to a channel within a Microsoft Teams group using the Microsoft Graph API.

### Parameters:
- $TeamsId (Mandatory): The ID of the Microsoft Teams group.
- $ChannelId (Mandatory): The ID of the channel.
- $Upn (Mandatory): An array of UPNs for users you want to add as owners to the channel.

Please note that these functions have dependencies on external configurations and API calls. Ensure that you have the necessary access permissions and configurations in place for these functions to work correctly.