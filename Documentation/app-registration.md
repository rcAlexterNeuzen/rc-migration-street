# App-Registration.psd1 for creating the App-Registration
This PowerShell Data File contains information related to an Azure Active Directory (Azure AD) application registration for provisioning teams using Microsoft Graph API.

## GraphApi
* ID: 00000003-0000-0000-c000-000000000000
* Description: This is the ID for Microsoft Graph API, which is used to interact with Microsoft 365 services, including Teams.
## AppRegistration
* Name: Rapic Circle Migration - Teams Provisioning
* Description: App registration for provisioning teams.
* SignInAudience: AzureADMyOrg
* web: https://portal.azure.com

## Permissions (perm)
The perm array contains a list of permission IDs (also known as application permissions) required by the registered application to perform specific operations on Microsoft Teams and SharePoint resources.

* Channel.Create (Application)

ID: f3a65bd4-b703-46df-8f7e-0174fea562aa
Description: Allows the application to create new channels within a Microsoft Teams team.
* Channel.ReadBasic.All (Application)

ID: 59a6b24b-4225-4393-8165-ebaec5f55d7a
Description: Allows the application to read basic information about channels in Microsoft Teams.
* ChannelMember.ReadWrite.All (Application)

ID: 35930dcf-aceb-4bd1-b99a-8ffed403c974
Description: Allows the application to read and write information about members in Microsoft Teams channels.
* ChannelSettings.Read.All (Application)

ID: c97b873f-f59f-49aa-8a0e-52b32d762124
Description: Allows the application to read settings of channels in Microsoft Teams.
* Sites.Read.All (Application)

ID: 332a536c-c7ef-4017-ab91-336970924f0d
Description: Allows the application to read information about SharePoint sites in the organization.
* Team.Create (Application)

ID: 23fc2474-f741-46ce-8465-674744c5c361
Description: Allows the application to create new Microsoft Teams.
* Team.ReadBasic.All (Application)

ID: 2280dda6-0bfd-44ee-a2f4-cb867cfc4c1e
Description: Allows the application to read basic information about Microsoft Teams.
* TeamMember.ReadWrite.All (Application)

ID: 0121dc95-1b9f-4aed-8bac-58c5ac466691
Description: Allows the application to read and write information about members in Microsoft Teams.