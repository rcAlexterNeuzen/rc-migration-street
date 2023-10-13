<#
.SYNOPSIS
	Removes the Site Collection Administrator from the user's OneDrives provided in the csv

.DESCRIPTION
	Removes the Site Collection Administrator from the user's OneDrives provided in the csv. The OneDrive url will be obtained using the upn/email address of the user

.PARAMETER CSVWithUPNs
    Path to the csv with UPN's

.PARAMETER Admin
    Name of the Administrator that should be removed

.EXAMPLE
    C:\PS> Remove-SiteCollectionAdminFromOneDrive.ps1 -CSVWithUPNs "C:\tmp\Homedrives.csv" -Admin "RC-Service@hocg.onmicrosoft.com"

.NOTES
	Authors : Wim Meurer
	From 	: Rapid Circle
	Date    : 13-09-2023
	Version	: 1.0
#>


[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the path to the csv with UPN's")]
    [string]$CSVWithUPNs,
    [Parameter(Mandatory = $true, HelpMessage = "Name of the Administrator that should be removed")]
    [string]$Admin
)

$UPNs = Import-Csv -Path $CSVWithUPNs -Delimiter ";"

Connect-PnPOnline -Url "https://hocg-admin.sharepoint.com" -Interactive

#Get all OneDrive Url's
$Script:OneDriveSites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'"

function Remove-OwnerOneDrive {
    param (
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
		[string]$Email,
        [parameter(Mandatory = $true, ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $false)]
        [string]$SiteCollectionAdministrator)


    $OneDriveSite = $Script:OneDriveSites | Where-Object Owner -eq $Email

    $OneDriveUrl = $OneDriveSite.Url

    Connect-PnPOnline -Url $OneDriveUrl -Interactive
    
    #Remove OneDrive Site collection admin
    Remove-PnPSiteCollectionAdmin -Owners $SiteCollectionAdministrator

    Write-Host "`t Account '$SiteCollectionAdministrator' removed as site collection admin at OneDrive '$OneDriveUrl' from user with UPN '$Email'"
}

$UPNs[0..($UPNs.count - 1)] | ForEach-Object { $_ | Remove-OwnerOneDrive -SiteCollectionAdministrator $Admin}