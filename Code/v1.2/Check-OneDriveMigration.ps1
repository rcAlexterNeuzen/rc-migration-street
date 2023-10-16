param (
    [string]$CreatedAfter,
    [string]$ModifiedAfter
)

#region default for each script
Try {
    Import-Module ".\modules\rc-migration-module.psm1" -DisableNameChecking -ErrorAction Stop
}
catch {
    $fault = $_
    Write-Host "ERROR importing the default MIGRATION module needed for logging and connections" -ForegroundColor Red
    write-Host "$($Fault.ErrorDetails.Message)"
    break
}
if (!($isMacOs)) {
    $Switch = "\"
}
else {
    $Switch = "/"
}

# version check
If ($PSVersionTable.PSVersion.Major -ne 5) {
    Write-Host "[ERROR] - You need to run this script in Windows Powershell version 5.1" -ForegroundColor Red
    break
}


# check if SPMT is installed
try {
    $check = Get-Item -Path "$env:UserProfile\Documents\WindowsPowerShell\Modules\Microsoft.SharePoint.MigrationTool.PowerShell" -ErrorAction Stop
}
catch {
    $CheckError = $_
    Write-Host "ERROR - Cannot find SharepointMigrationTool in path: $env:UserProfile\Documents\WindowsPowerShell\Modules\Microsoft.SharePoint.MigrationTool.PowerShell " -ForegroundColor Red
    Write-Host "ERROR - $($CheckError.Exception.Message)" -ForegroundColor Red
    Break
}

$Global:file = "." + $switch + "log" + $switch + "$(Get-Date -format "ddMMyyy-HHmm")-OneDriveMigration.txt"

$UploadFolder = "." + $switch + "MigrationResults" + $Switch + "OneDrive" + $switch + "$(Get-Date -Format "dd-MM-yyyy")"
if (Test-Path -path $UploadFolder) {
    Log-Message -file $file -Status Done -Message "Upload folder is found: $uploadfolder"
}
else {
    New-Item -Path $UploadFolder -ItemType Directory | Out-Null
    Log-Message -file $file -Status warning -Message "Upload folder is created: $uploadFolder"
}


# start provisioning teams
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "Start script for OneDrive migration $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "

#region checks
## checking if ImportExcel is present
try {
    Import-Module ImportExcel -Force -ErrorAction stop
    $installExcel = $false
}
catch {
    $installExcel = $true
}

## import custom modules

Try {
    $modules = Get-ChildItem -Path $("." + $Switch + "modules") -file
    $modules = $modules | Where-Object { $_.Extension -eq ".psm1" }
    Log-Message -File $file -Status DONE -Message "Getting modules from $("." + $Switch + "modules")"
}
Catch {
    $moduleError = $_
    Log-Message -File $file -Status error -Message "$($moduleError.Exception.Message)"
    break
}

ForEach ($Module in $modules) {
    try {
        Import-Module $Module.FullName -Force -DisableNameChecking -ErrorAction Stop
        Log-Message -File $file -status DONE -Message "- Imported $($module.FullName)"
        
    }
    catch {
        $moduleError = $_
        Log-Message -file $file -Status ERROR -Message "- Error importing $($module.FullName) : $($moduleError.Exception.Message)"
        break
    }
}

Try {
    Import-Module "$env:UserProfile\Documents\WindowsPowerShell\Modules\Microsoft.SharePoint.MigrationTool.PowerShell\microsoft.sharepoint.migrationtool.powershell.dll"
    Log-Message -file $file -Status DONE -Message "- Imported microsoft.sharepoint.migrationtool.powershell.dll "

}
Catch {
    $importError = $_
    Log-Message -file $file -Status ERROR -Message "$($importError.Exception.Message) "
    break
}


# getting datafiles
Try {
    $datafiles = Get-ChildItem -Path $("." + $Switch + "import-data") -file -ErrorAction Stop
    if (!($datafiles)) {
        Log-Message -file $file -Status ERROR -Message "There are no PowerShell Data files found in $("." + $Switch + "import-data") "
        break
    }
    $datafiles = $datafiles | Where-Object { $_.Extension -eq ".psd1" }
    Log-Message -file $file -Status WARNING -Message "Getting data files to import from $("." + $Switch + "import-data") "
}
Catch {
    $dataError = $_
    Log-Message -file $file -Status ERROR -Message "$($dataError.Exception.Message) "
    break
}

# importing customer credentials for sharepoint

try {
    $Activity = "Importing Customer credentials for sharepoint"
    Write-Progress -Activity $Activity 
    $Customer = Import-PowerShellDataFile -Path $("." + $Switch + "import-data" + $switch + "sharepoint-information.psd1") -ErrorAction Stop
    Log-Message -file $file -status DONE -Message "Customer sharepoint credentials are imported"
    Write-Progress -Activity $Activity -Completed
}
catch {
    $connectionError = $_
    Log-Message -file $file -Status Error -Message "$($connectionError.Exception.Message)"
    Write-Progress -Activity $Activity -Completed
    break
}


try {
    $Activity = "Importing Customer Teams information"
    Write-Progress -Activity $Activity 
    $Sharepoint = Import-PowerShellDataFile -Path $("." + $Switch + "import-data" + $switch + "teams-and-spo-information.psd1") -ErrorAction Stop
    Log-Message -file $file -status DONE -Message "Customer Teams information are imported"
    Write-Progress -Activity $Activity -Completed
}
catch {
    $connectionError = $_
    Log-Message -file $file -Status Error -Message "$($connectionError.Exception.Message)"
    Write-Progress -Activity $Activity -Completed
    break
}


## if needed installing excel module
if ($installExcel) {
    Install-Requirements -module "ImportExcel"
    Log-Message -file $file -status DONE -Message "Module ImportExcel is installed"
}

## getting security info
try {
    $connectionDetails = Import-clixml -Path $("." + $Switch + "security" + $switch + "connectionDetails.xml") -ErrorAction Stop
    $Global:Appid = $connectionDetails.appid
    $Global:TenantName = $connectionDetails.TenantName
    $Global:Thumbprint = $ConnectionDetails.ThumbPrint
    $Global:ClientSecret = $ConnectionDetails.ClientSecret
    $Global:TenantId = $connectionDetails.TenantId
    Log-Message -file $file -Status Done -Message "Connection details are imported"
}
catch {
    $connectionError = $_
    Log-Message -file $file -Status Error -Message "$($connectionError.Exception.Message)"
    break
}

## getting token

try {
    ## getting token 
    $Activity = "Getting token from Graph API"
    Write-Progress -Activity $Activity
    if ($PSVersionTable.PSVersion.Major -ne 5) {
        $token = Get-TokenForGraphAPIWithCertificate -appid $connectionDetails.appid -tenantname $connectionDetails.tenantname -Thumbprint $connectionDetails.ThumbPrint
    }
    else {
        $token = Get-TokenForGraphAPI -appid $connectionDetails.appid -tenantid $connectionDetails.TenantId -clientsecret $ConnectionDetails.clientsecret
    }
    if ($Token.access_token -ne $null) {
        Log-Message -File $file -Status done -Message "Token for Graph API is present"
        Write-Progress -Activity $Activity -Completed
    }
    else {
        Log-Message -File $file -Status WARNING -Message "No token was retrieved for Graph API. Trying again ..."
        Write-Progress -Activity $Activity -Completed
    }
}
catch {
    #do nothing
    break
}

#endregion checks
#endregion 

## getting customer data from excel sheet
$CustomerFile = "." + $switch + "CustomerData" + $switch + "MigrationStreet.xlsx"
$FolderExport = "." + $switch + "CustomerData" + $Switch + "Migration Assessment"

If (!(Test-Path $FolderExport)) {
    $Create = New-Item -Path $FolderExport -ItemType Directory 
}

$WorksheetName = "OneDrive"
try {
    $Activity = "Importing $Worksheetname from $Customerfile"
    Write-Progress -Activity $Activity
    $ToMigrate = Import-Excel -Path $CustomerFile -WorksheetName $WorksheetName -ErrorAction Stop
    Log-Message -file $file -Status Done -Message "Onedrive sheet is imported: $CustomerFile"
    Write-Progress -Activity $Activity -Completed
}
catch {
    $ProvisionError = $_
    Log-Message -file $file -Status error -Message "$($ProvisionError.Exception.Message)"
    Write-Progress -Activity $Activity -Completed
    break
}
## getting sharepoint information
## getting current Sites
try {
    $Activity = "Getting all sharepoint sites"
    Write-Progress -Activity $Activity
    $apiUri = "https://graph.microsoft.com/v1.0/sites"
    $Sites = RunQueryandEnumerateResults -ApiUri $apiUri

    Log-Message -file $file -Status Done -Message "All current sites from tenant are indexed"
    Write-Progress -Activity $Activity -Completed
}
catch {
    $ProvisionError = $_
    Log-Message -file $file -Status ERROR -Message "$($ProvisionError.Exception.Message)"
    Write-Progress -Activity $Activity -Completed
    break
}

# checking if all paths exists
$FailedFolderCheck = @()
$CorrectFolderCheck = @()
ForEach ($path in $toMigrate) {
    if ($Path.Homedir -eq $null) {
        # skip
    }
    else {
        if (!(Test-Path $Path.HomeDir -ErrorAction SilentlyContinue)) {
            Log-Message -file $file -Status Error -Message "$($Path.HomeDir) does not exist. Skipping ..."

            $item = [PSCustomObject]@{
                Share             = $Path.Homedir
                UserPrincipalName = $Path.UserPrincipalName
                Folder            = $Path.Folder
                Error             = "Cannot find folder"
            }
            $FailedFolderCheck += $item
        }
        else {
            Log-Message -file $file -Status done -Message "$($Path.HomeDir) was found"
            $CorrectFolderCheck += $path
        }
    }
}

if (!($CorrectFolderCheck)) {
    Log-Message -file $file -Status Error -Message "Nothing to migrate"
    break
}
# connecting sharepoint

Try {
    $Activity = "Connecting to sharepoint online"
    Write-Progress -Activity $Activity
    Connect-SPOService -Url $Customer.SPOUrl
    Log-Message -file $file -Status Done -Message "Connected to sharepoint online"
    Write-Progress -Activity $Activity -Completed
}
catch {
    $ProvisionError = $_
    Log-Message -file $file -Status Error -Message "$($ProvisionError.Exception.Message)"
    Write-Progress -Activity $Activity -Completed
    break
}


If ($Customer.weblogin -eq $true) {
    Write-Host "Logging in SPMT with web credentials"
    if ((!($CreatedAfter)) -and (!($ModifiedAfter))) {
        Connect-SPMT -LoginWithWeb $true -Scanonly $true
    }
    else {
        Connect-SPMT -LoginWithWeb $true -createdAfter $CreatedAfter -ModifiedAfter $ModifiedAfter -Scanonly $true
    }
}
else {
    $Global:SPOUrl = $Customer.SPOurl
    $Global:UserName = $Customer.Username
    $Global:PassWord = ConvertTo-SecureString -String $Customer.password -AsPlainText -Force
    $Global:SPOCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Global:UserName, $Global:PassWord
    
    if ((!($CreatedAfter)) -and (!($ModifiedAfter))) {
        Connect-SPMT -Scanonly $true
    }
    else {
        Connect-SPMT -createdAfter $CreatedAfter -ModifiedAfter $ModifiedAfter -Scanonly $true
    }
}



Log-Message -file $file  " "
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "Start OneDrive assessment $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "


ForEach ($migration in $CorrectFolderCheck) {
    try {
        $site = $Sites | Where-Object { $_.weburl -like "*$($migration.UserPrincipalName.replace("@","_").replace(".","_"))*" }

        $OneDriveSiteUrl = $site.WebUrl
        $SiteCollAdmin = $Customer.Username

        try {
            Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SiteCollAdmin -IsSiteCollectionAdmin $True -ErrorAction SilentlyContinue | Out-Null
            Log-Message -file $file -Status added -Message "Added $sitecolladmin to $($Site.WebUrl)"
        }
        catch {
            $setError = $_
            Log-Message -file $file -Status error -Message "Cannot add $sitecolladmin because $($Seterror.Exception.Message)"
        }

        if ($Migration.folder -eq "") {
            $job = Add-SPMTTask -FileShareSource $Migration.HomeDir -TargetSiteUrl $Site.WebUrl -TargetList "Documents"
        }
        else {
            $Job = Add-SPMTTask -FileShareSource $Migration.Homedir -TargetSiteUrl $Site.WebUrl  -TargetList "Documents"  -TargetListRelativePath $Migration.Folder 
        }

        Log-Message -file $file -Status Done -Message "Created task for $($migration.Homedir) to $($Site.WebUrl)"

        # getting task id
        #$migration.MigrationId = ((Get-SPMTMigration).StatusOfTasks | Where-Object { $_.TargetURI -eq $Site.webUrl -and $_.SourceURI -eq $Migration.HomeDir }).TaskId.guid 

        Write-Progress -Activity $Activity -Completed
    }
    catch {
        $ProvisionError = $_
        Log-Message -file $file -Status error -Message "$($Migration.UserPrincipalName) cannot migrate: $($ProvisionError.Exception.Message)"
        Write-Progress -Activity $Activity -Completed
    }
}


# starting migration
Try {
    $Activity = "Starting FileShare migration"
    Write-Progress -Activity $Activity
    Start-SPMTMigration -NoShow
    Log-Message -file $file -Status Done -Message "Migration of OneDrive is started"
    Write-Progress -Activity $Activity -Completed
}
catch {
    $StartError = $_
    Log-Message -file $file -Status ERROR -Message "$($StartError.Exception.Message)"
    Write-Progress -Activity $Activity -Completed
    break
}

$session = Get-SPMTMigration
Log-Message -File $File -status Waiting -Message "Waiting for job to complete."
while ($session.Status -ne "Finished") {
    Write-Host "[" -NoNewline
    Write-Host "WAITING" -ForegroundColor Yellow -NoNewline
    Write-Host "] - Waiting for job to complete. Refreshing every 60 seconds"
    Start-Sleep -Seconds 60
}

ForEach ($migration in $CorrectFolderCheck) {
    if ($migration.homedir -eq $null) {
        # skip
    }
    else {
        try {
            $site = $Sites | Where-Object { $_.weburl -like "*/$($migration.UserPrincipalName.replace("@","_").replace(".","_"))" }

            $OneDriveSiteUrl = $site.WebUrl
            $SiteCollAdmin = $Customer.Username

            try {
                Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SiteCollAdmin -IsSiteCollectionAdmin $false -ErrorAction SilentlyContinue | out-Null
                Log-Message -file $file -Status DONE -Message "Removed $sitecolladmin from $($Site.WebUrl)"
            }
            catch {
                $setError = $_
                Log-Message -file $file -Status ERROR -Message "Cannot delete $sitecolladmin because $($Seterror.Exception.Message)"
            }
        }
        catch {
            $ProvisionError = $_
            Log-Message -file $file -Status ERROR -Message "$($Migration.UserPrincipalName) cannot be removed: $($ProvisionError.Exception.Message)"
        }
    }
}

#zipping 
$number = $Session.ReportFolderPath.split("$switch").count - 2
$ArchiveName = $Session.ReportFolderPath.split("$switch")[$number] + ".zip"
$folderToZip = $session.ReportFolderPath

Compress-Archive -LiteralPath $folderToZip -DestinationPath $($UploadFolder + $switch + $ArchiveName) -CompressionLevel Fastest

# exporting results
$Export = @()
ForEach ($Task in $Session.StatusOfTasks) {
    $ExportFile = $Task.ReportFolder + $switch + "ItemReport_R1.csv"
    $CSV = Import-CSV $ExportFile
    $Export += $CSV
}
$Folders = $Export | Where-Object { $_.Type -eq "Folder" -and $_.Status -eq "Scan Finished" }
$ErrorFiles = $Export | Where-Object { $_.Status -ne "Skipped" -or $_.Status -ne "Failed" }
$FolderExport = $UploadFolder + $switch + "$(Get-Date -Format "dd-MM-yyyy")-CheckOneDriveMigration.xlsx"
$FailedFolderCheck | Export-Excel -WorksheetName "ERROR" -Path $FolderExport -BoldTopRow -AutoSize -AutoFilter
$errorFiles | Export-Excel -WorksheetName "ERROR" -Path $FolderExport -BoldTopRow -AutoSize -AutoFilter
$Export | Export-Excel -WorksheetName "$(Get-Date -Format "dd-MM-yyyy") - All" -Path $FolderExport -BoldTopRow -AutoSize -AutoFilter
$Folders | Export-Excel -WorksheetName "$(Get-Date -Format "dd-MM-yyyy") - Folders" -Path $FolderExport -BoldTopRow -AutoSize -AutoFilter

Log-Message -file $file -Status Done -Message "Exported job information to $FolderExport"

if ($Customer.Teams) {
    #uploading to teams
    $FilesToUpload = Get-ChildItem -File $UploadFolder
    $apiUri = "https://graph.microsoft.com/v1.0/teams/$($Sharepoint.Ids.TeamsId)/channels/$($Sharepoint.ids.HomeFolderChannelid)/filesFolder"
    $drive = Invoke-Restmethod -method Get -headers @{Authorization = "Bearer $($Token.access_token)" } -Uri $apiUri
    $driveId = $drive.id
    $driveItemId = $drive.parentReference.driveid

    Log-message -file $file -Status INFO -Message "Uploading all logging to Teams"
    ForEach ($upload in $FilesToUpload) {
        if ($Upload.name -like "WF*.zip") {
            $folderlocation = $sharepoint.folders.Homefolderfolder + "SPMT%20Logging/$(Get-Date -Format "dd-MM-yyyy")"
        }
        else {
            $folderlocation = $sharepoint.folders.Homefolderfolder + "Custom%20Logging/$(Get-Date -Format "dd-MM-yyyy")"
        }
        $locationToUpload = "https://graph.microsoft.com/v1.0/drives/$driveItemId/root:/$folderlocation/$($upload.name)" + ':/content'

        try {
            $silent = Invoke-Restmethod -Method PUT -headers @{Authorization = "Bearer $($Token.access_token)" } -uri $locationToUpload -InFile $($upload.fullname) -ContentType 'multipart/form-data'
            Log-Message -File $file -Status DONE -Message "- $($upload.name) is uploaded to teams"
        }
        catch {
            $webError = $_
            Log-message -File $file -status Error -Message "- Error uploading $($upload.name) to teams: $($webError.exception.message)"
        }
    }
}

Log-Message -File $file ""
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "End Onedrive migrations check $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "

if ($Customer.Teams){
    # send mail to teams channel
    $apiUri = "https://graph.microsoft.com/v1.0/Teams/$($Sharepoint.Ids.TeamsId)/channels"
    $channel = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.access_token)" } -Uri $apiUri -Method Get).value | Where-Object { $_.DisplayName -eq $Sharepoint.HomeDir }
    $emailTo = $channel.email

    $message = "<b>The homefolder migration checks are done</b><br>"
    $message += "<br>"
    $message += "The logging can be found in this channel files and folders.<br>"
    $message += "A total of $($CorrectFolderCheck.count) shares are checked to the requested OneDrives<br>"
    $message += "<br>"
    $message += "Kind regards,<br>"
    $message += "The Rapid Circle Migration Street"

    Send-MailToinform -to $emailTo -From $customer.MigrationMail -Subject "$($Customer.CompanyName) - $(Get-Date -Format "dd-MM-yyyy hh:mm") Homefolder migrations checks are done" -Message $message
}