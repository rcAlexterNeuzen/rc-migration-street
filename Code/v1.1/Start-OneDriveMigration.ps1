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
        $token = Get-TokenForGraphAPIWithCertificate -appid $appid -tenantname $tenantname -Thumbprint $ThumbPrint
    }
    else {
        $token = Get-TokenForGraphAPI -appid $appid -tenantid $TenantId -clientsecret $clientsecret
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
$CSV = "." + $switch + "CustomerData" + $switch + "OneDriveJobInformation.csv"
$WorksheetName = "OneDrive"

try {
    $Activity = "Importing $CustomerFile with worksheetname $Worksheetname"
    Write-Progress -Activity $Activity
    $ToMigrate = Import-Excel -Path $CustomerFile -WorksheetName $WorksheetName -ErrorAction Stop
    Log-Message -file $file -Status done -Message "OneDrive sheet to migrate is imported"
    Write-Progress -Activity $Activity -Completed
}
catch {
    $ProvisionError = $_
    Log-Message -file $file -Status Error -Message "$($ProvisionError.Exception.Message)"
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
            Write-Host "[" -NoNewline
            Write-Host "DONE" -ForegroundColor green -NoNewline
            Write-Host "] - $($Path.Homedir) was found"
            $CorrectFolderCheck += $path
        }
    }
}

if (!($CorrectFolderCheck)) {
    Log-Message -file $file -Status Error -Message "Nothing to migrate"
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

}
catch {
    $ProvisionError = $_
    Log-Message -file $file -Status Error -Message "$($ProvisionError.Exception.Message)"
    Write-Progress -Activity $Activity -Completed
    break
}


# connecting sharepoint

$ToMigrate | Add-Member -MemberType NoteProperty -Name "MigrationId" -Value "" -force
$ToMigrate | Add-Member -MemberType NoteProperty -Name "ReportFolder" -Value "" -force
$ToMigrate | Add-Member -MemberType NoteProperty -Name "Status" -Value "" -force
$ToMigrate | Add-Member -MemberType NoteProperty -Name "ScannedTotalFiles" -Value "" -force
$ToMigrate | Add-Member -MemberType NoteProperty -Name "ScannedBadFiles" -Value "" -force
$toMigrate | Add-Member -MemberType NoteProperty -Name "FilesToBeMigrated" -Value "" -Force
$ToMigrate | Add-Member -MemberType NoteProperty -Name "MigratedFiles" -Value "" -force
$ToMigrate | Add-Member -MemberType NoteProperty -Name "MigratingProgressPercentage" -value "" -Force
$ToMigrate | Add-Member -MemberType NoteProperty -Name "FailedFiles" -Value "" -force
$ToMigrate | Add-Member -MemberType NoteProperty -Name "ErrorMsg" -Value "" -force
$ToMigrate | Add-Member -MemberType NoteProperty -Name "Date" -Value "" -Force

$sSPOAdminCenterUrl = $Customer.SPOurl
$username = $Customer.Username
$Password = $Customer.password | ConvertTo-SecureString -AsPlainText -Force
$cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $userName, $password


Try {
    $Activity = "Connecting to sharepoint online"
    Write-Progress -Activity $Activity
    Connect-SPOService -Url $sSPOAdminCenterUrl
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
        Connect-SPMT -LoginWithWeb $true 
    }
    else {
        Connect-SPMT -LoginWithWeb $true -createdAfter $CreatedAfter -ModifiedAfter $ModifiedAfter
    }
}
else {
    $Global:SPOUrl = $Customer.SPOurl
    $Global:UserName = $Customer.Username
    $Global:PassWord = ConvertTo-SecureString -String $Customer.password -AsPlainText -Force
    $Global:SPOCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Global:UserName, $Global:PassWord

    if ((!($CreatedAfter)) -and (!($ModifiedAfter))) {
        Connect-SPMT
    }
    else {
        Connect-SPMT -createdAfter $CreatedAfter -ModifiedAfter $ModifiedAfter
    }
}

Log-Message -file $file  " "
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "Start Onedrive migrations $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "

$Summary = @()
# creating task
ForEach ($migration in $CorrectFolderCheck) {
    if ($Migration.homedir -eq $null) {
        #skip
    }
    else {
        try {
            $site = $Sites | Where-Object { $_.weburl -like "*$($migration.UserPrincipalName.replace("@","_").replace(".","_"))*" }

            $OneDriveSiteUrl = $site.WebUrl
            $SiteCollAdmin = $Customer.Username

            try {
                Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SiteCollAdmin -IsSiteCollectionAdmin $True -ErrorAction SilentlyContinue | out-Null
                Log-Message -file $file -Status ADDED -Message "Added $sitecolladmin to $($Site.WebUrl)"
            }
            catch {
                $setError = $_
                Log-Message -file $file -Status ERROR -Message "Cannot add $sitecolladmin because $($Seterror.Exception.Message)"
            }

            if ($Migration.folder -eq "") {
                $job = Add-SPMTTask -FileShareSource $Migration.HomeDir -TargetSiteUrl $Site.WebUrl -TargetList "Documents"
            }
            else {
                $Job = Add-SPMTTask -FileShareSource $Migration.Homedir -TargetSiteUrl $Site.WebUrl  -TargetList "Documents"  -TargetListRelativePath $Migration.Folder 
            }

            Log-Message -file $file -Status done -Message "Created task for $($migration.Homedir) to $($Site.WebUrl)"

            # getting task id
            $migration.MigrationId = ((Get-SPMTMigration).StatusOfTasks | Where-Object { $_.TargetURI -eq $Site.webUrl }).TaskId

            Write-Progress -Activity $Activity -Completed
        }
        catch {
            $ProvisionError = $_
            Log-Message -file $file -Status ERROR -Message "$($Migration.UserPrincipalName) cannot migrate: $($ProvisionError.Exception.Message)"
            Write-Progress -Activity $Activity -Completed
        }
    }
}


# starting migration
Try {
    $Activity = "Starting Onedrive migration"
    Write-Progress -Activity $Activity
    Start-SPMTMigration -NoShow
    Log-Message -file $file -Status done -Message "Migration of OneDrives is started"
    Write-Progress -Activity $Activity -Completed
}
catch {
    $StartError = $_
    Log-Message -file $file -Status ERROR -Message "$($StartError.Exception.Message)"
    Write-Progress -Activity $Activity -Completed
    break
}

# monitoring
$session = Get-SPMTMigration
Log-Message -File $File -Status WAITING -Message "Waiting for job to complete. Check $CSV for current status"
while ($session.Status -ne "Finished") {
    Write-Host "[" -NoNewline
    Write-Host "WAITING" -ForegroundColor Yellow -NoNewline
    Write-Host "] - $(Get-Date -Format "dd-MM-yyyy HH:mm") Waiting for job to complete. Check $CSV for current status. Refreshing every 60 seconds"
    ForEach ($Task in $Session.StatusOfTasks) {
        $Report = $ToMigrate | Where-Object { $_.MigrationId -eq $Task.TaskId }
        $Report.ScannedTotalFiles = $Task.NumScannedTotalFiles
        $Report.ScannedBadFiles = $task.NumScannedBadFiles
        $report.FilesToBeMigrated = $task.NumFileWillBeMigrated
        $report.MigratedFiles = $task.NumActuallyMigratedFiles
        $report.FailedFiles = $task.NumFailedFiles
        $report.MigratingProgressPercentage = $task.MigratingProgressPercentage
        $report.Errormsg = $task.Errormsg
        $report.Status = $task.Status
        $Report.Date = (Get-Date -Format "dd-MM-yyyy")
    }
    $toMigrate | Export-CSV -Path $csv -encoding UTF8 
    Start-Sleep -seconds 60
}

Log-Message -File $File -Status FINISHED -Message "Migration session is completed"

ForEach ($Task in $Session.StatusOfTasks) {
    $Report = $ToMigrate | Where-Object { $_.Homedir -eq $Task.SourceURI }
    $Report.ReportFolder = $task.ReportFolder
    $Report.ScannedTotalFiles = $Task.NumScannedTotalFiles
    $Report.ScannedBadFiles = $task.NumScannedBadFiles
    $report.FilesToBeMigrated = $task.NumFileWillBeMigrated
    $report.MigratedFiles = $task.NumActuallyMigratedFiles
    $report.FailedFiles = $task.NumFailedFiles
    $report.MigratingProgressPercentage = $task.MigratingProgressPercentage
    $report.Errormsg = $task.Errormsg
    $report.Status = $task.Status
    $Report.Date = (Get-Date -Format "dd-MM-yyyy")
}
$toMigrate | Export-CSV -Path $csv -encoding UTF8 

$FailFile = $($Session.ReportFolderPath) + $Switch + "FailureSummaryReport.csv"
$Sumfile = $($Session.ReportFolderPath) + $Switch + "SummaryReport.csv"
if (Test-Path $FailFile) {
    $fail = Import-CSV $failFile
    $FailedFiles += $fail
}
$sum = Import-CSV $Sumfile
$summary += $sum

Log-Message -File $File -Status DONE -Message "Exported job information to $customerfile"

#zipping 
$number = $Session.ReportFolderPath.split("$switch").count - 2
$ArchiveName = $Session.ReportFolderPath.split("$switch")[$number] + ".zip"
$folderToZip = $session.ReportFolderPath

Compress-Archive -LiteralPath $folderToZip -DestinationPath $($UploadFolder + $switch + $ArchiveName) -CompressionLevel Fastest

$Failed = $Session.StatusOfTasks | Where-Object { $_.Status -ne "COMPLETED" }

Stop-SPMTMigration
Unregister-SPMTMigration

Log-Message -file $file -Status WARNING -Message "Sharepoint Migration Tool disconnected."

Log-Message -file $file  " "

while ($Failed.count -ne 0) {
    Log-Message -File $file -Status WARNING -Message "Restarting failed migrations: $($Failed.count) tasks needs to be re-added"
    If ($Customer.weblogin -eq $true) {
        Write-Host "Logging in SPMT with web credentials"
        if ((!($CreatedAfter)) -and (!($ModifiedAfter))) {
            Connect-SPMT -LoginWithWeb $true 
        }
        else {
            Connect-SPMT -LoginWithWeb $true -createdAfter $CreatedAfter -ModifiedAfter $ModifiedAfter
        }
    }
    else {
        $Global:SPOUrl = $Customer.SPOurl
        $Global:UserName = $Customer.Username
        $Global:PassWord = ConvertTo-SecureString -String $Customer.password -AsPlainText -Force
        $Global:SPOCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Global:UserName, $Global:PassWord
    
        if ((!($CreatedAfter)) -and (!($ModifiedAfter))) {
            Connect-SPMT
        }
        else {
            Connect-SPMT -createdAfter $CreatedAfter -ModifiedAfter $ModifiedAfter
        }
    }
    ForEach ($Task in $Failed) { 
        $FolderInfo = $ToMigrate | Where-Object { $_.Homedir -eq $Task.SourceURI }
        if ($FolderInfo.folder -eq "") {
            $job = Add-SPMTTask -FileShareSource $Task.SourceURI -TargetSiteUrl $Task.TargetUri -TargetList "Documents"
        }
        else {
            $Job = Add-SPMTTask -FileShareSource $Task.SourceURI -TargetSiteUrl $Task.TargetUri  -TargetList "Documents"  -TargetListRelativePath $Folderinfo.Folder 
        }
        Log-Message -File $file -Status ADDED -Message "Added $($task.SourceURI) to $($Task.TargetURI) again"

    }
    
    Try {
        $Activity = "Starting Onedrive migration"
        Write-Progress -Activity $Activity
        Start-SPMTMigration -NoShow
        Log-Message -file $file -Status Done -Message "Restart of OneDrive migrations is started"
        Write-Progress -Activity $Activity -Completed
    }
    catch {
        $StartError = $_
        Log-Message -file $file -Status ERROR -Message "$($StartError.Exception.Message)"
        Write-Progress -Activity $Activity -Completed
        break
    }

    $session = Get-SPMTMigration
    Log-Message -File $File -Status WAITING -Message "Waiting for job to complete. Check $CSV for current status"
    while ($session.Status -ne "Finished") {

        Write-Host "[" -NoNewline
        Write-Host "WAITING" -ForegroundColor Yellow -NoNewline
        Write-Host "] - $(Get-Date -Format "dd-MM-yyyy HH:mm") Waiting for job to complete. Check $CSV for current status. Refreshing every 60 seconds"
        ForEach ($Task in $Session.StatusOfTasks) {
            $Report = $ToMigrate | Where-Object { $_.homedir -eq $Task.SourceURI }
            $Report.ScannedTotalFiles = $Task.NumScannedTotalFiles
            $Report.ScannedBadFiles = $task.NumScannedBadFiles
            $report.FilesToBeMigrated = $task.NumFileWillBeMigrated
            $report.MigratedFiles = $task.NumActuallyMigratedFiles
            $report.FailedFiles = $task.NumFailedFiles
            $report.MigratingProgressPercentage = $task.MigratingProgressPercentage
            $report.Errormsg = $task.Errormsg
            $report.Status = $task.Status
            $Report.Date = (Get-Date -Format "dd-MM-yyyy")
        }
        $toMigrate | Export-CSV -Path $csv -encoding UTF8 
        Start-Sleep -seconds 60
    }
    $Failed = $Session.StatusOfTasks | Where-Object { $_.Status -ne "COMPLETED" }
    $FailFile = $($Session.ReportFolderPath) + $Switch + "FailureSummaryReport.csv"
    $Sumfile = $($Session.ReportFolderPath) + $Switch + "SummaryReport.csv"
    if (Test-Path $FailFile) {
        $fail = Import-CSV $failFile
        $FailedFiles += $fail
    }
    $sum = Import-CSV $Sumfile
    $summary += $sum
    Unregister-SPMTMigration
    #zipping 
    $number = $Session.ReportFolderPath.split("$switch").count - 2
    $ArchiveName = $Session.ReportFolderPath.split("$switch")[$number] + ".zip"
    $folderToZip = $session.ReportFolderPath

    Compress-Archive -LiteralPath $folderToZip -DestinationPath $($UploadFolder + $switch + $ArchiveName) -CompressionLevel Fastest

} 

Log-Message -File $File -Status FINISHED -Message "Migration session is completed"
$ExportFile = $UploadFolder + $Switch + "$(Get-Date -Format "dd-MM-yyyy HHmm")-OneDriveMigrationResults.xlsx"
$ExportCSV = $UploadFolder + $Switch + "$(Get-Date -Format "dd-MM-yyyy HHmm")-OneDriveMigrationResults.csv"
$ChangedCSV = $UploadFolder + $Switch + "$(Get-Date -Format "dd-MM-yyyy HHmm")-OneDriveMigrationChanged.csv"

$Summary | Export-Excel -WorksheetName "Report" -Path $ExportFIle -AutoSize -AutoFilter -BoldTopRow
$Summary | Export-CSV -Path $exportcsv -Encoding UTF8

if ($FailedFiles) {
    $FailedFiles | Where-Object { $_.Message -ne "Scan File Failure:The parent folder was not migrated" } | Export-Excel -WorksheetName "Failed Files" -Path $ExportFile -AutoSize -AutoFilter -BoldTopRow
    $FailedFiles | Where-Object { $_.Message -ne "Scan File Failure:The parent folder was not migrated" } | Export-CSV -Path $($UploadFolder + $Switch + "$(Get-Date -Format "dd-MM-yyyy HHmm")-OneDriveMigrationResults-FailedFiles.csv")
}
If ($FailedFolderCheck) {
    $FailedFolderCheck | Export-Excel -WorksheetName "FailedFolderCheck" -Path $ExportFile -AutoSize -AutoFilter -BoldTopRow
    $FailedFolderCheck | Export-CSV -Path $($UploadFolder + $Switch + "$(Get-Date -Format "dd-MM-yyyy HHmm")-OneDriveMigrationResults-FailedFolders.csv")

}

$Files = Get-ChildItem -Path ($Session.ReportFolderPath) -File -Recurse
$Files = $files | Where-Object { $_.Name -like "ItemReport*.csv" }
$Changed = @()
ForEach ($rep in $Files) {
    $add = Import-CSV $rep.FullName
    $Add = $add | Where-Object { $_.Message -eq "Migrated SuccessFully" }
    $Changed += $add
}

$Changed | Export-CSV $ChangedCSV

Log-Message -File $file -Status Done -Message "Migration results are exported to $ExportFile"
# sending results per email
$Message = "<html><head>"
$message += "<style>
    table, th, td {
      border: 1px solid black;
      border-collapse: collapse;
    }
    th, td {
      padding: 5px;
    }
    th {
      text-align: left;
    }
    </style><body>"
$Message += "Hi,<br>"
$message += "<br>"
$Message += "The migration is completed. <br>"
$message += "<br>"
$message += "In total there are $($Results.count) jobs completed.<br>"
$message += "Logging can be found in $($session.ReportFolderPath)<br>"
$message += "<table><tr><th>Folder</th><th>OneDrive</th><th>TotalScannedItems</th><th>ToBeMigrated</th><th>Migrated</th><th>NotMigrated</th><th>StartTime</th><th>EndTime</th><th>Duration</th></tr>"
ForEach ($line in $results) {
    $message += "<tr><th>$($Line.Source)</th><th>$($Line.Destination)</th><th>$($Line.'Total scanned item')</th><th>$($Line.'Total to be migrated')</th><th>$($Line.'Migrated Items')</th><th>$($Line.'Items not migrated')</th><th>$($Line.'Start Time')</th><th>$($Line.'End Time')</th><th>$($Line.Duration)</th></tr>"
}
$message += "<br>"
$message += "Kind regards,<br>"
$Message += "<br>"
$message += "Rapid Circle Migration Street"

Try {
    Send-MailToInform -To $($Customer.MigrationMail) -From $($customer.username) -Subject "Migrations @ $($customer.CompanyName)" -Message $message
    Log-message -file $file -Status Done -Message "Mail is sent to $($Customer.MigrationMail)"
}
Catch {
    $Fault = $_
    Log-message -file $file -Status ERROR -Message "Message was not sent: $($fault.Exception.Message)"
}

# removing service account OneDrive

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

try {
    $copy = Copy-Item -Path $customerfile -Destination $UploadFolder
    Log-Message -file $file -Status DONE -Message "Imported $CustomerFile is copied to $uploadfolder"
}
catch {
    # do nothing
}

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
            $folderlocation = "Homefolder%20Migrations/SPMT%20logging/$(Get-Date -Format "dd-MM-yyyy")"
        }
        else {
            $folderlocation = "Homefolder%20Migrations/custom%20logging/$(Get-Date -Format "dd-MM-yyyy")"
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
Log-Message -file $file  "End Onedrive migrations $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "

if ($Customer.Teams){
    # send mail to teams channel
    $apiUri = "https://graph.microsoft.com/v1.0/Teams/$($Sharepoint.Ids.TeamsId)/channels"
    $channel = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.access_token)" } -Uri $apiUri -Method Get).value | Where-Object { $_.DisplayName -eq $sharepoint.Homedir }
    $emailTo = $channel.email

    $message = "<b>The homefolder migrations are done</b><br>"
    $message += "<br>"
    $message += "The logging can be found in this channel files and folders.<br>"
    $message += "A total of $($CorrectFolderCheck.count) shares are migrated to the requested Onedrives<br>"
    $message += "<br>"
    $message += "Kind regards,<br>"
    $message += "The Rapid Circle Migration Street"

    Send-MailToinform -to $emailTo -From $customer.MigrationMail -Subject "$($Customer.CompanyName) - $(Get-Date -Format "dd-MM-yyyy hh:mm") Homefolder migrations are done" -Message $message
}
