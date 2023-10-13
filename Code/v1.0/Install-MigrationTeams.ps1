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
    
$Global:file = "." + $switch + "log" + $switch + "$(Get-Date -format "ddMMyyy-HHmm")-InstallMigrationTeams.txt"
    
# start provisioning teams
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "Start script for creating Rapid Circle Migration Team $(Get-date -format "dd-MM-yyyy - HH:mm")"
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

#importing sharepoint site setup information

try {
    $Activity = "Importing Rapid Circle details for migration sharepoint"
    Write-Progress -Activity $Activity 
    $sharepoint = Import-PowerShellDataFile -Path $("." + $Switch + "import-data" + $switch + "teams-and-spo-information.psd1") -ErrorAction Stop
    Log-Message -file $file -status DONE -Message "Details are imported"
    Write-Progress -Activity $Activity -Completed
}
catch {
    $connectionError = $_
    Log-Message -file $file -Status Error -Message "$($connectionError.Exception.Message)"
    Write-Progress -Activity $Activity -Completed
    break
}


#region creating teams 
try {
    $apiUri = "https://graph.microsoft.com/v1.0/teams"
    $Teams = RunQueryandEnumerateResults -apiUri $apiUri
    Log-Message -file $file -Status done -Message "All current teams from tenant are indexed"
}
catch {
    $ProvisionError = $_
    Log-Message -file $file -Status error -Message "$($ProvisionError.Exception.Message)"
    break
}

#checking if teams exists
if ($Teams.displayname -eq $Sharepoint.Teamsname){
    Log-Message -File $file -Status Error -Message "Teams with the name $($Sharepoint.TeamsName) already exists"
    break
}

Log-Message -File $File -Status WAITING -Message "Creating Teams with the name $($Sharepoint.TeamsName)"
try {
    $Create = Create-Teams -Teamsname $Sharepoint.TeamsName -Owner $Customer.UserName -Description "A new team with the name $($Sharepoint.TeamsName) is created"

    do {
        $Apiuri = "https://graph.microsoft.com/v1.0/teams" + '?$filter=startsWith(displayName, ' + "'$($Sharepoint.TeamsName)')"
        $check = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.access_token)" } -Uri $apiUri -Method Get).value
        start-sleep -seconds 2
    } while ($Check.count -eq 0)
    
    $sharepoint.Ids.TeamsId = $check.Id

    # creating a folder which will be delete also but is to trigger the provisioning
    $apiUri = "https://graph.microsoft.com/v1.0/Teams/$($check.id)/channels"
    $channels = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.access_token)" } -Uri $apiUri -Method Get).value | Where-Object { $_.DisplayName -eq "General" }
    $Sharepoint.ids.GeneralChannelId = $channels.id
    $drive = Invoke-Restmethod -method Get -headers @{Authorization = "Bearer $($Token.access_token)" } -Uri https://graph.microsoft.com/v1.0/teams/$($check.id)/channels/$($Channels.id)/filesFolder
    $driveId = $drive.id
    $driveitemId = $drive.parentReference.DriveId
    $folderToCreate = "TemporaryFolder"
    $header = @{
        'Authorization' = "BEARER $($Token.access_token)"
        'Content-type'  = "application/json"
    }
    $params = @{
        name                                = $FolderToCreate
        folder                              = @{
        }
        "@microsoft.graph.conflictBehavior" = "rename"
    }
    $params = $params | ConvertTo-Json
    $createFolder = Invoke-Restmethod -Method POST -Uri "https://graph.microsoft.com/v1.0/drives/$driveitemid/items/$DriveId/children" -Headers $header -body $params
    Log-Message -File $file -Status DONE -Message "- Sharepoint drive is provisoned"
    
    #deleting folder
    $apiUri = "https://graph.microsoft.com/v1.0/drives/$driveItemId/items/$($createFolder.id)"
    Invoke-RestMethod -Method DELETE -uri $apiUri -Headers $header


    Log-Message -File $file -Status Done -Message "Teams is created with the name $($Sharepoint.Teamsname)"

}
catch {
    $CreateError = $_
    Log-Message -file $file -Status error -Message "Error creating $($Sharepoint.Teamsame) - $($CreateError.Exception.Message)"
    break
}

#endregion creating teams

#region creating channels
ForEach ($Channel in $Sharepoint.Channels) {
    try {
        $Membership = "Standard"
        $Activity = "Adding channel with name $Channel to team $($Sharepoint.TeamsName)"
        Write-Progress -Activity $Activity 
        $Create = Create-Channel -TeamsId $Check.id -ChannelName $Channel -Description "Channel for $Channel" -owner $Customer.Username -Type $membership

        Log-Message -file $file -Status done -Message "$Channel channel is created with owner: $($Customer.Username) in $($Sharepoint.TeamsName)"

        # creating a folder which will be delete also but is to trigger the provisioning
        $apiUri = "https://graph.microsoft.com/v1.0/Teams/$($check.id)/channels"
        $channels = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.access_token)" } -Uri $apiUri -Method Get).value | Where-Object { $_.DisplayName -eq $Channel }
        
        switch ($channel){
            "Fileshare Migrations" { 
                $Sharepoint.ids.FileshareChannelid = $Channels.id
                $sharepoint.Mails.FileshareChannel = $channels.email
            }
            "Homefolder Migrations" { 
                $Sharepoint.ids.HomeFolderChannelid = $Channels.id
                $sharepoint.mails.HomeFolderChannel = $channels.email
            }
            "Errors" { 
                $Sharepoint.ids.ErrorChannelid = $Channels.id
                $sharepoint.mails.ErrorChannel = $channels.id
            }
        }

        $drive = Invoke-Restmethod -method Get -headers @{Authorization = "Bearer $($Token.access_token)" } -Uri https://graph.microsoft.com/v1.0/teams/$($check.id)/channels/$($Channels.id)/filesFolder
        $driveId = $drive.id
        $driveitemId = $drive.parentReference.DriveId
        $foldersToCreate = "SPMT Logging", "Custom Logging"
        $header = @{
            'Authorization' = "BEARER $($Token.access_token)"
            'Content-type'  = "application/json"
        }
        foreach ($FolderToCreate in $FoldersToCreate) {
            $params = @{
                name                                = $FolderToCreate
                folder                              = @{
                }
                "@microsoft.graph.conflictBehavior" = "rename"
            }
            $params = $params | ConvertTo-Json
            $createFolder = Invoke-Restmethod -Method POST -Uri "https://graph.microsoft.com/v1.0/drives/$driveitemid/items/$DriveId/children" -Headers $header -body $params
            Log-Message -File $file -Status DONE -Message "- $FolderToCreate folder is created in $channel"
        }

        
        Write-Progress -Activity $Activity -Completed
    }
    catch {
        $CreateError = $_
        Log-Message -file $file -Status error -Message "Error creating $channel Channel - $($CreateError.Exception.Message)"
        $Team.Status = "ERROR"
        $Team.Information = $($CreateError.Exception.Message)
        Write-Progress -Activity $Activity -Completed
    }
}

#endregion creating channels
Log-Message -File $file -Status WAITING -Message "Waiting for Sharepoint site to be provisioned"
do {
    $apiUri = "https://graph.microsoft.com/v1.0/sites"
    $sites = RunQueryandEnumerateResults -ApiUri $apiUri
    start-sleep -seconds 2
} while (!($sites | Where-Object {$_.Displayname -eq $Sharepoint.TeamsName}))

$sharepoint.SharepointSite = ($sites | Where-Object {$_.Displayname -eq $Sharepoint.TeamsName}).weburl

#exporting Channel information
$PsdFile = "." + $switch + "import-data" + $switch + "teams-and-spo-information.psd1"
ConvertTo-Psd -InputObject $sharepoint -Depth 4 | out-File $PsdFile
Log-Message -File $file -Status DONE -Message "All teams and sharepoint information exported to $psdfile"

Log-Message -File $file -Message ""
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "End creating Rapid Circle Migration Teams $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "

