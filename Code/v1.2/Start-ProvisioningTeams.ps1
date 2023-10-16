
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

$Global:file = "." + $switch + "log" + $switch + "$(Get-Date -format "ddMMyyy-HHmm")-TeamsProvisioning.txt"

# start provisioning teams
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "Start script for Teams provisioning $(Get-date -format "dd-MM-yyyy - HH:mm")"
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

## getting customer data from excel sheet
$CustomerFile = "." + $switch + "CustomerData" + $switch + "MigrationStreet.xlsx"
$WorksheetName = "TeamsToProvision"

try {
    $TeamsToProvision = Import-Excel -Path $CustomerFile -WorksheetName $WorksheetName -ErrorAction Stop
    Log-Message -file $file -Status done -Message "Teams sheet to provision is imported: $CustomerFile"
}
catch {
    $ProvisionError = $_
    Log-Message -file $file -Status error -Message "$($ProvisionError.Exception.Message)"
    break
}

$WorksheetName = "ChannelsToProvision"

try {
    $ChannelsToProvision = Import-Excel -Path $CustomerFile -WorksheetName $WorksheetName -ErrorAction Stop
    Log-Message -file $file -Status done -Message "Channels sheet to provision is imported: $CustomerFile"
}
catch {
    $ProvisionError = $_
    Log-Message -file $file -Status error -Message "$($ProvisionError.Exception.Message)"
    break
}


Log-Message -file $file " "
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "Start provisioning teams $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "


## Filter which is done 
[array]$TeamsToDo = $TeamsToProvision | Where-Object { $_.Status -eq "New" }
If ($TeamsToDo -eq $null) {
    Log-Message -File $file ""
    Log-Message -file $file -Status error -Message "NO New Teams are needed to be provsioned"
    break
}

Log-Message -file $file -Status ADDED -Message "Adding collumns TeamsId and CreationDateTime to array"
$TeamsToDo | Add-Member -MemberType NoteProperty -Name "TeamsId" -Value "" -force
$TeamsToDo | Add-Member -MemberType NoteProperty -Name "CreationDateTime" -Value "" -force
$TeamsToDo | Add-Member -MemberType NoteProperty -Name "Information" -Value "" -force


[array]$ChannelsToDo = $ChannelsToProvision | Where-Object { $_.Status -eq "New" }
If ($ChannelsToDo -eq $null) {
    Log-Message -File $file ""
    Log-Message -file $file -Status error -Message " NO New channels are needed to be provsioned"
    break
}
Log-Message -file $file -Status ADDED -Message "Adding collumns ChannelId and CreationDateTime to array"
$ChannelsToDo | Add-Member -MemberType NoteProperty -Name "TeamsId" -Value "" -force
$ChannelsToDo | Add-Member -MemberType NoteProperty -Name "ChannelId" -Value "" -force
$ChannelsToDo | Add-Member -MemberType NoteProperty -Name "webUrl" -Value "" -force
$ChannelsToDo | Add-Member -MemberType NoteProperty -Name "CreationDateTime" -Value "" -force
$ChannelsToDo | Add-Member -MemberType NoteProperty -Name "Information" -Value "" -force



## getting current teams
try {
    $apiUri = "https://graph.microsoft.com/v1.0/teams"
    $Teams = RunQueryandEnumerateResults -ApiUri $apiUri
    Log-Message -file $file -Status done -Message "All current teams from tenant are indexed"
}
catch {
    $ProvisionError = $_
    Log-Message -file $file -Status error -Message "$($ProvisionError.Exception.Message)"
    break
}


ForEach ($team in $TeamsToDo) {
    $Activity = "Creating teams: $($Team.Teamname)"
    Write-Progress -Activity $Activity 
    
    if ($Teams.DisplayName -contains $team.TeamName) {
        Log-Message -file $file -Status SKIPPING -Message "Skipping $($Team.Teamname) for creation because it already exist with this name"
    }
    else {
        $SetAdditionalOwner = $False
        $Owners = $Team.Owner.Split(",")
        If ($Owners -gt 1) {
            $Owner = $Owners[0]
            $SetAdditionalOwner = $true
        }
        else {
            $Owner = $Team.Owner
        }

        $Members = $null
        $SetMembers = $false
        try {
            $Members = $Team.Members.split(",")
        }
        catch {
            # do nothing
        }
        if ($Members -gt 1) {
            $SetMembers = $True
        }

        try {
            $Create = Create-Teams -Teamsname $Team.TeamName -Owner $owner -Description "A new team with the name $($Team.TeamName)"

            do {
                $Apiuri = "https://graph.microsoft.com/v1.0/teams" + '?$filter=startsWith(displayName, ' + "'$($team.TeamName)')"
                $check = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.access_token)" } -Uri $apiUri -Method Get).value
                start-sleep -seconds 2
            } while ($Check.count -eq 0)

            ForEach ($Item in ($TeamsToProvision | Where-Object { $_.TeamName -eq $Team.TeamName })) {
                $item.TeamsId = $check.id
            }
            ForEach ($Item in ($ChannelsToProvision | Where-Object { $_.TeamName -eq $Team.TeamName })) {
                $item.TeamsId = $check.id
            }

            # creating a folder which will be delete also but is to trigger the provisioning
            $apiUri = "https://graph.microsoft.com/v1.0/Teams/$($check.id)/channels"
            $channels = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.access_token)" } -Uri $apiUri -Method Get).value | Where-Object { $_.DisplayName -eq "General" }
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
            
            
            #deleting folder
            $apiUri = "https://graph.microsoft.com/v1.0/drives/$driveItemId/items/$($createFolder.id)"
            Invoke-RestMethod -Method DELETE -uri $apiUri -Headers $header


            Log-Message -file $file -Status DONE -Message "$($Team.Teamname) is created with owner: $owner"
            Write-progress -Activity $Activity -Completed

            if ($SetAdditionalOwner -eq $True) {
                $Owners = $Owners | Select-Object -Skip 1

                ForEach ($Owner in $Owners) {
                    $Activity = "Adding owner to $($Team.Teamname) with upn $Owner"
                    Write-Progress -Activity $Activity 
                    
                    try {
                        $Set = Add-OwnerToTeams -TeamsName $Team.TeamName -upn $Owner -TeamsId $Check.id
                        Log-Message -file $file -Status added -Message "- Additional owner set to $($Team.TeamName) with the name $Owner"
                        Write-Progress -Activity $Activity -Completed
                    }
                    catch {
                        $FaultError = $_
                        Log-Message -file $file -Status error -Message "$($FaultError.Exception.Message)"
                        Write-Progress -Activity $Activity -Completed
                    }
                    Write-Progress -Activity $Activity -Completed
                }
            }
            if ($SetMembers -eq $True) {
                ForEach ($Member in $members) {
                    $Activity = "Adding member to $($Team.TeamName) with upn $Member"
                    Write-Progress -Activity $Activity 
                    
                    try {
                        $Set = Add-MemberToTeams -TeamsId $Team.TeamsId -upn $Member
                        Log-Message -file $file -Status added -Message "- Additional member set to $($Team.TeamName) with the name $Member"
                        Write-Progress -Activity $Activity -Completed
                    }
                    catch {
                        $FaultError = $_
                        Log-Message -file $file -Status error -Message "$($FaultError.Exception.Message)"
                        Write-Progress -Activity $Activity -Completed
                    }
                }
            }
            $Team.Status = "DONE"
            $Team.CreationDateTime = Get-Date -Format "dd-MM-yyy HH:mm"
            $Team.information = "Succesfull created"
        }
        catch {
            $CreateError = $_
            Log-Message -file $file -Status error -Message "Error creating $($Team.Teamname) - $($CreateError.Exception.Message)"
            $Team.Status = "ERROR"
            $Team.CreationDateTime = Get-Date -Format "dd-MM-yyy HH:mm"
            $Team.information = "Error creating teams : $($CreateError.Exception.Message)"
            Write-Progress -Activity $Activity -Completed
        }

    }
}

Log-Message -File $file ""
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "Adding additional channels to created teams"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "

# adding channels

ForEach ($team in $ChannelsToDo) {
    switch ($Team.Type) {
        private { $Membership = "Private" }
        standard { $membership = "Standard" }
        shared { $Membership = "Shared" }
    }
    $SetAdditionalOwner = $False
    $Owners = $Team.Owner.Split(",")
    If ($Owners -gt 1) {
        $Owner = $Owners[0]
        $SetAdditionalOwner = $true
    }
    else {
        $Owner = $Team.Owner
    }

    $Members = $null
    $SetMembers = $false
    try {
        $Members = $Team.Members.split(",")
    }
    catch {
        # do nothing
    }
    if ($Members -gt 1) {
        $SetMembers = $True
    }

    
    try {
        $Activity = "Adding channel with name $($Team.channelname) to team $($Team.TeamName)"
        Write-Progress -Activity $Activity 
        $Create = Create-Channel -TeamsId $Team.Teamsid -ChannelName $Team.ChannelName -Description $Team.Description -owner $Owner -Type $membership
        Log-Message -file $file -Status done -Message "$($Team.ChannelName) is created with owner: $owner in $($team.TeamName)"
        $Team.ChannelId = $Create.id
        $Team.Status = "DONE"
        $Team.CreationDateTime = Get-Date -Format "dd-MM-yyy HH:mm"
        $Team.information = "Succesfull created"
        Write-Progress -Activity $Activity -Completed
    }
    catch {
        $CreateError = $_
        Log-Message -file $file -Status error -Message "Error creating $($Team.Teamname) - $($CreateError.Exception.Message)"
        $Team.Status = "ERROR"
        $Team.Information = $($CreateError.Exception.Message)
        Write-Progress -Activity $Activity -Completed
    }


    if ($SetAdditionalOwner -eq $True) {
        $Owners = $Owners | Select-Object -Skip 1

        ForEach ($Owner in $Owners) {
            $Activity = "Adding owner to teams channel $($Team.ChannelName) with upn $owner"
            Write-Progress -Activity $Activity 
            try {
                $Set = Add-OwnerToTeamsChannel -TeamsId $Team.TeamsId -upn $Owner -ChannelId $team.ChannelId
                Log-Message -file $file -Status added -Message "- Additional owner set to $($Team.ChannelName) in team $($Team.TeamName) with the name $Owner"
                Write-Progress -Activity $Activity -Completed
            }
            catch {
                $FaultError = $_
                Log-Message -file $file -Status error -Message "$($FaultError.Exception.Message)"
                Write-Progress -Activity $Activity -Completed
            }
        }
    }
    if ($SetMembers -eq $True) {
        ForEach ($Member in $members) {
            $Activity = "Adding member $Member to $($Team.ChannelName)"
            Write-Progress -Activity $Activity 
            try {
                $Set = Add-MemberToTeamsChannel -TeamsId $Team.TeamsId -upn $Member -ChannelId $team.ChannelId
                Log-Message -file $file -Status added -Message "- Additional member added to $($Team.TeamName) with channel $($Team.ChannelName) with the name $Member"
                Write-Progress -Activity $Activity -Completed
            }
            catch {
                $FaultError = $_
                Log-Message -file $file -Status error -Message "$($FaultError.Exception.Message)"
                Write-Progress -Activity $Activity -Completed
            }
        }
    }

    if ($Membership -eq "private"){
        do {
            try {  
                $folderlocation = (Invoke-Restmethod -method Get -headers @{Authorization = "Bearer $($Token.access_token)" } -Uri https://graph.microsoft.com/v1.0/teams/$($Team.TeamsId)/channels/$($Create.id)/filesFolder)
                $mustRetry = 0
            }
            catch {
                $weberror = $_
                $mustRetry = 1
                Start-Sleep -seconds 2
            }
            
        } while (
            $mustRetry -eq 1
        )
    
        $Team.WebUrl = $Folderlocation.Weburl
    }

}

Log-Message -File $file ""

$TeamsToDo | Export-Excel -WorksheetName "TeamsToProvision" -Path $CustomerFile -AutoSize -BoldTopRow -FreezeTopRow
$ChannelsToDo | Export-Excel -WorkSheetname "ChannelsToProvision" -path $CustomerFile -AutoSize -BoldTopRow -FreezeTopRow

Log-Message -file $file -Status done -Message "Write back XLSX with updated data"
Log-Message -File $file ""
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "End provisioning teams $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "



