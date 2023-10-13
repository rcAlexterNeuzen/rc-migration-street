#group.readwrite.all

#region default for each script
Function Log-Message([String]$Message, [string]$file) {
    $message = "[$(Get-Date -format "HH:mm:ss")] " + $message
    Add-Content -Path $file $Message
}
if (!($isMacOs)) {
    $Switch = "\"
}
else {
    $Switch = "/"
}

$file = "." + $switch + "log" + $switch + "$(Get-Date -format "ddMMyyy-HHmm")-change-teams-group-to-dynamic.txt"

# start provisioning teams
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "Start script for changing groups to dynamic $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "

Write-Host "----------------------------------------------"
Write-Host  "Start script for changing groups to dynamic $(Get-date -format "dd-MM-yyyy - HH:mm")"
Write-Host  "----------------------------------------------"
Write-Host  " "

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
Write-host "[" -NoNewline
Try {
    $modules = Get-ChildItem -Path $("." + $Switch + "modules") -file
    $modules = $modules | Where-Object { $_.Extension -eq ".psm1" }
    Write-Host "DONE" -ForegroundColor Green -NoNewline
    Write-Host "] - " -nonewLine 
    Write-Host "Getting modules from $("." + $Switch + "modules")" -ForegroundColor Yellow
    Log-Message -file $file "[DONE] - Getting modules from $("." + $Switch + "modules")"
}
Catch {
    $moduleError = $_
    Write-Host "ERROR" -ForegroundColor Red -NoNewline
    Write-Host "] - " -NoNewline
    Write-Host "$($moduleError.Exception.Message)" -ForegroundColor red
    Log-Message -file $file "[ERROR] - $($moduleError.Exception.Message)"
    break
}

ForEach ($Module in $modules) {
    Write-Host "[" -NoNewline
    try {
        Import-Module $Module.FullName -Force -DisableNameChecking -ErrorAction Stop
        Write-Host "DONE" -ForegroundColor Green -NoNewline
        Write-Host "] - - Imported $($module.FullName)"
        Log-Message -File $file "[DONE] - - Imported $($module.FullName)"
    }
    catch {
        $moduleError = $_
        Write-host "ERROR" -ForegroundColor Red -NoNewline
        Write-host "] - Error importing $($module.FullName) : $($moduleError.Exception.Message)"
        Log-Message -file $file "[ERROR] - - Error importing $($module.FullName) : $($moduleError.Exception.Message)"
        break
    }
}

# getting datafiles
Write-host "[" -NoNewline
Try {
    $datafiles = Get-ChildItem -Path $("." + $Switch + "import-data") -file -ErrorAction Stop
    if (!($datafiles)) {
        Write-Host "ERROR" -ForegroundColor Red -NoNewline
        Write-Host "] - There are no PowerShell Data files found in $("." + $Switch + "import-data")"
        Log-Message -file $file "[ERROR] - There are no PowerShell Data files found in $("." + $Switch + "import-data") "
        break
    }
    $datafiles = $datafiles | Where-Object { $_.Extension -eq ".psd1" }
    Write-Host "DONE" -ForegroundColor Green -NoNewline
    Write-Host "] - " -nonewline
    Write-Host "Getting data files to import from $("." + $Switch + "import-data")" -ForegroundColor Yellow
    Log-Message -file $file "[DONE] - Getting data files to import from $("." + $Switch + "import-data") "
}
Catch {
    $dataError = $_
    Write-Host "ERROR" -ForegroundColor Red -NoNewline
    Write-Host "] - " -NoNewline
    Write-Host "$($dataError.Exception.Message)" -ForegroundColor red
    Log-Message -file $file "[ERROR] - $($dataError.Exception.Message) "
    break
}

## if needed installing excel module
if ($installExcel) {
    Install-Requirements -module "ImportExcel"
    Log-Message -file $file "[DONE] - Module ImportExcel is installed"
}

## getting security info
Write-host "[" -NoNewline
try {
    $connectionDetails = Import-clixml -Path $("." + $Switch + "security" + $switch + "connectionDetails.xml") -ErrorAction Stop
    Write-Host "DONE" -ForegroundColor Green -NoNewline
    Write-Host "] - Connection details are imported"
    Log-Message -file $file "[DONE] - Connection details are imported"
}
catch {
    $connectionError = $_
    Write-Host "ERROR" -ForegroundColor Red -NoNewline
    Write-Host "] - $($connectionError.Exception.Message)"
    Log-Message -file $file "[ERROR] - $($connectionError.Exception.Message)"
    break
}

## getting token
Write-Host "[" -NoNewLine
try {
    ## getting token 
    $token = Get-TokenForGraphAPIWithCertificate -appid $connectionDetails.appid -tenantname $connectionDetails.tenantname -Thumbprint $connectionDetails.ThumbPrint
    if ($Token.access_token -ne $null) {
        Write-Host "DONE" -ForegroundColor Green -NoNewLine
        Write-Host "] - Token for Graph API is present"
        Log-Message -File $file "[DONE] - Token for Graph API is present"
    }
    else {
        Write-Host "WARNING" -ForegroundColor Yellow -NoNewline
        Write-Host "] - No token was retrieved for Graph API. Trying again ..."
        Log-Message -File $file "[WARNING] - No token was retrieved for Graph API. Trying again ..."
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

Write-host "[" -NoNewline
try {
    $TeamsToProvision = Import-Excel -Path $CustomerFile -WorksheetName $WorksheetName -ErrorAction Stop
    Write-Host "DONE" -ForegroundColor Green -NoNewline
    Write-Host "] - Teams sheet to provision is imported"
    Log-Message -file $file "[DONE] - Teams sheet to provision is imported: $CustomerFile"
}
catch {
    $ProvisionError = $_
    Write-Host "ERROR" -ForegroundColor Red -NoNewline
    Write-Host "] - $($ProvisionError.Exception.Message)"
    Log-Message -file $file "[ERROR] - $($ProvisionError.Exception.Message)"
    break
}

$TeamsToChange = $TeamsToProvision | Where-Object {$_.dynamic -eq $true}

ForEach ($TeamToChange in $TeamsToChange) {
    # getting group id
    #$apiUri = "https://graph.microsoft.com/v1.0/groups" + '?$filter=startsWith(displayName, ' + "'$($TeamToChange.TeamName)')"
    $apiUri = "https://graph.microsoft.com/v1.0/groups" + '?$filter=startsWith(displayName, ' + "'TestGroupDynamic1')"
    $GroupId = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.access_token)" } -Uri $apiUri -Method Get).value.id

    try {
        $Change = Set-DynamicGroup -GroupId $GroupId -MembershipRule $TeamToChange.DynamicRule
        Write-host "[" -NoNewline
        Write-Host "DONE" -ForegroundColor Green -NoNewline
        Write-Host "] - $($TeamToChange.TeamName) group is changed to Dynamic with rule: $MembershipRule"
        Log-Message -file $file "[DONE] - $($TeamToChange.TeamName) group is changed to Dynamic with rule: $MembershipRule"
    }
    catch {
        $ChangeError = $_
        Write-host "[" -NoNewline
        Write-Host "ERROR" -ForegroundColor Red -NoNewline
        Write-Host "] - $($ChangeError.Exception.Message)"
        Log-Message -file $file "[ERROR] - $($ChangeError.Exception.Message)"
        break
    }
}






