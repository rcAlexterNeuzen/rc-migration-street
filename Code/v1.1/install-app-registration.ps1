# create app registration for teams provisioning
#region appregistration
#first we create an app registration with the required permissions
if (!($isMacOs)) {
    $Switch = "\"
}
else {
    $Switch = "/"
}

Add-Type -AssemblyName System.Security
Try {
    Import-Module ".\modules\rc-migration-module.psm1" -DisableNameChecking -ErrorAction Stop
    Import-Module ".\modules\rc-required-modules.psm1" -DisableNameChecking -ErrorAction Stop
}
catch {
        $fault = $_
        Write-Host "ERROR importing the default MIGRATION module needed for logging and connections" -ForegroundColor Red
        write-Host "$($Fault.ErrorDetails.Message)"
        break
    }


$Global:file = "." + $switch + "log" + $switch + "$(Get-Date -format "ddMMyyy-HHmm")-install-app-registration.txt"

# start installation app registration
Log-Message -file $file "----------------------------------------------------------------"
Log-Message -file $file  "Start installation app registration $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------------------------"
Log-Message -file $file  " "


#importing module
Try {
    Import-Module Az.Accounts -Force -ErrorAction Stop
    Log-Message -file $file -Status Done -Message "AZ.Resources module imported"
}
Catch {
    $moduleError = $_
    Log-Message -file $file -Status error -Message "$($moduleError.Exception.Message)"

    
    try {
        Install-Module Az.Accounts -Force -ErrorAction Stop
        Log-Message -file $file -Status Done -Message "- Installed AZ.Resources module "
        Import-Module AZ.Accounts
    }
    catch {
        $moduleError = $_
        Log-Message -file $file -Status error -Message "$($moduleError.Exception.Message)"
        break
    }
}

Connect-AZAccount


try {
    $ImportData = Import-PowerShellDataFile $("." + $switch + "import-data" + $switch + "app-registration.psd1")
    Log-Message -file $file -Status Done -Message "App Registration details are imported"
}
catch {
    $connectionError = $_
    Log-Message -file $file -Status error -Message "$($connectionError.Exception.Message)"
    break
}


$DisplayName = $ImportData.AppRegistration.name
$Description = $ImportData.AppRegistration.Description
$signInAudience = $ImportData.AppRegistration.SignInAudience
$web = $ImportData.AppRegistration.web

try {
    $info = New-AzADApplication -DisplayName $DisplayName -Description $Description -SignInAudience $signInAudience -ReplyUrls $web 
    Log-Message -file $file -Status Done -Message "App Registration is created with the name $($ImportData.AppRegistration.name)"
}
catch {
    $connectionError = $_
    Log-Message -file $file -Status error -Message "$($connectionError.Exception.Message)"
    break
}

$Permissions = $importData.perm

ForEach ($permission in $permissions) {
    try {
        Add-AzADAppPermission -ObjectId $info.id -ApiId $importData.GraphApi -PermissionId $permission -Type Role 
        Log-Message -file $file -Status Done -Message "- Added permission with code: $permission"
    }
    catch {
        $connectionError = $_
        Log-Message -file $file -Status error -Message "$($connectionError.Exception.Message)"
        break
    }
    
}

do {
    Log-Message -file $file -Status WAITING -Message "[WAITING] - Waiting for app to be available ..."
    $check = Get-AzADAppPermission -ObjectId $info.id
    start-sleep -Seconds 2
} while (
    $Check = $null
) 

# wait before app is live in azure
Start-Sleep -Seconds 20

# giving admin consent to the new permissions
$Tenantid = (Get-AzContext).Tenant.Id
$applicationid = $info.AppId

Log-Message -file $file -Status WARNING -Message "Please provide admin consent for the application"
start-process https://login.microsoftonline.com/$Tenantid/adminconsent?client_id=$applicationid

$answer = Read-Host "Did you provide admin consent to the application (Y/N)?"
if ($answer -ne "Y") {
    Write-Host "It is required to give admin consent to the permissions"
    Log-Message -file $file -Status ERROR -Message "No correct answer was given for admin consent"
    break
}

# creating a secret for the app

try {
    $AppSecret = Get-AzADApplication -ApplicationId $applicationId | New-AzADAppCredential -StartDate (Get-Date) -EndDate (Get-Date).AddYears(2)
    Log-Message -file $file -Status Done -Message "Creating an app secret for $($ImportData.AppRegistration.name) application"
}
catch {
    $connectionError = $_
    Log-Message -file $file -Status error -Message "$($connectionError.Exception.Message)"
    break
}



# creating a secret for the app
$password = Get-RandomCharacters -length 9 -characters 'abcdefghiklmnoprstuvwxyz'
$password += Get-RandomCharacters -length 1 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
$password += Get-RandomCharacters -length 2 -characters '1234567890'
$password += Get-RandomCharacters -Length 2 -characters '!@#$'
$password = Scramble-String $password

$certname = $ImportData.AppRegistration.name
$CerExportPath = "." + $switch + "security" + $switch + $certname + ".cer"
$pfxExportPath = "." + $switch + "security" + $switch + $certname + ".pfx"

try {
    $mycert = New-SelfSignedCertificate -Subject "CN=$CertName" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256
    Log-Message -file $file -Status Done -Message "Creating certificate for $($ImportData.AppRegistration.name) application"
}
catch {
    $connectionError = $_
    Log-Message -file $file -Status error -Message "$($connectionError.Exception.Message)"
    break
}

# Export certificate to .pfx file
try {
    $mycert | Export-PfxCertificate -FilePath $pfxExportPath -Password (ConvertTo-SecureString $password -asplaintext -force) -InformationAction SilentlyContinue | Out-Null
    Log-Message -file $file -Status Done -Message "Exporting certificate pfx for $($ImportData.AppRegistration.name) application to $PfxExportPath"
}
catch {
    $connectionError = $_
    Log-Message -file $file -Status error -Message "$($connectionError.Exception.Message)"
    break
}

# Export certificate to .cer file
try {
    $mycert | Export-Certificate -FilePath $CerExportPath -InformationAction SilentlyContinue | Out-Null
    Log-Message -file $file -Status Done -Message "Exporting certificate cer for $($ImportData.AppRegistration.name) application to $CerExportPath"
}
catch {
    $connectionError = $_
    Log-Message -file $file -Status error -Message "$($connectionError.Exception.Message)"
    break
}

$passwordoutput = "." + $switch + "security" + $switch + "certificate_password.txt"
Log-Message -file $file -Status WARNING -Message "Your password for certificate is: $Password"
[Byte[]] $key = (1..16)
$password = $password | ConvertTo-SecureString -AsPlainText -Force
Log-Message -file $file -Status WARNING -Message "It will be saved to $passwordoutput encrypted"
$password | out-file $passwordoutput

$cer = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2((Get-ChildItem $CerExportPath).fullName)
$binCert = $cer.GetRawCertData() 
$credValue = [System.Convert]::ToBase64String($binCert)

try {
    $CertSecret = Get-AzADApplication -ApplicationId $applicationId | New-AzADAppCredential -CertValue $credValue -StartDate (Get-Date) -EndDate $Cer.NotAfter
    Log-Message -file $file -Status Done -Message "Uploading certificate for $($ImportData.AppRegistration.name) application"
}
catch {
    $connectionError = $_
    Log-Message -file $file -Status ERROR -Message "$($connectionError.Exception.Message)"
    break
}

# tenantname 
$tenantName = ((Get-AZTenant).name) + ".onmicrosoft.com"

# servicePrincipal
$ServicePrincipal = (Get-AzADServicePrincipal -ApplicationId $applicationId).Id

# combine all information
$connectionDetails = @{
    'TenantId'           = $Tenantid
    'Appid'              = $applicationid
    'ClientSecret'       = $AppSecret.SecretText 
    'ThumbPrint'         = $cer.Thumbprint
    'TenantName'         = $tenantName
    'ServicePrincipal'   = $ServicePrincipal
}


# exporting details to xml
$detailsFile = "." + $switch + "security" + $switch + "connectionDetails.xml"
try {
    $ConnectionDetails | Export-Clixml -path $detailsFile    
    Log-Message -file $file -Status Done -Message "Exporting connection details to $DetailsFile"
}
catch {
    $connectionError = $_
    Log-Message -file $file -Status error -Message "$($connectionError.Exception.Message)"
    break
}

Disconnect-AZAccount | out-Null
Log-Message -file $file -Status INFO -Message "Disconnecting Azure"
#endregion appregistration

$psdfile = "." + $switch + "import-data" + $switch + "sharepoint-information.psd1"
If (!(Test-path $psdfile)){
    Log-Message -file $file -Status WARNING -Message "No SharePoint information file is found. Please provide the correct answers in the next questions"
    $SPOUsername = Read-Host "[QUESTION] - What is the SharePoint Admin username?"
        $SPOPassword = Read-Host "[QUESTION] - What is the SharePoint Admin password?" -MaskInput
        $SPOadminUrl = Read-Host "[QUESTION] - What is the SharePoint Admin url?"
        $CompanyName = Read-Host "[QUESTION] - What is the CompanyName?"
        $Teams = Read-Host "[QUESTION] - Will there be a teams for the migration data?"
        $Weblogin = Read-Host "[QUESTION] - Does the migration account have MFA enabled?"

        if ($Weblogin -eq "Y") {
            $Weblogin = $true
        }
        Else {
            $Weblogin = $false
        }
        if ($teams -eq "Y") {
            $teams = $true
        }
        Else {
            $teams = $false
        }
    
        Log-Message -file $file -Status INFO -Message "Exporting information to $psdfile ... " -NoNewline
        $hashString = @{
            Username      = $SPOUsername
            Password      = $SPOPassword
            SPOUrl        = $SPOadminUrl
            CompanyName   = $CompanyName
            Teams         = $teams
            Weblogin      = $weblogin
        }
        ConvertTo-Psd -InputObject $hashString -Depth 1 | out-file $psdfile

}
else {
    Log-Message -file $file -Status WARNING -Message "SharePoint information file is found." 
    $question = Read-Host "[QUESTION] - Would you like to replace the file? (Y/N)"
    If ($question -ne "Y") {
        Log-Message -file $file -Status INFO -Message "$psdFile will not be replaced. Exit script."
    }
    If ($question -eq "Y") {
        Log-Message -file $file -Status WARNING -Message "Loading questions for SharePoint information ..."
        $SPOUsername = Read-Host "[QUESTION] - What is the SharePoint Admin username?"
        $SPOPassword = Read-Host "[QUESTION] - What is the SharePoint Admin password?" -MaskInput
        $SPOadminUrl = Read-Host "[QUESTION] - What is the SharePoint Admin url?"
        $CompanyName = Read-Host "[QUESTION] - What is the CompanyName?"
        $Teams = Read-Host "[QUESTION] - Will there be a teams for the migration data?"
        $Weblogin = Read-Host "[QUESTION] - Does the migration account have MFA enabled?"

        if ($Weblogin -eq "Y") {
            $Weblogin = $true
        }
        Else {
            $Weblogin = $false
        }
        if ($teams -eq "Yes" -or $teams -eq "y") {
            $teams = $true
        }
        Else {
            $teams = $false
        }
    
        Log-Message -file $file -Status INFO -Message "Exporting information to $psdfile ... " -NoNewline
        $hashString = @{
            Username      = $SPOUsername
            Password      = $SPOPassword
            SPOUrl        = $SPOadminUrl
            CompanyName   = $CompanyName
            Teams         = $teams
            Weblogin      = $weblogin
        }
        ConvertTo-Psd -InputObject $hashString -Depth 1 | out-file $psdfile
        
    }
}

Log-Message -file $file "----------------------------------------------------------------"
Log-Message -file $file  "END installation app registration $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------------------------"
Log-Message -file $file  " "