# 
param(
    [Parameter()]
    [string]$InstallFolder
)

function Show-Message([String]$Message, [string]$Status) {
    $Status = $Status.ToUpper()
 
    if (!($Status)) {
        Write-Host "$message"
    }
    else {
        # to console
        Write-Host "[" -NoNewline
        switch ($status) {
            INFO { Write-host "INFO" -NoNewline }
            WARNING { Write-host "WARNING" -NoNewline -ForegroundColor Yellow }
            WAITING { Write-host "WAITING" -NoNewline -ForegroundColor Yellow }
            SKIPPING { Write-host "SKIPPING" -NoNewline -ForegroundColor Yellow }
            UPDATING { Write-host "UPDATING" -NoNewline -ForegroundColor Yellow }
            INSTALL { Write-host "INSTALL" -NoNewline -ForegroundColor Yellow }
            ERROR { Write-host "ERROR" -NoNewline -ForegroundColor RED }
            ADDED { Write-host "ADDED" -NoNewline }
            DONE { Write-Host "DONE" -ForegroundColor green -NoNewLine }
            FINISHED { Write-Host "FINISHED" -ForegroundColor green -NoNewLine }
        }
        Write-Host "] - $message"
    }
}
function Disable-InternetExplorerESC {
    $AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
    Set-ItemProperty -Path $AdminKey -Name "IsInstalled" -Value 0
    Stop-Process -Name Explorer
    Write-Host "IE Enhanced Security Configuration (ESC) has been disabled." -ForegroundColor Green
}

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
            [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Show-Message -Status ERROR -Message "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
    Break
}

function Test-LicenseKey {
    $licenseKey = Get-RCLicenseKey
    $apiUrl = "http://api.checkyourlic.org:80/checklicense"

    try {
        $body = @{
            licenseKey = $licenseKey
        } | ConvertTo-Json

        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Body $body -ContentType "application/json" #-ErrorAction Stop
    }
    catch {
        $oops = $_.ErrorDetails.Message | ConvertFrom-Json
        if (-not $Oops.Message -eq "License key is valid") {
            write-Host "License validation failed" -ForegroundColor Red
            return $false
        }
        if ($Oops.Message -eq "License key is valid") {
            return $true
        }
        if ($Oops.Message -eq "Invalid license key") {
            write-Host "License validation failed" -ForegroundColor Red
            return $false
        }	
        if ($Oops.Message -eq "License key has expired") {
            write-Host "License validation failed. Key is expired" -ForegroundColor Red
            return $false
        }
    }
    if ($response.Message -eq "License key is valid") {
        return $true
    }
    else {
        return $false
    }	
}

function Get-RCLicenseKey {
    [CmdletBinding()]
    param()
    
    try {
        $registryPath = "HKLM:\SOFTWARE\RAPIDCIRCLE"
        
        # Check if registry path exists
        if (-not (Test-Path -Path $registryPath)) {
            Write-Warning "Registry path not found: $registryPath"
            return $null
        }

        # Get license key value
        $licenseKey = Get-ItemProperty -Path $registryPath -Name "licenseKey" -ErrorAction SilentlyContinue
        
        if ($null -eq $licenseKey) {
            Write-Warning "License key not found in registry"
            return $null
        }

        return $licenseKey.licenseKey
    }
    catch {
        Write-Error "Failed to read registry: $_"
        return $null
    }
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Disable-InternetExplorerESC
clear

$password = Read-Host "Please enter password to unzip package" -AsSecureString

if (!($password)) {
    Show-Message -Status ERROR -Message "A password needs to be given to install the Migration street"
    break
}

# check if Edge is installed

$check = (Get-Item "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -ErrorAction SilentlyContinue).VersionInfo 
if (!($Check)) {
    Show-Message -Status ERROR -Message "Microsoft Edge is not installed."

    md -Path $env:temp\edgeinstall -erroraction SilentlyContinue | Out-Null
    $Download = join-path $env:temp\edgeinstall MicrosoftEdgeEnterpriseX64.msi

    Invoke-WebRequest 'https://msedge.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/8ec28e1e-d2ae-4d26-b1e6-324aa5318db1/MicrosoftEdgeEnterpriseX64.msi'  -OutFile $Download
    Show-Message -Status INSTALL -Message "Installing Microsoft Edge now .."
    Start-Process "$Download" -ArgumentList "/quiet" -wait

    $check = (Get-Item "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -ErrorAction SilentlyContinue).VersionInfo 
    if (!($Check)) {
        Show-Message -Status Error -Message "Microsoft Edge is still not installed. Please install manually"
        break
    }

    Show-Message -Status DONE -Message "Microsoft Edge is installed ..."

}
# check for 7ZIP Powershell Module
$NuGet = Get-PackageProvider -Name "NuGet" -ErrorAction SilentlyContinue
if (!($NuGet)) {
    Try {
        Show-Message -status WARNING -Message "NuGet Package provider not installed. Installing now ..."
        Install-PackageProvider -Name NuGet -Force
    }
    Catch {
        $Fault = $_
        Show-Message -Status ERROR -Message "Cannot install NuGet PackageProvider"
        break
    }
}

try {
    Import-Module -name 7Zip4Powershell -ErrorAction Stop
    Show-Message -Status DONE -Message "7Zip4Powershell is imported"
}
Catch {
    $fault = $_
    Show-Message -Status Error -Message "7Zip4Powershell is not installed..."
    Show-Message -Status Warning -Message "Installing 7Zip4Powershell now ..."

    Install-Module -name 7Zip4Powershell -Force

    try {
        Import-Module -Name 7Zip4Powershell -ErrorAction Stop
        Show-Message -Status DONE -Message "7Zip4Powershell is imported"
    }
    catch {
        $fault = $_
        Show-Message -Status Error -Message "$($fault.Exception.Message)"
        Show-Message -Status Error -Message "7Zip4Powershell is not installed and could not be installed... aborting ..."
        break
    }
}

try {
    Import-Module -name CredentialManager -ErrorAction Stop
    Show-Message -Status DONE -Message "CredentialManager is imported"
}
Catch {
    $fault = $_
    Show-Message -Status Error -Message "CredentialManager is not installed..."
    Show-Message -Status Warning -Message "Installing CredentialManager now ..."

    Install-Module -name CredentialManager -Force

    try {
        Import-Module -Name CredentialManager -ErrorAction Stop
        Show-Message -Status DONE -Message "CredentialManager is imported"
    }
    catch {
        $fault = $_
        Show-Message -Status Error -Message "$($fault.Exception.Message)"
        Show-Message -Status Error -Message "CredentialManager is not installed and could not be installed... aborting ..."
        break
    }
}

#region begin

if (!($Installfolder)) {
    $InstallFolder = "C:\RCScripts"
}

Show-Message -message "----------------------------------------------------------"
Show-Message -message "Installation Rapid Circle Migration Street - $installfolder"
Show-Message -message "----------------------------------------------------------"


$repo = "rcalexterneuzen/rc-migration-street"
$filename = "rc-migration-street.zip"
$releases = "https://api.github.com/repos/$repo/releases"

if (!(Test-Path $InstallFolder)) {
    Try {
        $create = New-Item -Path $InstallFolder -ItemType Directory
        Show-Message -Status DONE -Message "Created location for Rapid Circle Migration Street: $installfolder"
        $New = $true
    }
    Catch {
        $CreateError = $_
        Show-Message -Status ERROR -Message "Cannot create location for Rapid Circle Migration Street"
        Show-Message -Status ERROR -Message "$($webError.Exception.Message)"
        break
    }
}
else {
    Show-Message -Status DONE -Message "Folder already exists: $installfolder"
    $inplace = $true
}

try {
    $tag = (Invoke-WebRequest $releases | ConvertFrom-Json)[0].tag_name
}
catch {
    $webError = $_
    Show-Message -Status ERROR -Message "Error getting the latest release from Github $repo"
    Show-Message -Status ERROR -Message "$($webError.Exception.Message)"
    break
}

if ($new) {
    Show-Message -Status UPDATING -Message "Rapid Circle Migration Street version $tag will be installed"

    Try {
        $download = "https://github.com/$repo/releases/download/$tag/$filename"
        $name = $filename.Split(".")[0]
        $zip = "$name-$tag.zip"
        $dir = "$name-$tag"
        Show-Message -Status Updating -Message "Updating to version $tag : $zip will be downloaded"
    }
    Catch {
        $webError = $_
        Show-Message -Status ERROR -Message "Error getting the latest release: $($weberror.Exception.Message)"
        break
    }

    # download zip package
    Try {
        $location = $InstallFolder + "\" + $zip
        $download = Invoke-WebRequest $download -Out $location
        Show-Message -Status DONE -Message "- Downloaded $zip to $location"
    }
    Catch {
        $webError = $_
        Show-Message -Status ERROR -Message "Error getting the download: $($weberror.Exception.Message)"
        break
    }

    Try {
        $extract = Expand-7Zip $location -TargetPath $InstallFolder -SecurePassword $password -ErrorAction SilentlyContinue
        #$extract = Expand-Archive $location -Destination $InstallFolder -Force
        Show-Message -Status DONE -Message "- Extracted $location to $InstallFolder"
    }
    Catch {
        $webError = $_
        Show-Message -Status ERROR -Message "Error extracting $location to $installfolder : $($weberror.Exception.Message)"
        break
    }

    Try {
        $remove = Remove-Item $location -Recurse -force -ErrorAction SilentlyContinue
        Show-Message -Status DONE -Message "- $location removed from $InstallFolder"
    }
    Catch {
        $webError = $_
        Show-Message -Status ERROR -Message "- Cannot remove $location from $InstallFolder : $($weberror.Exception.Message)"
        break
    }

    Try {
        $toRemove = $InstallFolder + "\" + "__MACOSX"
        $remove = Remove-Item $toRemove -Recurse -force -ErrorAction SilentlyContinue
    }
    Catch {
        $webError = $_
        Show-Message -Status ERROR -Message "- - Cannot remove $ToRemove from $InstallFolder : $($weberror.Exception.Message)"
        break
    }

    $files = Get-ChildItem -Path $InstallFolder -file -Recurse
    $files = $files | Where-Object { $_.Name -eq ".DS_store" }

    if ($files.count -eq 0) {
        Show-Message -Status DONE -Message "- Rapid Circle Migration Street is downloaded and extracted"
    }
    else {
        ForEach ($File in $files) {
            try {
                Remove-Item $File.Fullname -force
            }
            catch {
                $delError = $_
                Write-Host "$($DelError.Exception.Message)"
            }
        }
        Show-Message -Status DONE -Message "- Cleanup of Rapid Circle Migration Street is done"
    }

    $files = Get-ChildItem -Path "$($InstallFolder)\Modules" -file -Recurse
    $modules = $files | Where-Object { $_.Extension -eq ".psm1" }

    if (!(Test-path -Path "C:\Program Files\WindowsPowershell\Modules\RapidCircle")){
        New-Item -Path "C:\Program Files\WindowsPowershell\Modules"  -Name "RapidCircle" -ItemType "Directory"
        }
        
    if ($files.count -eq 0) {
        Show-Message -Status DONE -Message "- NO modules found to copy"
    }
    else {
        ForEach ($File in $modules) {
            try {
                Move-Item $File.Fullname -Destination "C:\Program Files\WindowsPowerShell\Modules\RapidCircle" -force
            }
            catch {
                $delError = $_
                Write-Host "$($DelError.Exception.Message)"
            }
        }
        Show-Message -Status DONE -Message "- Cleanup of Rapid Circle Migration Street is done"
    }
    
    Write-Host "[" -NoNewline
    Write-Host "QUESTION" -ForegroundColor Yellow -NoNewline
    Write-Host "] - Would you like to start the installation of the Migration Street? (Y/N)" -NoNewline
    $Question = Read-Host " "

    if ($question -ne "Y") {
        Show-Message -Status DONE -Message "Rapid Circle Migration Street is downloaded and unzipped to $InstallFolder"
        break
    }
    else {
        CD $InstallFolder

        Write-Host "[" -NoNewline
        Write-Host "WAITING" -NoNewline -ForegroundColor YELLOW
        Write-Host "] - Rapid Circle Migration Street requirements are being installed" -NoNewline
        # starting installation
        $requirements = $InstallFolder + "\install-requirements.ps1"
        Start-Process Powershell -verb runAs -ArgumentList "-file $requirements" -wait
        Write-host "    DONE" -ForegroundColor Green

        Write-Host "[" -NoNewline
        Write-Host "WAITING" -NoNewline -ForegroundColor YELLOW
        Write-Host "] - Rapid Circle Migration Street app registration is being installed" -NoNewline
        $appregistration = $InstallFolder + "\install-app-registration.ps1"
        Start-Process Powershell -verb runAs -ArgumentList "-file $appregistration" -wait
        Write-host "    DONE" -ForegroundColor Green

        $content = Import-PowerShellDataFile ".\import-data\sharepoint-information.psd1"
        if ($Content.teams -eq $true) {

            Write-Host "[" -NoNewline
            Write-Host "WAITING" -NoNewline -ForegroundColor YELLOW
            Write-Host "] - Rapid Circle Migration Street Teams is being installed" -NoNewline
            $teamsinstall = $InstallFolder + "\install-Migrationteams.ps1"
            Start-Process Powershell -verb runAs -ArgumentList "-file $teamsinstall" -wait
            Write-host "    DONE" -ForegroundColor Green
        }
    }
}
if ($Inplace) {
    $version = $Tag.Split("v")[1]
    try {
        $info = Get-Content "$InstallFolder\version.json" -ErrorAction stop | ConvertFrom-Json
    }
    catch {
        $weberror = $_
        Show-Message -Status ERROR -Message "There was an error getting version.json : $($Weberror.Exception.Message)"
        break
    }

    if ($info.Version -eq $version) {
        Show-Message -Status FINISHED -Message "Rapid Circle Migration Street is up-to-date"
        break
    }
    else {
        Show-Message -Status UPDATING -Message "Rapid Circle Migration Street will be updated to version $($version)"

        # creating backup
        if (!(Test-Path $InstallFolder\Backup)) {
            Try {
                $create = New-Item -Path "$InstallFolder\Backup" -ItemType Directory
                Show-Message -Status DONE -Message "- Creating backup folder for Rapid Circle Migration Street"

            }
            Catch {
                $CreateError = $_
                Show-Message -Status ERROR -Message "Cannot create backup folder for Rapid Circle Migration Street"
                Show-Message -Status ERROR -Message "$($createError.Exception.Message)"
                break
            }
        }
        else {
            $archive = $InstallFolder + "\Backup\$(Get-Date -Format "ddMMyyyy")-Backup-Version-$($info.version).zip"
            $Exclude = "BACKUP"
            $files = Get-ChildItem -Path $installFolder -Exclude $Exclude
            Compress-Archive -path $files -DestinationPath $archive -CompressionLevel Fastest -force 
            Show-Message -Status DONE -Message "- Created a backup: $($archive)"
        }
        Try {
            $download = "https://github.com/$repo/releases/download/$tag/$filename"
            $name = $filename.Split(".")[0]
            $zip = "$name-$tag.zip"
            $dir = "$name-$tag"
            Show-Message -Status Updating -Message "Updating to version $version : $($zip) will be downloaded"
        }
        Catch {
            $webError = $_
            Show-Message -Status ERROR -Message "Error getting the latest release: $($weberror.Exception.Message)"
            break
        }

        # download zip package
        Try {
  
            if (!(Test-Path "$($env:temp)\RCStreet")) {
                $out = New-Item -Name "RCStreet" -Path $Env:temp -ItemType Directory -ErrorAction Stop
            }

            $location = $env:temp + "\RCStreet\" + $zip
            $download = Invoke-WebRequest $download -Out $location
            Show-Message -Status DONE -Message "- Downloaded $zip to $location"
        }
        Catch {
            $webError = $_
            Show-Message -Status ERROR -Message "Error getting the download: $($weberror.Exception.Message)"
        }

        Try {
            $extract = Expand-7Zip $location -TargetPath $InstallFolder -SecurePassword $password -ErrorAction SilentlyContinue
            Show-Message -Status DONE -Message "- Extracted $location to $InstallFolder"
        }
        Catch {
            $webError = $_
            Show-Message -Status ERROR -Message "Error extracting $location to $installfolder : $($weberror.Exception.Message)"
            break
        }

        Try {
            $remove = Remove-Item $location -Recurse -force -ErrorAction SilentlyContinue
            Show-Message -Status DONE -Message "- $location removed from $InstallFolder"
        }
        Catch {
            $webError = $_
            Show-Message -Status ERROR -Message "- Cannot remove $location from $InstallFolder : $($weberror.Exception.Message)"
            break
        }

        Try {
            $toRemove = $InstallFolder + "\" + "__MACOSX"
            $remove = Remove-Item $toRemove -Recurse -force -ErrorAction SilentlyContinue
        }
        Catch {
            $webError = $_
            Show-Message -Status ERROR -Message "- - Cannot remove $ToRemove from $InstallFolder : $($weberror.Exception.Message)"
            break
        }

        $files = Get-ChildItem -Path $InstallFolder -file -Recurse
        $files = $files | Where-Object { $_.Name -eq ".DS_store" }

        if ($files.count -eq 0) {
            Show-Message -Status DONE -Message "- Rapid Circle Migration Street is downloaded and extracted"
        }
        else {
            ForEach ($File in $files) {
                try {
                    Remove-Item $File.Fullname -force
                }
                catch {
                    $delError = $_
                    Write-Host "$($DelError.Exception.Message)"
                }
            }
            Show-Message -Status DONE -Message "- Cleanup of Rapid Circle Migration Street is done"
        }
    $files = Get-ChildItem -Path "$($InstallFolder)\Modules" -file -Recurse
    $modules = $files | Where-Object { $_.Extension -eq ".psm1" }

    if (!(Test-path -Path "C:\Program Files\WindowsPowershell\Modules\RapidCircle")){
        New-Item -Path "C:\Program Files\WindowsPowershell\Modules"  -Name "RapidCircle" -ItemType "Directory"
        }
        
    if ($files.count -eq 0) {
        Show-Message -Status DONE -Message "- NO modules found to copy"
    }
    else {
        ForEach ($File in $modules) {
            try {
                Move-Item $File.Fullname -Destination "C:\Program Files\WindowsPowerShell\Modules\RapidCircle" -force
            }
            catch {
                $delError = $_
                Write-Host "$($DelError.Exception.Message)"
            }
        }
        Show-Message -Status DONE -Message "- Cleanup of Rapid Circle Migration Street is done"
    }

        Show-Message -Status DONE -Message "Rapid Circle Migration Street is downloaded and unzipped to $InstallFolder"

    }
}


#endregion begin
