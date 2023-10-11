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
## check if powershell elevated is started
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
            [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
    Break
}

$Global:file = "." + $switch + "log" + $switch + "$(Get-Date -format "ddMMyyy-HHmm")-Install-Requirements.txt"

# start provisioning teams
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "Start script for Installing Requirements $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "

## import custom modules

Try {
    $Modules = (Import-PowerShellDataFile ./import-data/required-modules.psd1).modules
    Log-Message -file $file -Status Done -Message "Getting modules to install"
}
Catch {
    $moduleError = $_
    Log-Message -file $file -Status error -Message "$($moduleError.Exception.Message)"
    break
}


Try {
    $Import = Import-Module ./modules/rc-required-modules.psm1
    Log-Message -file $file -Status Done -Message "Imported the install module"
}
Catch {
    $moduleError = $_
    Log-Message -file $file -Status error -Message "$($moduleError.Exception.Message)"
    break
}

If (!(Get-PackageProvider | Where-Object {$_.Name -eq "NuGet"})){
    Try {
        Install-PackageProvider -Name NuGet -Force
        Log-Message -file $file -Status Done -Message "Installing Nuget Package provider"
    }
    Catch {
        $moduleError = $_
        Log-Message -file $file -Status error -Message "$($moduleError.Exception.Message)"
        break
    }
}

Install-requirements -Modules $Modules -File $File

$OutputFolder = ".\downloads"
if (Test-Path -path $OutputFolder) {
    Log-Message -file $file -Status Done -Message "Download folder is found"
}
else {
    New-Item -Path $OutputFolder -ItemType Directory | Out-Null
    Log-Message -file $file -Status warning -Message "Downloads folder is created"
}

# # SharePoint Online Client Components SDK
# $FileToDownload = ".\downloads\sharepointclientcomponents.msi"
# if (Test-Path -path $FileToDownload) {
#     Write-Host "[" -NoNewline
#     Write-Host "DONE" -NoNewline -ForegroundColor Green
#     Write-Host "] - - Sharepoint Online Client SDK is found"
#     Log-Message -file $file "[DONE] - - Sharepoint Online Client SDK is found"
# }
# else {
#     Try {
#         Invoke-WebRequest -Uri "https://download.microsoft.com/download/B/3/D/B3DA6839-B852-41B3-A9DF-0AFA926242F2/sharepointclientcomponents_16-6906-1200_x64-en-us.msi" -OutFile "$OutputFolder\sharepointclientcomponents.msi"
#         Write-Host "[" -NoNewline
#         Write-Host "DOWNLOADED" -ForegroundColor yellow -NoNewline
#         Write-Host "] - Sharepoint Online Client SDK is downloaded"
#         Log-Message -file $file "[DOWNLOADED] - Sharepoint Online Client SDK is downloaded"
#         Log-Message -File $file "[DONWLOADED] - File can be found in: $($OutputFolder)\sharepointclientcomponents.msi"
#         $SDK = $true
#     }
#     catch {
#         $ErrorToDisplay = $_
#         Write-Host "[" -NoNewline
#         Write-Host "ERROR" -ForegroundColor Red -NoNewline
#         Write-Host "] - " -NoNewline
#         Write-Host "$($ErrorToDisplay.Exception.Message)" -ForegroundColor red
#         Log-Message -file $file -Message "[ERROR] - $($ErrorToDisplay.Exception.Message)"
#         $SDK = $false
#     }
# }   

$SPMTCheck = "$env:UserProfile\Documents\WindowsPowerShell\Modules\Microsoft.SharePoint.MigrationTool.PowerShell"
if (Test-Path $SPMTCheck) {
    Log-Message -file $file -Status Done -Message "Sharepoint Migration Tool is found"
    $SPMT = $false
}
else {
    Log-Message -file $file -Status WARNING -Message "Sharepoint Migration Tool is not found"

    # redirect to download

    $Downloadedfile = ".\downloads\spmtsetup.exe"
    if (Test-Path $Downloadedfile) {
        Log-Message -file $file -Status Done -Message "Sharepoint Migration Tool installation file is found"
        $SPMT = $true
    }
    else {
    
        Try {
            Invoke-WebRequest -Uri "https://spmt.sharepointonline.com/spmtinstaller/spmtbuild/ga/spmtsetup.exe" -OutFile "$OutputFolder\spmtsetup.exe"
            Log-Message -file $file -Status Done -Message "Sharepoint Migration Tool installation file is downloaded"
            Log-Message -file $file -Status Done -Message "File can be found in: $($OutputFolder)\spmtsetup.exe"
            $SPMT = $true
        }
        catch {
            $ErrorToDisplay = $_
            Log-Message -file $file -Status error -Message "$($ErrorToDisplay.Exception.Message)"
            $SPMT = $false
        }
    }
}

# checking vc_redist.x64
$VCREDIST = "registry::HKEY_CLASSES_ROOT\Installer\Dependencies\VC,redist.x64,amd64,14.36,bundle"
if (Test-Path $VCREDIST) {
    Log-Message -file $file -Status Done -Message "Microsoft Visual C++ Redistributable is present"
    $VCRE = $false
}
else {
    Log-Message -file $file -Status WARNING -Message "Microsoft Visual C++ Redistributable is not found"

    # redirect to download

    $Downloadedfile = ".\downloads\vcredist_x64.exe"
    if (Test-Path $Downloadedfile) {
        Log-Message -file $file -Status Done -Message "Microsoft Visual C++ Redistributable installation file is found"
        $VCRE = $true
    }
    else {
    
        Try {
            Invoke-WebRequest -Uri "https://aka.ms/vs/17/release/vc_redist.x64.exe" -OutFile "$OutputFolder\vcredist_x64.exe"
            Log-Message -file $file -Status Done -Message "Microsoft Visual C++ Redistributable installation file is downloaded"
            Log-Message -file $file -Status Done -Message "File can be found in: $($OutputFolder)\vcredist_x64.exe"
            $VCRE = $true
        }
        catch {
            $ErrorToDisplay = $_
            Log-Message -file $file -Status error -Message "$($ErrorToDisplay.Exception.Message)"
            $VCRE = $false
        }
    }
}

#checking .net framework version

if ((Get-ItemPropertyValue -LiteralPath 'HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release) -ge 393297){
    Log-Message -file $file -Status Done -Message ".NET Framework is 4.6 or higher"
    $framework = $false
}
else {
    Try {
        Invoke-Webrequest -Uri "https://download.visualstudio.microsoft.com/download/pr/014120d7-d689-4305-befd-3cb711108212/0fd66638cde16859462a6243a4629a50/ndp48-x86-x64-allos-enu.exe" -OutFile "$OutputFolder\FrameWork48.exe"
        Log-Message -file $file -Status Done -Message ".NET Framework 4.8 installation file is downloaded"
        Log-Message -file $file -Status Done -Message "File can be found in: $($OutputFolder)\Framework48.exe"
        $Framework = $true
    }
    catch {
        $ErrorToDisplay = $_
        Log-Message -file $file -Status error -Message "$($ErrorToDisplay.Exception.Message)"
        $Framework = $false
    }
}

Log-Message -File $file -Message ""

# starting installation
# If ($sdk){
#     try{
#         $installfile = ((get-ChildItem $outputfolder) | Where-Object {$_.Name -eq "sharepointclientcomponents.msi"}).fullname
#         Start-Process msiexec "/i $installfile /qn" -Wait
#         Write-Host "[" -NoNewline
#         Write-Host "INSTALLED" -ForegroundColor green -NoNewline
#         Write-Host "] - Sharepoint Online Client SDK is installed"
#         Log-Message -file $file "[INSTALLED] - Sharepoint Online Client SDK is installed"

#     }
#     catch {
#         $installError = $_
#         Write-Host "[" -NoNewline
#         Write-Host "ERROR" -ForegroundColor red -NoNewline
#         Write-host "] - Sharepoint Online Client SDK cannot be installed with error: $($Installerror.Exception.Message)"
#         Log-Message -File $file -Message "[ERROR] - Sharepoint Online Client SDK cannot be installed with error: $($Installerror.Exception.Message)"
#     }
    
# }
# else {
#     Write-Host "[" -NoNewline
#     Write-Host "ERROR" -ForegroundColor red -NoNewline
#     Write-host "] - Sharepoint Online Client SDK cannot be installed"
#     Log-Message -File $file -Message "[ERROR] - Sharepoint Online Client SDK cannot be installed"
# }

if ($VCRE){
    try{
        $installfile = ((get-ChildItem $outputfolder) | Where-Object {$_.Name -eq "vcredist_x64.exe"}).fullname
        Start-Process $installfile -ArgumentList "/q"  -Wait
        Log-Message -file $file -Status Done -Message "Microsoft Visual C++ Redistributable is installed"

    }
    catch {
        $installError = $_
        Log-Message -file $file -Status error -Message "Microsoft Visual C++ Redistributable cannot be installed with error: $($Installerror.Exception.Message)"
    }
}
else {
        Log-Message -file $file -Status error -Message "Microsoft Visual C++ Redistributable cannot be installed."
}


if ($SPMT){
    try{
        $installfile = ((get-ChildItem $outputfolder) | Where-Object {$_.Name -eq "spmtsetup.exe"}).fullname
        Start-Process $installfile -Wait
        Log-Message -file $file -Status Done -Message "Sharepoint migration tool is installed"

    }
    catch {
        $installError = $_
        Log-Message -file $file -Status error -Message "Sharepoint migration tool cannot be installed with error: $($Installerror.Exception.Message)"
    }
}
else {
    Log-Message -file $file -Status error -Message "Sharepoint migration tool cannot be installed"
}

if ($Framework){
    try{
        $installfile = ((get-ChildItem $outputfolder) | Where-Object {$_.Name -eq "framework48.exe"}).fullname
        Start-Process msiexec "/i $installfile /qn" -Wait
        Log-Message -file $file -Status Done -Message ".NET Framework 4.8 is installed"

    }
    catch {
        $installError = $_
        Log-Message -file $file -Status error -Message ".NET Framework 4.8 cannot be installed with error: $($Installerror.Exception.Message)"
    }
}
else {
    Log-Message -file $file -Status skipping -Message ".NET Framework 4.6 or higher is allready installed"
}

Log-Message -File $file ""
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "End installation requirements $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "
