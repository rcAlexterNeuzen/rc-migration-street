# 
param(
    [Parameter(Mandatory)]
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
            ERROR { Write-host "ERROR" -NoNewline -ForegroundColor RED }
            ADDED { Write-host "ADDED" -NoNewline }
            DONE { Write-Host "DONE" -ForegroundColor green -NoNewLine }
            FINISHED { Write-Host "FINISHED" -ForegroundColor green -NoNewLine }
        }
        Write-Host "] - $message"
    }
}

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
            [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Show-Message -Status ERROR -Message "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
    Break
}

clear

Show-Message -message "----------------------------------------------------------"
Show-Message -message "Installation Rapid Circle Migration Street - $installfolder"
Show-Message -message "----------------------------------------------------------"

#region begin
$InstallFolder = "C:\RCScripts1"
$repo = "rcalexterneuzen/rc-migration-street"
$filename = "rc-migration-street.zip"
$releases = "https://api.github.com/repos/$repo/releases"

if (!(Test-Path $InstallFolder)) {
    Try {
        $create = New-Item -Path $InstallFolder -ItemType Directory
        Show-Message -Status DONE -Message "Created location for Rapid Circle Migration Street: $installfolder"
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
        Show-Message -Status UPDATING -Message "Rapid Circle Migration Street will be updated to version $version"

        # creating backup
        if (!(Test-Path $InstallFolder\Backup)){
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
            Show-Message -Status DONE -Message "- Created a backup: $archive"
        }
        Try {
            $download = "https://github.com/$repo/releases/download/$tag/$filename"
            $name = $filename.Split(".")[0]
            $zip = "$name-$tag.zip"
            $dir = "$name-$tag"
            Show-Message -Status Updating -Message "Updating to version $version : $zip will be downloaded"
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
        }

        Try {
            $extract = Expand-Archive $location -Destination $InstallFolder -Force
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
}



#endregion begin