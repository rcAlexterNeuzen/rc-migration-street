# 

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
            [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
    Break
}

$targetLocation = "C:\RCScripts1"
$repo = "rcalexterneuzen/rc-migration-street"
$filename = "rc-migration-street.zip"
$releases = "https://api.github.com/repos/$repo/releases"

if (!(Test-Path $targetLocation)) {
    Try {
        $create = New-Item -Path $targetLocation -ItemType Directory
        Write-Host "[" -NoNewline
        Write-Host "DONE" -NoNewline -ForegroundColor Green
        Write-Host "] - Creating location for Rapid Circle Migration Street "
    }
    Catch {
        $CreateError = $_
        Write-Host "[" -NoNewline
        Write-Host "ERROR" -NoNewline -ForegroundColor RED
        Write-Host "] - Cannot create location for Rapid Circle Migration Street " -nonewLine
        Write-Host "$($webError.Exception.Message)" -ForegroundColor Red
        break
    }
}
else {
    Write-Host "[" -NoNewline
    Write-Host "DONE" -NoNewline -ForegroundColor Green
    Write-Host "] - Location for Rapid Circle Migration Street is in place"
    $inplace = $true
}

$tag = (Invoke-WebRequest $releases | ConvertFrom-Json)[0].tag_name

if ($Inplace) {
    $version = $Tag.Split("v")[1]
    $info = Get-Content $targetLocation\version.json -ErrorAction SilentlyContinue | ConvertFrom-Json

    if ($info.Version -eq $version) {
        Write-Host "[" -NoNewline
        Write-Host "DONE" -NoNewline -ForegroundColor Green
        Write-Host "] - Rapid Circle Migration Street is up-to-date"
        break
    }
    else {
        Write-Host "[" -NoNewline
        Write-Host "UPDATING" -NoNewline -ForegroundColor Yellow
        Write-Host "] - Rapid Circle Migration Street will be updated to version $version"
        Try {
            Write-Host "[" -NoNewline
            Write-Host "DONE" -NoNewline -ForegroundColor Green
            Write-Host "] - - Found releases: " -nonewLine
            $download = "https://github.com/$repo/releases/download/$tag/$filename"
            $name = $filename.Split(".")[0]
            $zip = "$name-$tag.zip"
            $dir = "$name-$tag"
            Write-Host "$zip" -ForegroundColor Green -NoNewline
            Write-Host " will be downloaded"
        }
        Catch {
            $webError = $_
            Write-Host "[" -NoNewline
            Write-Host "ERROR" -NoNewline -ForegroundColor Red
            Write-Host "] - - Error getting the latest release: " -nonewLine
            Write-Host "$($webError.Exception.Message)" -ForegroundColor Red
        }

        # download zip package
        Try {
            $location = $targetLocation + "\" + $zip
            $download = Invoke-WebRequest $download -Out $location
            Write-Host "[" -NoNewline
            Write-Host "DONE" -NoNewline -ForegroundColor Green
            Write-Host "] - - Downloaded $zip to $location"

        }
        Catch {
            $webError = $_
            Write-Host "[" -NoNewline
            Write-Host "ERROR" -NoNewline -ForegroundColor Red
            Write-Host "] - - Error getting the download: " 
            Write-Host "$($webError.Exception.Message)" -ForegroundColor Red
        }

        Try {
            $extract = Expand-Archive $location -Destination $targetLocation -Force
            Write-Host "[" -NoNewline
            Write-Host "DONE" -NoNewline -ForegroundColor Green
            Write-Host "] - - Extracted $location to $targetLocation"

        }
        Catch {
            $webError = $_
            Write-Host "[" -NoNewline
            Write-Host "ERROR" -NoNewline -ForegroundColor Red
            Write-Host "] - - Error getting the download: " 
            Write-Host "$($webError.Exception.Message)" -ForegroundColor Red
            break
        }

        Try {
            $remove = Remove-Item $location -Recurse -force -ErrorAction SilentlyContinue
            Write-Host "[" -NoNewline
            Write-Host "DONE" -NoNewline -ForegroundColor Green
            Write-Host "] - - $location removed from $targetLocation"

        }
        Catch {
            $webError = $_
            Write-Host "[" -NoNewline
            Write-Host "ERROR" -NoNewline -ForegroundColor Red
            Write-Host "] - - Cannot remove $location from $targetLocation : " 
            Write-Host "$($webError.Exception.Message)" -ForegroundColor Red
            break
        }

        Try {
            $toRemove = $targetLocation + "\" + "__MACOSX"
            $remove = Remove-Item $toRemove -Recurse -force -ErrorAction SilentlyContinue
            Write-Host "[" -NoNewline
            Write-Host "DONE" -NoNewline -ForegroundColor Green
            Write-Host "] - - $ToRemove is removed from $targetLocation"

        }
        Catch {
            $webError = $_
            Write-Host "[" -NoNewline
            Write-Host "ERROR" -NoNewline -ForegroundColor Red
            Write-Host "] - - Cannot remove $ToRemove from $targetLocation : " 
            Write-Host "$($webError.Exception.Message)" -ForegroundColor Red
            break
        }

        $files = Get-ChildItem -Path $targetLocation -file -Recurse
        $files = $files | Where-Object { $_.Name -eq ".DS_store" }

        if ($files.count -eq 0) {
            Write-Host "[" -NoNewline
            Write-Host "DONE" -NoNewline -ForegroundColor Green
            Write-Host "] - Rapid Circle Migration Street is downloaded"
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
            Write-Host "[" -NoNewline
            Write-Host "DONE" -NoNewline -ForegroundColor Green
            Write-Host "] - Cleanup of Rapid Circle Migration Street is done"
        }

        $question = Read-Host "Would you like to start the installation of the Migration Street? (Y/N)"
        if ($question -ne "Y") {
            Write-Host "[" -NoNewline
            Write-Host "DONE" -NoNewline -ForegroundColor Green
            Write-Host "] - Rapid Circle Migration Street is downloaded and unzipped to $TargetLocation"
            break
        }
        else {
            CD $targetLocation

            Write-Host "[" -NoNewline
            Write-Host "WAITING" -NoNewline -ForegroundColor YELLOW
            Write-Host "] - Rapid Circle Migration Street requirements are being installed" -NoNewline
            # starting installation
            $requirements = $targetLocation + "\install-requirements.ps1"
            Start-Process Powershell -verb runAs -ArgumentList "-file $requirements" -wait
            Write-host "    DONE" -ForegroundColor Green

            Write-Host "[" -NoNewline
            Write-Host "WAITING" -NoNewline -ForegroundColor YELLOW
            Write-Host "] - Rapid Circle Migration Street app registration is being installed" -NoNewline
            $appregistration = $targetLocation + "\install-app-registration.ps1"
            Start-Process Powershell -verb runAs -ArgumentList "-file $appregistration" -wait
            Write-host "    DONE" -ForegroundColor Green

            $content = Import-PowerShellDataFile ".\import-data\sharepoint-information.psd1"
            if ($Content.teams -eq $true) {

                Write-Host "[" -NoNewline
                Write-Host "WAITING" -NoNewline -ForegroundColor YELLOW
                Write-Host "] - Rapid Circle Migration Street Teams is being installed" -NoNewline
                $teamsinstall = $targetLocation + "\install-Migrationteams.ps1"
                Start-Process Powershell -verb runAs -ArgumentList "-file $teamsinstall" -wait
                Write-host "    DONE" -ForegroundColor Green
            }

        }
    }
}



