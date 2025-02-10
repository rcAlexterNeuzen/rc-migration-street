[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$InstallFolder = "C:\RCScripts",
    [Parameter(Mandatory=$true)]
    [SecureString]$Password,
    [switch]$Force
)

#region Configuration
$script:Config = @{
    Repository = "rcalexterneuzen/rc-migration-street"
    ApiUrl = "https://www.checkyourlic.org:443"
    TempPath = Join-Path $env:TEMP "RCStreet"
    LogPath = Join-Path $env:TEMP "RCMigrationStreet_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    TranscriptPath = Join-Path $env:TEMP "RCMigrationStreet_$(Get-Date -Format 'yyyyMMdd_HHmmss')_transcript.log"
    RequiredModules = @('7Zip4Powershell', 'CredentialManager')
}

#region Helper Functions
function Write-RCLog {
    [CmdletBinding()]
    param(
        [string]$Message,
        [ValidateSet('INFO','WARNING','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )
    
    $Colors = @{
        'INFO' = 'White'
        'WARNING' = 'Yellow'
        'ERROR' = 'Red'
        'SUCCESS' = 'Green'
    }
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$TimeStamp][$Level] $Message" -ForegroundColor $Colors[$Level]
    Add-Content -Path $Config.LogPath -Value "[$TimeStamp][$Level] $Message"
}

function Test-Prerequisites {
    [CmdletBinding()]
    param()
    
    # Check admin rights
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Administrative privileges required"
    }
    
    # Install required modules
    foreach ($module in $Config.RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-RCLog "Installing required module: $module" -Level WARNING
            Install-Module -Name $module -Force -Scope CurrentUser
        }
        Import-Module $module -Force
    }
    
    # Ensure folders exist
    @($InstallFolder, "$InstallFolder\Backup", $Config.TempPath) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -Path $_ -ItemType Directory -Force | Out-Null
        }
    }
}

function Get-LatestVersion {
    [CmdletBinding()]
    param()
    
    try {
        $releases = "https://api.github.com/repos/$($Config.Repository)/releases"
        $latest = (Invoke-RestMethod -Uri $releases)[0]
        return @{
            Tag = $latest.tag_name
            Version = $latest.tag_name -replace 'v',''
            DownloadUrl = $latest.assets | Where-Object name -eq 'rc-migration-street.zip' | Select-Object -ExpandProperty browser_download_url
        }
    }
    catch {
        throw "Failed to get latest version: $_"
    }
}

function New-RCBackup {
    [CmdletBinding()]
    param(
        [string]$Version
    )
    
    $backupPath = Join-Path "$InstallFolder\Backup" "Backup-$(Get-Date -Format 'yyyyMMdd')-v$Version.zip"
    $files = Get-ChildItem -Path $InstallFolder -Exclude 'Backup'
    Compress-Archive -Path $files -DestinationPath $backupPath -Force
    return $backupPath
}

function Install-RCMigrationStreet {
    [CmdletBinding()]
    param(
        [string]$DownloadUrl,
        [string]$Version,
        [string]$BackupPath
    )
    
    try {
        # Download
        $zipPath = Join-Path $Config.TempPath "rc-migration-street-$Version.zip"
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $zipPath
        
        # Extract
        Expand-7Zip -ArchiveFileName $zipPath -TargetPath $InstallFolder -SecurePassword $Password
        
        # Cleanup
        Remove-Item -Path $zipPath -Force
        Remove-Item -Path "$InstallFolder\__MACOSX" -Recurse -Force -ErrorAction SilentlyContinue
        Get-ChildItem -Path $InstallFolder -Filter ".DS_Store" -Recurse | Remove-Item -Force
        
        return $true
    }
    catch {
        if ($BackupPath) {
            Write-RCLog "Installation failed, restoring backup..." -Level WARNING
            Expand-Archive -Path $BackupPath -DestinationPath $InstallFolder -Force
        }
        throw
    }
}

#region Main Execution
try {
    Start-Transcript -Path $Config.TranscriptPath
    Write-RCLog "Starting RC Migration Street installation" -Level INFO
    
    # Initialize
    Test-Prerequisites
    
    # Get versions
    $latest = Get-LatestVersion
    $current = if (Test-Path "$InstallFolder\version.json") {
        (Get-Content "$InstallFolder\version.json" | ConvertFrom-Json).Version
    } else { $null }
    
    # Check if update needed
    if ($current -eq $latest.Version -and -not $Force) {
        Write-RCLog "Already running latest version: $current" -Level SUCCESS
        return
    }
    
    # Perform installation/update
    Write-RCLog "Installing/Updating to version $($latest.Version)" -Level INFO
    $backup = if ($current) { New-RCBackup -Version $current }
    
    if (Install-RCMigrationStreet -DownloadUrl $latest.DownloadUrl -Version $latest.Version -BackupPath $backup) {
        Write-RCLog "Installation completed successfully" -Level SUCCESS
    }
}
catch {
    Write-RCLog $_.Exception.Message -Level ERROR
    exit 1
}
finally {
    Stop-Transcript
    $originalpaths = (Get-ItemProperty -Path ‘Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment’ -Name PSModulePath).PSModulePath

# Add your new path to below after the ;

    $newPath=$originalpaths+’;$InstallFolder\modules’

Set-ItemProperty -Path ‘Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment’ -Name PSModulePath –Value $newPath
    Remove-Item -Path $Config.TempPath -Recurse -Force -ErrorAction SilentlyContinue
}
