function Send-MailToinform {
    param(
        [Parameter(Mandatory)]
        [string]$to,
        [Parameter(Mandatory)]
        [string]$from,
        [Parameter(Mandatory)]
        [string]$Subject,        
        [Parameter(Mandatory)]
        [String]$Message
    )

    $emailSender = $from
    $emailSubject = $subject
    #endregion 1
    try {
        ## getting token 
        if ($PSVersionTable.PSVersion.Major -ne 5) {
            $token = Get-TokenForGraphAPIWithCertificate -appid $appid -tenantname $tenantname -Thumbprint $ThumbPrint
        }
        else {
            $token = Get-TokenForGraphAPI -appid $appid -tenantid $TenantId -clientsecret $clientsecret
        }
        if ($Token.access_token -ne $null) {
            
           Write-Host "Token is still valid"

        }
        else {
            Write-Host "Token is not valid. Aborting..."
            break

        }
    }
    catch {
        #do nothing
        $Out = $_.Exception.Message
        break
    }

    $header = @{
        'Authorization' = "BEARER $($Token.access_token)"
        'Content-Type' = "Application/json"
    }

    $filename = (Get-Item -path $File).fullName
    $name = (Get-Item -path $File).name
    $base64string = [Convert]::ToBase64String([IO.File]::ReadAllBytes($Filename))

    #region 2: Run
    $params = @{
        message = @{
            subject = $emailSubject
            body = @{
                contentType = "HTML"
                content = $message
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $to
                    }
                }
            )
            ccRecipients = @(
                @{
                    emailaddress = @{
                        address = "migration-onboard@rapidcircle.com"
                    }
                }
            )
            attachments = @(
                @{
                    "@odata.type" = "#microsoft.graph.fileAttachment"
                    name = $Name
                    contentType = "text/plain"
                    contentBytes = $base64string
                }
            )
        }
        saveToSentItems = $true
    }
    # Send the message
    $Body = $Params | ConvertTo-JSON -Depth 6
    Try {
    $out = Invoke-RestMethod -method POST -Header $header -uri "https://graph.microsoft.com/v1.0/users/$emailSender/sendMail" -body $body
    $out = "Message is sent"
    }
    catch {
        $out = $_
    }

    return $out
}

function Log-Message([String]$Message, [string]$file, [string]$Status) {
    if (!($Status)) {
        Write-Host $message
        $messageLog = "[$(Get-Date -format "HH:mm:ss")] " + $message
        Add-Content -Path $file $MessageLog
    }
    else {
        $Status = $Status.ToUpper()
        $messageLog = "[$(Get-Date -format "HH:mm:ss")] " + "[$STATUS] - " + $message
        Add-Content -Path $file $MessageLog


        # to console
        Write-Host "[" -NoNewline
        switch ($status) {
            INFO { Write-host "INFO" -NoNewline }
            WARNING { Write-host "WARNING" -NoNewline -ForegroundColor Yellow }
            WAITING { Write-host "WAITING" -NoNewline -ForegroundColor Yellow }
            SKIPPING { Write-host "SKIPPING" -NoNewline -ForegroundColor Yellow }
            ERROR { Write-host "ERROR" -NoNewline -ForegroundColor RED }
            ADDED { Write-host "ADDED" -NoNewline }
            DONE { Write-Host "DONE" -ForegroundColor green -NoNewLine }
            FINISHED { Write-Host "FINISHED" -ForegroundColor green -NoNewLine }
        }
        Write-Host "] - $message"
    }
}

function Connect-SPMT {
    param (
        [bool]$ScanOnly,
        [string]$CreatedAfter,
        [string]$ModifiedAfter,
        [bool]$LoginWithWeb = $false
    )

    $extensions = "tmp", "pst"

    if (!($CreatedAfter)) {
        $CreatedAfter = "01-01-1970"
    }
    if (!($ModifiedAfter)) {
        $ModifiedAfter = "01-01-1970"
    }

    $ConnectionError = $null
    $count = 0
    do {
        if ($count -eq 4){
            Log-Message -file $file -status Error -Message "There was a problem connecting: $($ConnectionError.Exception.Message)"
            break
        }
        Try {
            $Activity = "Connecting to sharepoint migration services"
            Write-Progress -Activity $Activity
            #Register-SPMTMigration -SPOCredential $Global:SPOCredential -Force
            # scan only with weblogin
            if ($Scanonly -and $LoginWithWeb) {
                Register-SPMTMigration -Force -ScanOnly $True -IncludeHiddenFiles $False -MigrateWithoutRootFolder -DeleteTempFilesWhenMigrationDone $true -SkipFilesWithExtension $extensions -EnableMultiRound $true -MigrateFilesCreatedAfter $CreatedAfter -MigrateFilesModifiedAfter $ModifiedAfter
            }       
            # scan only with local credentials
            elseif ($ScanOnly) {
                Register-SPMTMigration -SPOCredential $Global:SPOCredential -Force -ScanOnly $True -IncludeHiddenFiles $False -MigrateWithoutRootFolder -DeleteTempFilesWhenMigrationDone $true -SkipFilesWithExtension $extensions -EnableMultiRound $true  -MigrateFilesCreatedAfter $CreatedAfter -MigrateFilesModifiedAfter $ModifiedAfter
            }
            # migration session with weblogin
            elseif ($LoginWithWeb) {
                Register-SPMTMigration -Force -IncludeHiddenFiles $False -MigrateWithoutRootFolder -DeleteTempFilesWhenMigrationDone $true -SkipFilesWithExtension $extensions -EnableMultiRound $true  -MigrateFilesCreatedAfter $CreatedAfter -MigrateFilesModifiedAfter $ModifiedAfter
            }
            # migration session with local credentials
            else {
                Register-SPMTMigration -SPOCredential $Global:SPOCredential -Force -IncludeHiddenFiles $False -MigrateWithoutRootFolder -DeleteTempFilesWhenMigrationDone $true -SkipFilesWithExtension v -EnableMultiRound $true  -MigrateFilesCreatedAfter $CreatedAfter -MigrateFilesModifiedAfter $ModifiedAfter
            }

            Log-Message -File $file -Status "DONE" -Message "Connected to sharepoint migration services"
            Write-Progress -Activity $Activity -Completed
            $ConnectionError = $null
        }
        catch {
            $ConnectionError = $_
            $Count++
            Log-Message -File $file -Status "WAITING" -Message "- Retrying to connect ..."
            Write-Progress -Activity $Activity -Completed
            # break
        }
    }
    while (
        $ConnectionError -ne $null
    )

}

function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs = ""
    return [String]$characters[$random]
}
function Scramble-String([string]$inputString) {     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}

function New-FileShareMigration {
    param (
        [string]$CreatedAfter,
        [string]$ModifiedAfter,
        [array]$ToMigrate
        )

        

}