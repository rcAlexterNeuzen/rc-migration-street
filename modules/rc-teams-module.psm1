# teams provisioning module
function Get-TeamsChannels {
    #Channel.ReadBasic.All
    param(
        [Parameter(Mandatory)]
        [string]$teamsid
        )

    $apiUri = "https://graph.microsoft.com/v1.0/teams/$teamsid/Channels"
    $channels = RunQueryandEnumerateResults -ApiUri $apiUri
    return $channels
}
function Create-Teams {
    # #Team.Create
    param(
        [Parameter(Mandatory)]
        [string]$Teamsname,
        [Parameter(Mandatory)]
        [string]$Owner,
        [Parameter(Mandatory)]
        [string]$Description

    )
    #$token = Get-TokenForGraphAPI -appid $appid -clientsecret $clientSecret -Tenantid $tenantid
    $token = Get-TokenForGraphAPIWithCertificate -appid $appid -tenantname $tenantname -Thumbprint $ThumbPrint

     $userDatabind = "https://graph.microsoft.com/v1.0/users('$($owner)')"

     $params = @{
        "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
        displayName    = $Teamsname
        description    = $description
        members        = @(
            @{
                "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
                "user@odata.bind" = $userDatabind
                roles             = @(
                    "owner"
                )
            }
        )
    }

     $body = $params | ConvertTo-Json -Depth 5
     $uri = "https://graph.microsoft.com/v1.0/teams"
     $header = @{
        'Authorization' = "BEARER $($Token.access_token)"
        'Content-type'  = "application/json"
    }
    
    try {  
        $createTeams = Invoke-Restmethod -Method POST -Uri $uri -headers $header -Body $body
        return $createTeams
    }
    catch {
        $webError = $_
        $webError = ($weberror | convertFrom-json).error.innererror.code
        return $weberror
    }
}
function Create-Channel {
    param(
        [Parameter(Mandatory)]
        [string]$TeamsId,
        [Parameter(Mandatory)]
        [string]$ChannelName,
        [Parameter(Mandatory)]
        [string]$Description,
        [Parameter(Mandatory)]
        [string]$owner,
        [Parameter(Mandatory)]
        [string]$Type
    )

   #$token = Get-TokenForGraphAPI -appid $appid -clientsecret $clientSecret -Tenantid $tenantid
   $token = Get-TokenForGraphAPIWithCertificate -appid $appid -tenantname $tenantname -Thumbprint $ThumbPrint

    $userDatabind = "https://graph.microsoft.com/v1.0/users('$($owner)')"

    $params = @{
        "@odata.type"  = "#Microsoft.Graph.channel"
        membershipType = $Type
        displayName    = $ChannelName
        description    = $description
        members        = @(
            @{
                "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
                "user@odata.bind" = $userDatabind
                roles             = @(
                    "owner"
                )
            }
        )
    }

    $body = $params | ConvertTo-Json -Depth 5
    $apiUri = "https://graph.microsoft.com/v1.0/Teams/$teamsId/channels"
    $header = @{
        'Authorization' = "BEARER $($Token.access_token)"
        'Content-type'  = "application/json"
    }

    try {  
        $createChannel = Invoke-Restmethod -Method POST -Uri $apiUri -headers $header -Body $body
    }
    catch {
        $webError = $_
        $createChannel = ($weberror | convertFrom-json).error.innererror.code
    }
    return $createChannel
}
function Add-MemberToTeams {
    #TeamMember.ReadWrite.All
    param(
        [Parameter(Mandatory)]
        [string]$TeamsId,
        [Parameter(Mandatory)]
        [string]$Upn
    )

    #$token = Get-TokenForGraphAPI -appid $appid -clientsecret $clientSecret -Tenantid $tenantid
    $token = Get-TokenForGraphAPIWithCertificate -appid $appid -tenantname $tenantname -Thumbprint $ThumbPrint

    $uri = "https://graph.microsoft.com/v1.0/teams/$TeamsId/members"
    $userDatabind = "https://graph.microsoft.com/v1.0/users('$($upn)')"

    $params = @{
        "@odata.type"  = "#microsoft.graph.aadUserConversationMember"
        "user@odata.bind" = $userDatabind
        # roles = @(
        # )
    }

    $header = @{
        'Authorization' = "BEARER $($Token.access_token)"
        'Content-type'  = "application/json"
    }

    $body = $params | ConvertTo-Json -Depth 5

    Try {
        $addMember = Invoke-Restmethod -method Post -Uri $uri -Headers $header -body $body

    }

    catch {
        $Fault = $_.Exception.Message
        return $Fault
    }

    return $addMember
}
function Add-MemberToTeamsChannel {
    param(
        [Parameter(Mandatory)]
        [string]$TeamsId,
        [Parameter(Mandatory)]
        [string]$ChannelId,
        [Parameter(Mandatory)]
        [string]$Upn
    )

   #$token = Get-TokenForGraphAPI -appid $appid -clientsecret $clientSecret -Tenantid $tenantid
    $token = Get-TokenForGraphAPIWithCertificate -appid $appid -tenantname $tenantname -Thumbprint $ThumbPrint
    
    ForEach ($user in $upn){
        $params = @{
            "@odata.type" = "#microsoft.graph.aadUserConversationMember"
            "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($user)')"
        }

        $uri = "https://graph.microsoft.com/v1.0/teams/$teamsid/channels/$channelid/members"
        $body = $params | ConvertTo-Json -Depth 5
        $header = @{
            'Authorization' = "BEARER $($Token.access_token)"
            'Content-type'  = "application/json"
            roles = @(
                "member"
            )
        }
        $AddUser = Invoke-Restmethod -method Post -Uri $uri -Body $body -Headers $header

        return $adduser
    }
}
function Add-OwnerToTeams {
    #TeamMember.ReadWrite.All
    param(
        [Parameter(Mandatory)]
        [string]$TeamsName,
        [Parameter(Mandatory)]
        [string]$Upn,
        [Parameter(Mandatory)]
        [string]$TeamsId
    )

   #$token = Get-TokenForGraphAPI -appid $appid -clientsecret $clientSecret -Tenantid $tenantid
   $token = Get-TokenForGraphAPIWithCertificate -appid $appid -tenantname $tenantname -Thumbprint $ThumbPrint
    
     $uri = "https://graph.microsoft.com/v1.0/teams/$TeamsId/members"
    $userDatabind = "https://graph.microsoft.com/v1.0/users('$($upn)')"

    $params = @{
        "@odata.type"  = "#microsoft.graph.aadUserConversationMember"
        "user@odata.bind" = $userDatabind
        roles = @(
            "owner"
        )
    }

    $header = @{
        'Authorization' = "BEARER $($Token.access_token)"
        'Content-type'  = "application/json"
    }

    $body = $params | ConvertTo-Json -Depth 5

    Try {
        $addOwner = Invoke-Restmethod -method Post -Uri $uri -Headers $header -body $body
        return $addOwner
    }

    catch {

    }
}
function Add-OwnerToTeamsChannel {
    param(
        [Parameter(Mandatory)]
        [string]$TeamsId,
        [Parameter(Mandatory)]
        [string]$ChannelId,
        [Parameter(Mandatory)]
        [string]$Upn
    )

   #$token = Get-TokenForGraphAPI -appid $appid -clientsecret $clientSecret -Tenantid $tenantid
   $token = Get-TokenForGraphAPIWithCertificate -appid $appid -tenantname $tenantname -Thumbprint $ThumbPrint
    

    ForEach ($user in $upn){
        $params = @{
            "@odata.type" = "#microsoft.graph.aadUserConversationMember"
            roles = @(
                "owner"
            )
            "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($user)')"
        }

        $uri = "https://graph.microsoft.com/beta/teams/$teamsid/channels/$channelid/members"
        $body = $params | ConvertTo-Json -Depth 5
        $header = @{
            'Authorization' = "BEARER $($Token.access_token)"
            'Content-type'  = "application/json"
        }
        $AddOwner = Invoke-Restmethod -method Post -Uri $uri -Body $body -Headers $header

        return $addOwner
    }
}

function Send-MessageToTeamsChannel {
    param(
        [Parameter(Mandatory)]
        [string]$TeamsId,
        [Parameter(Mandatory)]
        [string]$ChannelId,
        [Parameter(Mandatory)]
        [string]$Message,
        [Parameter(Mandatory)]
        [string]$Type,
        [Parameter(Mandatory)]
        [array]$Attachment
    )

   #$token = Get-TokenForGraphAPI -appid $appid -clientsecret $clientSecret -Tenantid $tenantid
   $token = Get-TokenForGraphAPIWithCertificate -appid $appid -tenantname $tenantname -Thumbprint $ThumbPrint
    
    $params = @{
        body = @{
            content = "Hi<br>
            The migration is DONE.<br>
            $type migrations are completed at $(Get-Date -format "dd-MM-yyyy hh:mm"). Attached you will find the logging<br>
            <attachment id=$attachementId></attachment>"
        }
        attachments = @(
            @{
                id = $Attachment.id
                contentType = "reference"
                contentUrl = $Attachment.ContentUrl
                name = $Attachment.Name
            }
        )

    }

}

function Upload-MigrationLogging {
    param(
        [Parameter(Mandatory)]
        [string]$DriveId,
        [Parameter(Mandatory)]
        [string]$filename
    )

    $connectionDetails = Import-Clixml .\security\connectionDetails.xml

    #$token = Get-TokenForGraphAPI -appid $ConnectionDetails.appid -clientsecret $ConnectionDetails.clientSecret -Tenantid $ConnectionDetails.tenantid
    $token = Get-TokenForGraphAPIWithCertificate -appid $connectionDetails.appid -tenantname $connectionDetails.tenantname -Thumbprint $connectionDetails.ThumbPrint


    $uploadlocation = "https://graph.microsoft.com/v1.0/drives/$driveId/root:/Inventory/$filename" + ':/content'
    Try {
        $upload = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.access_token)" } -method PUT -uri $uploadlocation -InFile $PathExport -ContentType 'multipart/form-data'
    }
    catch {
        Write-Host "ERROR: CANNOT UPLOAD EXPORT"
        Write-Host $_
        return $upload
    }
    return $upload
}