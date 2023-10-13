function Set-DynamicGroup {
    #Group.ReadWrite.All
    param(
        [Parameter(Mandatory)]
        [string]$Groupid,
        [Parameter(Mandatory)]
        [string]$MembershipRule

    )
    $connectionDetails = Import-Clixml .\security\connectionDetails.xml

    #$token = Get-TokenForGraphAPI -appid $ConnectionDetails.appid -clientsecret $ConnectionDetails.clientSecret -Tenantid $ConnectionDetails.tenantid
    $token = Get-TokenForGraphAPIWithCertificate -appid $connectionDetails.appid -tenantname $connectionDetails.tenantname -Thumbprint $connectionDetails.ThumbPrint

    $params = @{
        groupTypes = @(
            "DynamicMembership"
        )
        membershipRuleProcessingState = "On"
        membershipRule = $MembershipRule
    }

     $body = $params | ConvertTo-Json -Depth 5
     $uri = "https://graph.microsoft.com/beta/groups/$groupId"
     $header = @{
        'Authorization' = "BEARER $($Token.access_token)"
        'Content-type'  = "application/json"
    }
    
    try {  
        $ChangeGroup = Invoke-Restmethod -Method PATCH -Uri $uri -headers $header -Body $body
        return $ChangeGroup
    }
    catch {
        $webError = $_
        $webError = ($weberror | convertFrom-json).error.innererror.code
        return $weberror
    }
}