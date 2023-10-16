# module for Microsoft Graph API
function Get-TokenForGraphAPI {
    param(
        [Parameter(Mandatory)]
        [string]$appid,
        [Parameter(Mandatory)]
        [string]$clientsecret,
        [Parameter(Mandatory)]
        [string]$tenantid
    )

    
    $Body = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $appid
        Client_Secret = $ClientSecret
    }
 
    $Connection = Invoke-RestMethod `
        -Uri https://login.microsoftonline.com/$($TenantID)/oauth2/v2.0/token `
        -Method POST `
        -Body $body
 
    #Get the Access Token
    return $connection

}
function Get-TokenForGraphAPIWithCertificate {
    param(
        [Parameter(Mandatory)]
        [string]$appid,
        [Parameter(Mandatory)]
        [string]$tenantname,
        [Parameter(Mandatory)]
        [string]$Thumbprint
    )

    Try {
        $Certificate = Get-Item Cert:\CurrentUser\My\$Thumbprint -ErrorAction SilentlyContinue
        if ($Certificate -eq $null) {
            $password = (Get-Content -Path .\security\certificate_password.txt | ConvertTo-SecureString -AsPlainText -Force)
            $Certificate = Get-PfxCertificate -FilePath '.\security\Rapid Circle Migration Street.pfx' -Password $password
        }
    }
    catch {
        #do nothing
    }


    $Scope = "https://graph.microsoft.com/.default"  
  
    # Create base64 hash of certificate  
    $CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())  
  
    # Create JWT timestamp for expiration  
    $StartDate = (Get-Date "1970-01-01T00:00:00Z" ).ToUniversalTime()  
    $JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds  
    $JWTExpiration = [math]::Round($JWTExpirationTimeSpan, 0)  
  
    # Create JWT validity start timestamp  
    $NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds  
    $NotBefore = [math]::Round($NotBeforeExpirationTimeSpan, 0)  
  
    # Create JWT header  
    $JWTHeader = @{  
        alg = "RS256"  
        typ = "JWT"  
        # Use the CertificateBase64Hash and replace/strip to match web encoding of base64  
        x5t = $CertificateBase64Hash -replace '\+', '-' -replace '/', '_' -replace '='  
    }  
  
    # Create JWT payload  
    $JWTPayLoad = @{  
        # What endpoint is allowed to use this JWT  
        aud = "https://login.microsoftonline.com/$TenantName/oauth2/token"  
  
        # Expiration timestamp  
        exp = $JWTExpiration  
  
        # Issuer = your application  
        iss = $AppId  
  
        # JWT ID: random guid  
        jti = [guid]::NewGuid()  
  
        # Not to be used before  
        nbf = $NotBefore  
  
        # JWT Subject  
        sub = $AppId  
    }  
  
    # Convert header and payload to base64  
    $JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))  
    $EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)  
  
    $JWTPayLoadToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))  
    $EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)  
  
    # Join header and Payload with "." to create a valid (unsigned) JWT  
    $JWT = $EncodedHeader + "." + $EncodedPayload  
  
    # Get the private key object of your certificate  
    $PrivateKey = ([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate))  
  
    # Define RSA signature and hashing algorithm  
    $RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1  
    $HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256  
  
  
    # Create a signature of the JWT  
    $Signature = [Convert]::ToBase64String(  
        $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT), $HashAlgorithm, $RSAPadding)  
    ) -replace '\+', '-' -replace '/', '_' -replace '='  
  
    # Join the signature to the JWT with "."  
    $JWT = $JWT + "." + $Signature  
  
    # Create a hash with body parameters  
    $Body = @{  
        client_id             = $AppId  
        client_assertion      = $JWT  
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"  
        scope                 = $Scope  
        grant_type            = "client_credentials"  
  
    }  
  
    $Url = "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token"  
  
    # Use the self-generated JWT as Authorization  
    $Header = @{  
        Authorization = "Bearer $JWT"  
    }  
  
    # Splat the parameters for Invoke-Restmethod for cleaner code  
    $PostSplat = @{  
        ContentType = 'application/x-www-form-urlencoded'  
        Method      = 'POST'  
        Body        = $Body  
        Uri         = $Url  
        Headers     = $Header  
    }  
  
    $Request = Invoke-RestMethod @PostSplat  

    # View access_token  
    return $Request

}
function RunQueryandEnumerateResults {
    param (
    [Parameter(Mandatory)]
    [string]$apiUri
    )

    #$token = Get-TokenForGraphAPI -appid $appid -clientsecret $clientSecret -Tenantid $tenantid
    $token = Get-TokenForGraphAPIWithCertificate -appid $appid -tenantname $tenantname -Thumbprint $ThumbPrint

    Try {
        $Results = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.access_token)" } -Uri $apiUri -Method Get)
    }
    catch {
        $webError = $_
        $mustRetry = 1
    }
    If ($mustRetry -and ($weberror.ErrorDetails.message -like "*Access token has expired or is not yet valid.*") -or $null -eq $Token.Access_Token) {
        #region connection
        # Get an access token for the Microsoft Graph API
        do {
            try {  
                 #$token = Get-TokenForGraphAPI -appid $appid -clientsecret $clientSecret -Tenantid $tenantid
                $token = Get-TokenForGraphAPIWithCertificate -appid $appid -tenantname $tenantname -Thumbprint $ThumbPrint
                $mustRetry = 0
            }
            catch {
                $webError = $_
                $mustRetry = 1
                Start-Sleep -seconds 2
            }
            
        } while (
            $mustRetry -eq 1
        )
        
        #endregion connection
        $Results = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.access_token)" } -Uri $apiUri -Method Get)
    }

    #Begin populating results
    [array]$ResultsValue = $Results.value

    #If there is a next page, query the next page until there are no more pages and append results to existing set
    if ($null -ne $results."@odata.nextLink") {
        $NextPageUri = $results."@odata.nextLink"
        ##While there is a next page, query it and loop, append results
        While ($null -ne $NextPageUri) {
            $NextPageRequest = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.access_token)" } -Uri $NextPageURI -Method Get)
            $NxtPageData = $NextPageRequest.Value
            $NextPageUri = $NextPageRequest."@odata.nextLink"
            $ResultsValue = $ResultsValue + $NxtPageData
        }
    }

    ##Return completed results
    return $ResultsValue

    
}