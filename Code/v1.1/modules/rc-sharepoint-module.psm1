# sharepoint module


function Create-List {
    # #Team.Create
    param(
        [Parameter(Mandatory)]
        [string]$DisplayName,
        [Parameter(Mandatory)]
        [string]$SiteId
    )
    $connectionDetails = Import-Clixml .\security\connectionDetails.xml

    #$token = Get-TokenForGraphAPI -appid $ConnectionDetails.appid -clientsecret $ConnectionDetails.clientSecret -Tenantid $ConnectionDetails.tenantid
    $token = Get-TokenForGraphAPIWithCertificate -appid $connectionDetails.appid -tenantname $connectionDetails.tenantname -Thumbprint $connectionDetails.ThumbPrint

    switch ($displayName) {
        "Fileshare Migrations" {
            $params = @{
                displayName = $displayName
                columns     = @(
                    @{
                        name = "Status"
                        choices = @(
                            "New"
                            "Finalize"
                            "Done"
                            "Error"
                        )
                        displayAs = "dropDownMenu"
                    }
                    @{
                        name = "Date"
                        format = "dateOnly"
                        dateTime = @{
                            }
                    }
                    @{
                        name   = "Share"
                        text = @{
                        }
                    }
                    @{
                        name   = "Team"
                        text = @{
                        }
                    }
                    @{
                        name   = "Channel"
                        text = @{
                        }
                    }
                    @{
                        name   = "Folder"
                        text = @{
                        }
                    }
                )
                list        = @{
                    template = "genericList"
                }
            }
        }
        "Homefolder Migrations" {
            $params = @{
                displayName = $displayName
                columns     = @(
                    @{
                        name = "Status"
                        text = @{
                        }
                    }
                    @{
                        name = "Date"
                        dateTime = @{
                            }
                    }
                    @{
                        name   = "Homedir"
                        text = @{
                        }
                    }
                    @{
                        name   = "UserPrincipalName"
                        chooseFromType = "peopleOnly"
                        personOrGroup = @{
                        }
                    }
                    @{
                        name   = "Folder"
                        text = @{
                        }
                    }
                )
                list        = @{
                    template = "genericList"
                }
            }
        }
        "Errorlist" {
            $params = @{
                displayName = $displayName
                columns     = @(
                    @{
                        name = "Status"
                        text = @{
                        }
                    }
                    @{
                        name = "Date"
                        dateTime = @{
                            }
                    }
                    @{
                        name   = "Source"
                        text = @{
                        }
                    }
                    @{
                        name   = "Destination"
                        text = @{
                        }
                    }
                    @{
                        name   = "Error"
                        text = @{
                        }
                    }
                )
                list        = @{
                    template = "genericList"
                }
            }
        }
    }
    


    $body = $params | ConvertTo-Json -Depth 5
    $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId"
    $header = @{
        'Authorization' = "BEARER $($Token.access_token)"
        'Content-type'  = "application/json"
    }
    

    try {  
        $createlist = Invoke-Restmethod -Method post -Uri $uri -headers $header -Body $body
        return $createlist
    }
    catch {
        $webError = $_
        $webError = ($weberror | convertFrom-json).error.innererror.code
        return $weberror
    }
}

function Create-ListColumns {
    param(
        [Parameter(Mandatory)]
        [string]$ListId,
        [Parameter(Mandatory)]
        [string]$SiteId,
        [Parameter(Mandatory)]
        [String]$Column

    )
    $connectionDetails = Import-Clixml .\security\connectionDetails.xml

    $token = Get-TokenForGraphAPI -appid $ConnectionDetails.appid -clientsecret $ConnectionDetails.clientSecret -Tenantid $ConnectionDetails.tenantid
    # $token = Get-TokenForGraphAPIWithCertificate -appid $connectionDetails.appid -tenantname $connectionDetails.tenantname -Thumbprint $connectionDetails.ThumbPrint

    $uri = "https://graph.microsoft.com/v1.0/sites/$siteid/lists/$listid/columns"

    $params = @{
        description         = $column
        enforceUniqueValues = $false
        hidden              = $false
        indexed             = $false
        name                = $column
        text                = @{
            allowMultipleLines          = $false
            appendChangesToExistingText = $false
            linesForEditing             = 0
            maxLength                   = 255
        }
    } 
    $body = $params | ConvertTo-Json -Depth 5
    $header = @{
        'Authorization' = "BEARER $($Token.access_token)"
        'Content-type'  = "application/json"
    }
    Invoke-Restmethod -Method POST -Uri $uri -headers $header -Body $body
    
    #     # try {  
    #     #     $createcolumn = Invoke-Restmethod -Method POST -Uri $uri -headers $header -Body $body
    #     #     return $createcolumn
    #     # }
    #     # catch {
    #     #     $webError = $_
    #     #     $webError = ($weberror | convertFrom-json).error.innererror.code
    #     #     return $weberror
    #     # }

}

function Update-Item {
    $Body = @{
        fields = @{
            Title = "Test"
        }
    }

    $GraphUrl = "https://graph.microsoft.com/v1.0/sites/$SiteID/lists/$ListTitle/items"

    $BodyJSON = $Body | ConvertTo-Json -Compress
    Invoke-RestMethod -Uri $GraphUrl -Method 'POST' -Body $BodyJSON -Headers $Header -ContentType "application/json" 
    
}