
$scriptFileName = ($PSCommandPath | Split-Path -Leaf) -replace '\..*$'
$pathMyDocuments = [environment]::GetFolderPath('MyDocuments')
$pathScriptFiles = "$pathMyDocuments\$scriptFileName"
$configFile = "$pathScriptFiles\$scriptFileName.xml"

try {
    $appProperties = Import-CliXml `
        -Path $configFile `
        -ErrorAction Stop
    # TODO Ask for values and create if missing
    $tenantId = $appProperties.TenantId
    $appId = $appProperties.AppId
    $appSecret = $appProperties.AppSecret
} catch {
    throw "Error importing app properties file $($configFile): $($_.ToString())"
}

# Construct URI and body needed for authentication
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $appId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $appSecret
    grant_type    = "client_credentials"
}

function Get-AzGraphToken
{
    try {
        $tokenRequest = Invoke-WebRequest `
            -Method Post `
            -Uri $uri `
            -ContentType "application/x-www-form-urlencoded" `
            -Body $body -UseBasicParsing `
            -ErrorAction Stop

        ($tokenRequest.Content | ConvertFrom-Json -ErrorAction Stop).access_token
    } catch {
        throw "Error getting OAuth 2.0 Token: $($_.ToString())"
    }
}

# Based on https://github.com/12Knocksinna/Office365itpros/blob/master/ReportDLMembershipsCountsGraph.PS1
function Get-AzGraphData
{
    param (
        [parameter(Mandatory=$true)][String]$AccessToken,
        [parameter(Mandatory=$true)][String]$Uri
    )
    $headers = @{
        'Content-Type'  = "application\json"
        'Authorization' = "Bearer $AccessToken" 
        'ConsistencyLevel' = "eventual"
    }

    # Invoke REST method and fetch data until there are no pages left.
    while ($uri) {
        $statusCode = 429 # Means requests are getting throttled. Setting here to trick the while loop 
                          # to iterate at least once.
        while ($statusCode -eq 429) {
            try {
                $result = Invoke-RestMethod `
                    -Headers $headers `
                    -Uri $Uri `
                    -UseBasicParsing `
                    -Method "GET" `
                    -ContentType "application/json" `
                    -ErrorAction Stop

                $statusCode = $result.StatusCode
            } catch {
                $statusCode = $_.Exception.Response.StatusCode.value__
                if ($statusCode -eq 429) {
                    Write-Warning "Got throttled by Microsoft. Sleeping for 45 seconds..."
                    Start-Sleep -Seconds 45
                } else {
                    Write-Error $_.Exception
                }
            }
        }

        # TODO Try to see why it's sometimes $result.value and sometimes $result. Apparently.
        if ($result.value) {
            $result.value
        } else {
            $result
        }

        $uri = $result.'@odata.nextlink'
    }
}

function Get-AzUserAuthenticationMethods
{
    Param(
        [Parameter(Mandatory=$false)][String]$UserPrincipalName
    )
    # See https://docs.microsoft.com/en-us/graph/authenticationmethods-get-started#api-reference
    $token = Get-AzGraphToken
    if ($UserPrincipalName) {
        $uri = "https://graph.microsoft.com/beta/users/$UserPrincipalName" + '?$select=userPrincipalName'
    } else {
        $uri = 'https://graph.microsoft.com/beta/users?$select=userPrincipalName'
    }
    Get-AzGraphData -AccessToken $token -Uri $uri | ForEach-Object {
        $upn = $_.userPrincipalName
        # NOTE UserAuthenticationMethod.Read.All is required to access the authentication methods API
        $uri = "https://graph.microsoft.com/beta/users/$upn/authentication/methods"
        Get-AzGraphData -AccessToken $token -Uri $uri | `
            Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $upn -PassThru
    }
}

function Find-AzADUserByAuthenticationMethod
{
    Param(
        [Parameter(Mandatory=$true)][String]$Type,
        [Parameter(Mandatory=$false)][String]$SearchString
    )
    $typeMap = @{
        EMailAddress = @{
            odataType = '#microsoft.graph.emailAuthenticationMethod'
            valuePropertyName = 'emailAddress'
            additionalMatchProperties = @()
        }
        Mobile = @{
            odataType = '#microsoft.graph.phoneAuthenticationMethod'
            valuePropertyName = 'phoneNumber'
            additionalMatchProperties = @(
                @{
                    Name = 'phoneType'
                    Value = 'mobile'
                }
            )
        }
    }
    $typeDef = $typeMap.$Type
    Get-AzUserAuthenticationMethods | ForEach-Object {
        if ($_.'@odata.type' -eq $typeDef.odataType) {
            foreach ($property in $typeDef.additionalMatchProperties) {
                if ($_.$property.Name -ne $property.Value) {
                    return $null
                }
            }
            if ($_.($typeDef.valuePropertyName) -match $SearchString) {
                return $_ | Select-Object UserPrincipalName, $typeDef.valuePropertyName
            }
        }
    }
}
