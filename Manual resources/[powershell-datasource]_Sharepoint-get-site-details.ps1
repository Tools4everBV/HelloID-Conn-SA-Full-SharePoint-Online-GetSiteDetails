$groupId = $datasource.selectedSite.Groupid
$siteUrl = $datasource.selectedSite.Site
# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

try {
    
        Write-Information -Message "Generating Microsoft Graph API Access Token user.."

        $baseUri = "https://login.microsoftonline.com/"
        $authUri = $baseUri + "$AADTenantID/oauth2/token"

        $body = @{
            grant_type      = "client_credentials"
            client_id       = "$AADAppId"
            client_secret   = "$AADAppSecret"
            resource        = "https://graph.microsoft.com"
        }
 
        $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
        $accessToken = $Response.access_token;

        #Add the authorization header to the request
        $authorization = @{
            Authorization = "Bearer $accesstoken";
            'Content-Type' = "application/json";
            Accept = "application/json";
        }
 
        $baseSearchUri = "https://graph.microsoft.com/"

        $siteUri = $baseSearchUri + "v1.0/groups/" + $groupId + "/sites/root"
        $siteResponse = Invoke-RestMethod -Uri $siteUri -Method Get -Headers $authorization -Verbose:$false          
        
        
        $resultCount = @($siteResponse).Count
        Write-Information -Message "Result count: $resultCount"
         
        if($resultCount -gt 0){         
            $returnObject = @{DisplayName=$siteResponse.DisplayName; Description=$siteResponse.Description; Created=$siteResponse.createdDateTime; LastModified=$siteResponse.lastModifiedDateTime; Name=$siteResponse.name; WebUrl=$siteResponse.webUrl; HostName=$siteResponse.SiteCollection.hostname}
            Write-Output $returnObject        
        } else {
            return
        }
    
} catch {
    
    Write-Error -Message ("Error getting site-details for SharePoint Site . Error: $($_.Exception.Message)" + $errorDetailsMessage)
    Write-Warning -Message "Error getting site-details for SharePoint Site"
     
    return
}
