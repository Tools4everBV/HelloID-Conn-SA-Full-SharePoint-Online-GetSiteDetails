# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

try {
    $searchValue = $datasource.searchValue
    $searchQuery = "*$searchValue*"
      
      
    if([String]::IsNullOrEmpty($searchValue) -eq $true){
        return
    }else{
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

        Write-Information -Message "Searching for: $searchQuery"
        #Add the authorization header to the request
        $authorization = @{
            Authorization = "Bearer $accesstoken";
            'Content-Type' = "application/json";
            Accept = "application/json";
        }
 
        $baseSearchUri = "https://graph.microsoft.com/"
        $searchUri = $baseSearchUri + "v1.0/groups"
        $groupsResponse = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false          

        #Write-Information ($teamsResponse.value | ConvertTo-Json)
        $groups = foreach($groupObject in $groupsResponse.value){
            
            if( $groupObject.displayName -like $searchQuery -or $groupObject.MailNickName -like $searchQuery -or $groupObject.Mailaddress -like $searchQuery ){
                $groupObject
            }
        }

        $resultCount = @($groups).Count
        Write-Information -Message "Result count: $resultCount"
         
        if($resultCount -gt 0){
            foreach($group in $groups){
                $siteUri = $searchUri + "/" + $($group.Id)+ "/sites/root"
                $site = Invoke-RestMethod -Uri $siteUri -Method Get -Headers $authorization -Verbose:$false
                #Write-Information $site.WebUrl
                $returnObject = @{DisplayName=$group.DisplayName; Description=$group.Description; MailNickName=$group.MailNickName; GroupId=$group.Id; Site=$site.WebUrl; SiteId=$site.id}
                Write-Output $returnObject
            }
        } else {
            return
        }
    }
} catch {
    
    Write-Error -Message ("Error searching for Sharepoint Sites. Error: $($_.Exception.Message)" + $errorDetailsMessage)
    Write-Warning -Message "Error searching for Sharepoint Sites"
     
    return
}
