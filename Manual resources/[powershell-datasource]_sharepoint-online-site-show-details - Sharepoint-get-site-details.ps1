$siteId = $datasource.selectedSite.SiteId
$siteUrl = $datasource.selectedSite.SPWebUrl

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

try{
    $actionMessage = "show properties"
    $properties = $datasource.selectedSite.psObject.properties | Sort-Object Name

    foreach($tmp in $properties)
    {
        $returnObject = @{name=$tmp.Name; value=$tmp.value}
        Write-Output $returnObject
    }
} catch {
    $ex = $PSItem
    Write-Warning "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    Write-Error "Error $($actionMessage). Error: $($ex.Exception.Message)"
    # exit # use when using multiple try/catch and the script must stop
}


