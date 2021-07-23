$siteId = $datasource.selectedSite.Url
$connected = $false

try {
	Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
	$pwd = ConvertTo-SecureString -string $SharePointAdminPWD -AsPlainText -Force
	$cred = New-Object System.Management.Automation.PSCredential $SharePointAdminUser, $pwd
	$null = Connect-SPOService -Url $SharePointBaseUrl -Credential $cred
    Write-Information "Connected to Microsoft SharePoint"
    $connected = $true
}
catch
{	
    Write-Error "Could not connect to Microsoft SharePoint. Error: $($_.Exception.Message)"
    Write-Warning "Failed to connect to Microsoft SharePoint"
}

if ($connected)
{
	try {
        $sites = Get-SPOSite -Identity $siteId
        
        if(@($sites).Count -eq 1){
         foreach($tmp in $sites.psObject.properties)
            {
                $returnObject = [ordered]@{name=$tmp.Name; value=$tmp.value}
                Write-Output $returnObject
            }
        }else{
            return
        }
	}
	catch
	{
		Write-Error "Error getting Site Details. Error: $($_.Exception.Message)"
		Write-Warning -Message "Error getting Site Details"
		return
	}
    finally
    {
        Disconnect-SPOService
        Remove-Module -Name Microsoft.Online.SharePoint.PowerShell
    }
}
else
{
	return
}

