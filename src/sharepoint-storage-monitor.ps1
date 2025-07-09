# SharePoint Storage Monitor
# This script collects SharePoint storage utilization and sends data to Azure Log Analytics

# Get the current script path
$scriptPath = $PSScriptRoot

# Import configuration if exists
$configPath = Join-Path -Path $scriptPath -ChildPath "config.json"
if (Test-Path $configPath) {
    $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
} else {
    # Default configuration
    $config = @{
        TenantId = $env:TenantId
        ClientId = $env:ClientId
        KeyVaultName = $env:KeyVaultName
        SecretName = $env:SecretName
        WorkspaceId = $env:WorkspaceId
        WorkspaceKey = $env:WorkspaceKey
        SharePointSites = @($env:SharePointSites -split ',')
        LogName = "SharePointStorageStats"
    }
}

# Function to check and install modules
function Ensure-Module {
    param (
        [string]$ModuleName,
        [string]$MinimumVersion = ""
    )
    
    Write-Output "Checking for module: $ModuleName"
    
    # Check if module is already installed
    if ($MinimumVersion -eq "") {
        $moduleInstalled = Get-Module -Name $ModuleName -ListAvailable
    } else {
        $moduleInstalled = Get-Module -Name $ModuleName -ListAvailable | Where-Object { $_.Version -ge $MinimumVersion }
    }
    
    # Install module if not present
    if (-not $moduleInstalled) {
        try {
            Write-Output "Installing module: $ModuleName"
            if ($MinimumVersion -eq "") {
                Install-Module -Name $ModuleName -Force -Scope CurrentUser
            } else {
                Install-Module -Name $ModuleName -Force -MinimumVersion $MinimumVersion -Scope CurrentUser
            }
            Write-Output "$ModuleName module installed successfully"
        }
        catch {
            Write-Error "Failed to install $ModuleName module: $_"
            throw
        }
    } else {
        Write-Output "$ModuleName module already installed"
    }
}

# Ensure required modules are installed
try {
    Ensure-Module -ModuleName "PnP.PowerShell" -MinimumVersion "1.12.0"
    Ensure-Module -ModuleName "Az.KeyVault" -MinimumVersion "4.0.0"
} catch {
    Write-Error "Error ensuring required modules: $_"
    throw
}

# Function to get secure credentials from Azure Key Vault
function Get-SecureCredentials {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$KeyVaultName,
        [string]$SecretName
    )
    
    try {
        Write-Output "Authenticating to Azure..."
        # Connect using managed identity if available, otherwise use client credentials
        if ($env:MSI_ENDPOINT) {
            Connect-AzAccount -Identity
        } else {
            # For local development, this assumes az cli is logged in
            # or environment variables are set up
            Connect-AzAccount -TenantId $TenantId
        }
        
        Write-Output "Getting secret from Key Vault..."
        $secret = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name $SecretName
        
        if ($null -eq $secret) {
            throw "Secret not found in Key Vault"
        }
        
        $securePassword = $secret.SecretValue
        $credential = New-Object System.Management.Automation.PSCredential($ClientId, $securePassword)
        
        return $credential
    }
    catch {
        Write-Error "Error getting secure credentials: $_"
        throw
    }
}

# Function to collect SharePoint storage statistics
function Get-SharePointStorageStats {
    param (
        [string]$TenantId,
        [System.Management.Automation.PSCredential]$Credential,
        [string[]]$SiteUrls
    )
    
    $storageData = @()
    
    try {
        # Connect to SharePoint Online
        Write-Output "Connecting to SharePoint Online..."
        Connect-PnPOnline -Url "https://$TenantId-admin.sharepoint.com" -Credential $Credential
        
        foreach ($siteUrl in $SiteUrls) {
            try {
                Write-Output "Getting storage info for site: $siteUrl"
                # Get site usage data
                $site = Get-PnPTenantSite -Url $siteUrl -Detailed
                
                if ($site) {
                    $storageInfo = [PSCustomObject]@{
                        Timestamp = (Get-Date).ToString("o")
                        SiteUrl = $siteUrl
                        SiteTitle = $site.Title
                        StorageUsed = $site.StorageUsageCurrent
                        StorageQuota = $site.StorageQuota
                        PercentageUsed = if ($site.StorageQuota -gt 0) { ($site.StorageUsageCurrent / $site.StorageQuota) * 100 } else { 0 }
                    }
                    $storageData += $storageInfo
                }
            }
            catch {
                Write-Error "Error getting data for site $siteUrl`: $_"
                # Continue with other sites even if one fails
            }
        }
        
        # Disconnect from SharePoint
        Disconnect-PnPOnline
        
        return $storageData
    }
    catch {
        Write-Error "Error collecting SharePoint storage stats: $_"
        throw
    }
}

# Function to send data to Log Analytics
function Send-LogAnalyticsData {
    param (
        [string]$WorkspaceId,
        [string]$WorkspaceKey,
        [string]$LogName,
        [array]$Data
    )
    
    try {
        Write-Output "Preparing data for Log Analytics..."
        
        # Create the function to create the authorization signature
        Function Build-Signature ($customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource) {
            $xHeaders = "x-ms-date:" + $date
            $stringToHash = $method + "`n" + $contentLength + "`n" + $contentType + "`n" + $xHeaders + "`n" + $resource
            
            $bytesToHash = [Text.Encoding]::UTF8.GetBytes($stringToHash)
            $keyBytes = [Convert]::FromBase64String($sharedKey)
            
            $sha256 = New-Object System.Security.Cryptography.HMACSHA256
            $sha256.Key = $keyBytes
            $calculatedHash = $sha256.ComputeHash($bytesToHash)
            $encodedHash = [Convert]::ToBase64String($calculatedHash)
            $authorization = 'SharedKey {0}:{1}' -f $customerId, $encodedHash
            return $authorization
        }
        
        # Create the function to create and post the request
        Function Post-LogAnalyticsData($customerId, $sharedKey, $body, $logType) {
            $method = "POST"
            $contentType = "application/json"
            $resource = "/api/logs"
            $rfc1123date = [DateTime]::UtcNow.ToString("r")
            $contentLength = $body.Length
            $signature = Build-Signature `
                -customerId $customerId `
                -sharedKey $sharedKey `
                -date $rfc1123date `
                -contentLength $contentLength `
                -method $method `
                -contentType $contentType `
                -resource $resource
            
            $uri = "https://" + $customerId + ".ods.opinsights.azure.com" + $resource + "?api-version=2016-04-01"
            
            $headers = @{
                "Authorization"        = $signature;
                "Log-Type"             = $logType;
                "x-ms-date"            = $rfc1123date;
                "time-generated-field" = "Timestamp";
            }
            
            $response = Invoke-WebRequest -Uri $uri -Method $method -ContentType $contentType -Headers $headers -Body $body -UseBasicParsing
            return $response.StatusCode
        }
        
        # Submit the data
        $jsonData = $Data | ConvertTo-Json
        Write-Output "Sending data to Log Analytics..."
        $statusCode = Post-LogAnalyticsData -customerId $WorkspaceId -sharedKey $WorkspaceKey -body ([System.Text.Encoding]::UTF8.GetBytes($jsonData)) -logType $LogName
        
        if ($statusCode -eq 200) {
            Write-Output "Data successfully sent to Log Analytics"
        } else {
            Write-Error "Error sending data to Log Analytics: StatusCode = $statusCode"
        }
    }
    catch {
        Write-Error "Error sending data to Log Analytics: $_"
        throw
    }
}

# Main execution
try {
    Write-Output "Starting SharePoint storage monitoring script..."
    
    # Get credentials from Key Vault
    $credential = Get-SecureCredentials -TenantId $config.TenantId -ClientId $config.ClientId -KeyVaultName $config.KeyVaultName -SecretName $config.SecretName
    
    # Get SharePoint storage statistics
    $storageData = Get-SharePointStorageStats -TenantId $config.TenantId -Credential $credential -SiteUrls $config.SharePointSites
    
    # Send data to Log Analytics
    if ($storageData.Count -gt 0) {
        Send-LogAnalyticsData -WorkspaceId $config.WorkspaceId -WorkspaceKey $config.WorkspaceKey -LogName $config.LogName -Data $storageData
        
        # Output the collected data for reference
        Write-Output "Storage data collected:"
        $storageData | Format-Table -AutoSize
    } else {
        Write-Warning "No storage data was collected."
    }
    
    Write-Output "Script execution completed successfully."
}
catch {
    Write-Error "Error in main execution: $_"
    throw
}