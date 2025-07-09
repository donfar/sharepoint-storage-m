# SharePoint Storage Monitor
# This script collects SharePoint storage utilization from multiple tenants and sends data to Azure Log Analytics
#
# Created: 2023-09-01
# Author: GitHub Spark

# Get the current script path
$scriptPath = $PSScriptRoot

# Import configuration if exists
$configPath = Join-Path -Path $scriptPath -ChildPath "config.json"
if (Test-Path $configPath) {
    $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
} else {
    # Default configuration (environment variables are used as fallback)
    $config = @{
        WorkspaceId = $env:WorkspaceId
        WorkspaceKey = $env:WorkspaceKey
        LogName = $env:LogName ?? "SharePointStorageStats"
        KeyVaultName = $env:KeyVaultName
        Tenants = @()
    }

    # If tenants are provided via environment variable
    if ($env:TenantsList) {
        try {
            $config.Tenants = $env:TenantsList | ConvertFrom-Json
        } catch {
            Write-Error "Failed to parse TenantsList environment variable: $_"
        }
    } else {
        # Single tenant fallback
        if ($env:TenantId -and $env:ClientId) {
            $config.Tenants = @(
                @{
                    TenantId = $env:TenantId
                    TenantName = $env:TenantName ?? "Default"
                    ClientId = $env:ClientId
                    SecretName = $env:SecretName
                    SharePointSites = $env:SharePointSites -split ","
                }
            )
        }
    }
}

# Function to check and install modules
function Ensure-Module {
    param (
        [string]$ModuleName,
        [string]$MinimumVersion = ""
    )
    
    Write-Verbose "Checking for module: $ModuleName"
    $module = Get-Module -Name $ModuleName -ListAvailable
    
    if (-not $module) {
        Write-Verbose "Module $ModuleName is not installed, installing now..."
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
            Write-Verbose "Module $ModuleName has been installed successfully"
        } catch {
            throw "Failed to install module $ModuleName. Error: $_"
        }
    }
    elseif ($MinimumVersion -and ($module.Version -lt [System.Version]::Parse($MinimumVersion))) {
        Write-Verbose "Module $ModuleName version $($module.Version) is below minimum required version $MinimumVersion, updating..."
        try {
            Update-Module -Name $ModuleName -Force
            Write-Verbose "Module $ModuleName has been updated successfully"
        } catch {
            throw "Failed to update module $ModuleName. Error: $_"
        }
    }
    else {
        Write-Verbose "Module $ModuleName is already installed with version $($module.Version)"
    }
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
        # Ensure Azure modules are installed
        Ensure-Module -ModuleName "Az.Accounts"
        Ensure-Module -ModuleName "Az.KeyVault"
        
        # Connect to Azure
        Write-Verbose "Connecting to Azure with managed identity"
        Connect-AzAccount -Identity -Tenant $TenantId | Out-Null
        
        # Get secret from Key Vault
        Write-Verbose "Getting secret $SecretName from Key Vault $KeyVaultName"
        $secret = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name $SecretName
        
        if (-not $secret) {
            throw "Secret $SecretName not found in Key Vault $KeyVaultName"
        }
        
        $secretText = $secret.SecretValue | ConvertFrom-SecureString -AsPlainText
        
        # Create credential object
        $password = ConvertTo-SecureString -String $secretText -AsPlainText -Force
        $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $password
        
        return $credential
    }
    catch {
        throw "Error getting credentials from Key Vault: $_"
    }
}

# Function to get SharePoint storage statistics
function Get-SharePointStorageStats {
    param (
        [string]$TenantId,
        [string]$TenantName,
        [System.Management.Automation.PSCredential]$Credential,
        [string[]]$SiteUrls
    )
    
    try {
        # Ensure PnP module is installed
        Ensure-Module -ModuleName "PnP.PowerShell" -MinimumVersion "1.12.0"
        
        $results = @()
        
        foreach ($siteUrl in $SiteUrls) {
            try {
                Write-Verbose "Connecting to SharePoint site: $siteUrl"
                Connect-PnPOnline -Url $siteUrl -Credentials $Credential
                
                # Get site information
                $site = Get-PnPSite -Includes StorageMaximumLevel, StorageUsage, StorageWarningLevel
                $web = Get-PnPWeb -Includes Title
                
                # Format results for Log Analytics
                $resultObject = [PSCustomObject]@{
                    TenantId = $TenantId
                    TenantName = $TenantName
                    SiteUrl = $siteUrl
                    SiteTitle = $web.Title
                    StorageUsed = [math]::Round($site.StorageUsage / 1024, 2)  # Convert to MB
                    StorageLimit = [math]::Round($site.StorageMaximumLevel / 1024, 2)  # Convert to MB
                    StorageWarning = [math]::Round($site.StorageWarningLevel / 1024, 2)  # Convert to MB
                    PercentageUsed = [math]::Round(($site.StorageUsage / $site.StorageMaximumLevel) * 100, 2)
                    CollectionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
                
                $results += $resultObject
                Write-Verbose "Collected data for site: $siteUrl"
            }
            catch {
                Write-Error "Error getting stats for site $siteUrl in tenant $TenantName`: $_"
            }
            finally {
                # Disconnect from SharePoint site
                Disconnect-PnPOnline
            }
        }
        
        return $results
    }
    catch {
        throw "Error in Get-SharePointStorageStats for tenant $TenantName`: $_"
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
    
    if (-not $Data -or $Data.Count -eq 0) {
        Write-Verbose "No data to send to Log Analytics"
        return
    }
    
    try {
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
                "Authorization" = $signature
                "Log-Type" = $logType
                "x-ms-date" = $rfc1123date
                "time-generated-field" = "CollectionDate"
            }
            
            $response = Invoke-WebRequest -Uri $uri -Method $method -ContentType $contentType -Headers $headers -Body $body -UseBasicParsing
            return $response.StatusCode
        }
        
        # Submit the data
        $bodyJson = $Data | ConvertTo-Json
        $statusCode = Post-LogAnalyticsData -customerId $WorkspaceId -sharedKey $WorkspaceKey -body $bodyJson -logType $LogName
        
        Write-Verbose "Log Analytics status code: $statusCode"
        if ($statusCode -ne 200) {
            Write-Warning "Error submitting data to Log Analytics: Status code $statusCode"
        } else {
            Write-Verbose "Successfully sent $($Data.Count) records to Log Analytics"
        }
    }
    catch {
        throw "Error sending data to Log Analytics: $_"
    }
}

# Main execution
try {
    Write-Output "Starting SharePoint Storage Monitor for multiple tenants"
    $totalResults = @()
    
    # Verify configuration
    if (-not $config.Tenants -or $config.Tenants.Count -eq 0) {
        throw "No tenants configured. Please provide tenant information in config.json or via environment variables."
    }
    
    if (-not $config.WorkspaceId -or -not $config.WorkspaceKey) {
        throw "Log Analytics workspace ID and key are required."
    }
    
    # Process each tenant
    foreach ($tenant in $config.Tenants) {
        try {
            Write-Output "Processing tenant: $($tenant.TenantName)"
            
            # Get credentials from Key Vault
            $credential = Get-SecureCredentials -TenantId $tenant.TenantId -ClientId $tenant.ClientId -KeyVaultName $config.KeyVaultName -SecretName $tenant.SecretName
            
            # Get SharePoint storage statistics
            $storageData = Get-SharePointStorageStats -TenantId $tenant.TenantId -TenantName $tenant.TenantName -Credential $credential -SiteUrls $tenant.SharePointSites
            
            if ($storageData.Count -gt 0) {
                $totalResults += $storageData
                Write-Output "Collected data from $($storageData.Count) sites in tenant $($tenant.TenantName)"
            } else {
                Write-Warning "No data collected from tenant $($tenant.TenantName)"
            }
        }
        catch {
            Write-Error "Error processing tenant $($tenant.TenantName): $_"
            # Continue with next tenant
        }
    }
    
    # Send all data to Log Analytics
    if ($totalResults.Count -gt 0) {
        Send-LogAnalyticsData -WorkspaceId $config.WorkspaceId -WorkspaceKey $config.WorkspaceKey -LogName $config.LogName -Data $totalResults
        Write-Output "Successfully sent $($totalResults.Count) total records to Log Analytics"
    } else {
        Write-Warning "No data was collected from any tenant"
    }
}
catch {
    Write-Error "Error in main execution: $_"
    throw
}