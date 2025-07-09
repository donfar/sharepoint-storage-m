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
        # Error handling configuration
        MaxRetries = $env:MaxRetries ?? 3
        RetryDelaySeconds = $env:RetryDelaySeconds ?? 5
        LogErrors = $env:LogErrors -eq "true"
        ErrorLogName = $env:ErrorLogName
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
        [string]$SecretName,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 5
    )
    
    $retryCount = 0
    $success = $false
    $lastError = $null
    
    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            # Ensure Azure modules are installed
            Ensure-Module -ModuleName "Az.Accounts"
            Ensure-Module -ModuleName "Az.KeyVault"
            
            # Connect to Azure
            Write-Verbose "Connecting to Azure with managed identity (Attempt $($retryCount + 1))"
            Connect-AzAccount -Identity -Tenant $TenantId -ErrorAction Stop | Out-Null
            
            # Get secret from Key Vault
            Write-Verbose "Getting secret $SecretName from Key Vault $KeyVaultName"
            $secret = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name $SecretName -ErrorAction Stop
            
            if (-not $secret) {
                throw "Secret $SecretName not found in Key Vault $KeyVaultName"
            }
            
            $secretText = $secret.SecretValue | ConvertFrom-SecureString -AsPlainText
            
            # Create credential object
            $password = ConvertTo-SecureString -String $secretText -AsPlainText -Force
            $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $password
            
            $success = $true
            return $credential
        }
        catch {
            $lastError = $_
            $retryCount++
            
            if ($retryCount -lt $MaxRetries) {
                $errorType = if ($_.Exception.GetType().Name) { $_.Exception.GetType().Name } else { "Unknown" }
                Write-Warning "Attempt $retryCount of $MaxRetries failed when getting credentials for tenant $TenantId. Error type: $errorType. Retrying in $RetryDelaySeconds seconds..."
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }
    }
    
    if (-not $success) {
        throw "Failed to get credentials after $MaxRetries attempts from Key Vault for tenant $TenantId. Last error: $lastError"
    }
}

# Function to get SharePoint storage statistics
function Get-SharePointStorageStats {
    param (
        [string]$TenantId,
        [string]$TenantName,
        [System.Management.Automation.PSCredential]$Credential,
        [string[]]$SiteUrls,
        [int]$MaxSiteRetries = 2,
        [int]$RetryDelaySeconds = 10
    )
    
    try {
        # Ensure PnP module is installed
        Ensure-Module -ModuleName "PnP.PowerShell" -MinimumVersion "1.12.0"
        
        $results = @()
        $siteErrors = @()
        $successCount = 0
        $failureCount = 0
        
        foreach ($siteUrl in $SiteUrls) {
            $retryCount = 0
            $siteSuccess = $false
            $lastError = $null
            
            while (-not $siteSuccess -and $retryCount -le $MaxSiteRetries) {
                try {
                    if ($retryCount -gt 0) {
                        Write-Verbose "Retry attempt $retryCount for site: $siteUrl"
                    }
                    
                    Write-Verbose "Connecting to SharePoint site: $siteUrl"
                    Connect-PnPOnline -Url $siteUrl -Credentials $Credential -ErrorAction Stop
                    
                    # Get site information with timeout handling
                    $timeoutTask = [System.Threading.Tasks.Task]::Run({
                        try {
                            $site = Get-PnPSite -Includes StorageMaximumLevel, StorageUsage, StorageWarningLevel -ErrorAction Stop
                            $web = Get-PnPWeb -Includes Title -ErrorAction Stop
                            return @{ Site = $site; Web = $web }
                        }
                        catch {
                            throw $_
                        }
                    })
                    
                    # Wait for the task with a timeout
                    if (-not [System.Threading.Tasks.Task]::WaitAll(@($timeoutTask), 30000)) {
                        throw "Operation timed out after 30 seconds when getting site information"
                    }
                    
                    $siteData = $timeoutTask.Result
                    $site = $siteData.Site
                    $web = $siteData.Web
                    
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
                        Status = "Success"
                    }
                    
                    $results += $resultObject
                    $successCount++
                    $siteSuccess = $true
                    Write-Verbose "Successfully collected data for site: $siteUrl"
                }
                catch {
                    $lastError = $_
                    $retryCount++
                    
                    # Categorize the error
                    $errorCategory = "Unknown"
                    $errorMessage = $_.Exception.Message
                    
                    if ($errorMessage -match "Access denied|Unauthorized|401|403") {
                        $errorCategory = "AccessDenied"
                    }
                    elseif ($errorMessage -match "timeout|timed out") {
                        $errorCategory = "Timeout"
                    }
                    elseif ($errorMessage -match "not found|404|doesn't exist") {
                        $errorCategory = "NotFound"
                    }
                    elseif ($errorMessage -match "network|connectivity|connection") {
                        $errorCategory = "NetworkIssue"
                    }
                    
                    if ($retryCount -le $MaxSiteRetries) {
                        Write-Warning "Attempt $retryCount of $MaxSiteRetries failed for site $siteUrl in tenant $TenantName. Error category: $errorCategory. Retrying in $RetryDelaySeconds seconds..."
                        Start-Sleep -Seconds $RetryDelaySeconds
                    }
                }
                finally {
                    # Always try to disconnect, but don't throw if it fails
                    try {
                        Disconnect-PnPOnline -ErrorAction SilentlyContinue
                    }
                    catch {
                        # Just continue if disconnect fails
                    }
                }
            }
            
            if (-not $siteSuccess) {
                $failureCount++
                
                # Add error info to results for tracking
                $errorResult = [PSCustomObject]@{
                    TenantId = $TenantId
                    TenantName = $TenantName
                    SiteUrl = $siteUrl
                    SiteTitle = "Error"
                    CollectionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Status = "Failed"
                    ErrorMessage = $lastError.Exception.Message
                }
                
                $siteErrors += $errorResult
                Write-Error "Failed to collect data for site $siteUrl in tenant $TenantName after $MaxSiteRetries retries. Error: $lastError"
            }
        }
        
        # Log statistics
        Write-Verbose "Tenant $TenantName processing complete. Success: $successCount, Failed: $failureCount"
        
        # If we have site errors, we can log them for monitoring
        if ($siteErrors.Count -gt 0) {
            Write-Verbose "Returning $($results.Count) successful results and $($siteErrors.Count) error records"
            return $results, $siteErrors
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
    $totalErrors = @()
    $tenantResults = @{}
    
    # Verify configuration
    if (-not $config.Tenants -or $config.Tenants.Count -eq 0) {
        throw "No tenants configured. Please provide tenant information in config.json or via environment variables."
    }
    
    if (-not $config.WorkspaceId -or -not $config.WorkspaceKey) {
        throw "Log Analytics workspace ID and key are required."
    }
    
    # Initialize result tracking
    $tenantStats = @{
        TotalTenants = $config.Tenants.Count
        ProcessedTenants = 0
        SuccessfulTenants = 0
        FailedTenants = 0
        TotalSitesAttempted = 0
        TotalSitesSuccessful = 0
        TenantDetails = @{}
    }
    
    # Process each tenant
    foreach ($tenant in $config.Tenants) {
        $tenantStats.ProcessedTenants++
        $tenantId = $tenant.TenantId
        $tenantName = $tenant.TenantName
        
        # Initialize tenant tracking
        $tenantStats.TenantDetails[$tenantName] = @{
            Status = "Processing"
            SitesAttempted = $tenant.SharePointSites.Count
            SitesSuccessful = 0
            ErrorType = ""
            ErrorMessage = ""
            StartTime = Get-Date
        }
        
        try {
            Write-Output "Processing tenant: $tenantName (Tenant $($tenantStats.ProcessedTenants) of $($tenantStats.TotalTenants))"
            
            # Track sites for this tenant
            $tenantStats.TotalSitesAttempted += $tenant.SharePointSites.Count
            
            # Get credentials from Key Vault with retry logic
            Write-Output "  Getting credentials for tenant: $tenantName"
            $credential = $null
            try {
                $credential = Get-SecureCredentials -TenantId $tenantId -ClientId $tenant.ClientId -KeyVaultName $config.KeyVaultName -SecretName $tenant.SecretName
            }
            catch {
                $errorMsg = "Failed to get credentials for tenant $tenantName`: $_"
                Write-Error $errorMsg
                $tenantStats.TenantDetails[$tenantName].Status = "Failed"
                $tenantStats.TenantDetails[$tenantName].ErrorType = "Credentials"
                $tenantStats.TenantDetails[$tenantName].ErrorMessage = $errorMsg
                $tenantStats.FailedTenants++
                continue
            }
            
            # Get SharePoint storage statistics
            Write-Output "  Collecting storage data from $($tenant.SharePointSites.Count) sites in tenant: $tenantName"
            $storageResults = Get-SharePointStorageStats -TenantId $tenantId -TenantName $tenantName -Credential $credential -SiteUrls $tenant.SharePointSites
            
            # Handle different return types (normal array vs array with errors)
            $storageData = @()
            $errorData = @()
            
            if ($storageResults -is [array] -and $storageResults.Count -eq 2) {
                # We have a split return with [results, errors]
                $storageData = $storageResults[0]
                $errorData = $storageResults[1]
            }
            else {
                # Just regular results
                $storageData = $storageResults
            }
            
            # Update tenant statistics
            $tenantStats.TenantDetails[$tenantName].SitesSuccessful = $storageData.Count
            $tenantStats.TotalSitesSuccessful += $storageData.Count
            
            # Process results
            if ($storageData.Count -gt 0) {
                $totalResults += $storageData
                $tenantResults[$tenantName] = $storageData
                $tenantStats.SuccessfulTenants++
                $tenantStats.TenantDetails[$tenantName].Status = "Success"
                Write-Output "  Collected data from $($storageData.Count) sites in tenant $tenantName"
                
                # If we had some errors but also some successes
                if ($errorData.Count -gt 0) {
                    Write-Warning "  Failed to collect data from $($errorData.Count) sites in tenant $tenantName"
                    $totalErrors += $errorData
                }
            } 
            else {
                Write-Warning "  No data collected from tenant $tenantName"
                $tenantStats.TenantDetails[$tenantName].Status = "NoData"
                
                if ($errorData.Count -gt 0) {
                    Write-Warning "  Failed to collect data from all $($errorData.Count) sites in tenant $tenantName"
                    $totalErrors += $errorData
                    $tenantStats.FailedTenants++
                }
            }
        }
        catch {
            $errorMsg = "Error processing tenant $tenantName`: $_"
            Write-Error $errorMsg
            $tenantStats.FailedTenants++
            $tenantStats.TenantDetails[$tenantName].Status = "Failed"
            $tenantStats.TenantDetails[$tenantName].ErrorType = "Processing"
            $tenantStats.TenantDetails[$tenantName].ErrorMessage = $errorMsg
            
            # Create an error entry for logging
            $errorEntry = [PSCustomObject]@{
                TenantId = $tenantId
                TenantName = $tenantName
                CollectionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Status = "TenantFailed"
                ErrorMessage = $_.Exception.Message
            }
            $totalErrors += $errorEntry
        }
        finally {
            $tenantStats.TenantDetails[$tenantName].EndTime = Get-Date
            $tenantStats.TenantDetails[$tenantName].Duration = [math]::Round(($tenantStats.TenantDetails[$tenantName].EndTime - $tenantStats.TenantDetails[$tenantName].StartTime).TotalSeconds, 1)
        }
    }
    
    # Send all success data to Log Analytics
    if ($totalResults.Count -gt 0) {
        Write-Output "Sending $($totalResults.Count) data records to Log Analytics"
        Send-LogAnalyticsData -WorkspaceId $config.WorkspaceId -WorkspaceKey $config.WorkspaceKey -LogName $config.LogName -Data $totalResults
    } else {
        Write-Warning "No successful data was collected from any tenant"
    }
    
    # Send error data to Log Analytics if configured
    if ($totalErrors.Count -gt 0 -and $config.LogErrors -ne $false) {
        $errorLogName = $config.ErrorLogName ?? ($config.LogName + "Errors")
        Write-Output "Sending $($totalErrors.Count) error records to Log Analytics as '$errorLogName'"
        Send-LogAnalyticsData -WorkspaceId $config.WorkspaceId -WorkspaceKey $config.WorkspaceKey -LogName $errorLogName -Data $totalErrors
    }
    
    # Output summary
    Write-Output ""
    Write-Output "--- Execution Summary ---"
    Write-Output "Tenants processed: $($tenantStats.ProcessedTenants) of $($tenantStats.TotalTenants)"
    Write-Output "Successful tenants: $($tenantStats.SuccessfulTenants)"
    Write-Output "Failed tenants: $($tenantStats.FailedTenants)"
    Write-Output "Total sites attempted: $($tenantStats.TotalSitesAttempted)"
    Write-Output "Total sites successful: $($tenantStats.TotalSitesSuccessful)"
    Write-Output ""
    
    # Output per-tenant results
    Write-Output "--- Tenant Results ---"
    foreach ($tn in $tenantStats.TenantDetails.Keys) {
        $t = $tenantStats.TenantDetails[$tn]
        $statusIcon = switch ($t.Status) {
            "Success" { "✓" }
            "Failed" { "✗" }
            "NoData" { "⚠" }
            default { "?" }
        }
        
        Write-Output "$statusIcon $tn`: $($t.SitesSuccessful) of $($t.SitesAttempted) sites processed in $($t.Duration)s"
        if ($t.Status -eq "Failed") {
            Write-Output "   Error: $($t.ErrorMessage)"
        }
    }
}
catch {
    Write-Error "Critical error in main execution: $_"
    throw
}