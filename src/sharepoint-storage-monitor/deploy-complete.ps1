# SharePoint Storage Monitor - Complete Deployment Script
# This script automates the entire deployment process for the SharePoint Storage Monitor application

param (
    [Parameter(Mandatory=$true)]
    [string]$ResourceGroupName,
    
    [Parameter(Mandatory=$true)]
    [string]$Location,
    
    [Parameter(Mandatory=$true)]
    [string]$FunctionAppName,
    
    [Parameter(Mandatory=$true)]
    [string]$KeyVaultName,
    
    [Parameter(Mandatory=$true)]
    [string]$WorkspaceId,
    
    [Parameter(Mandatory=$true)]
    [string]$WorkspaceKey,
    
    [Parameter(Mandatory=$false)]
    [string]$LogName = "SharePointStorageStats",
    
    [Parameter(Mandatory=$false)]
    [string]$TenantsJsonPath,

    [Parameter(Mandatory=$false)]
    [switch]$CreateDashboard = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipModuleCheck = $false
)

# Function to ensure required modules are installed
function Ensure-Module {
    param (
        [string]$ModuleName,
        [string]$MinimumVersion = ""
    )
    
    Write-Host "Checking for module: $ModuleName"
    
    # Check if module is already installed
    if ($MinimumVersion -eq "") {
        $moduleInstalled = Get-Module -Name $ModuleName -ListAvailable
    } else {
        $moduleInstalled = Get-Module -Name $ModuleName -ListAvailable | Where-Object { $_.Version -ge $MinimumVersion }
    }
    
    # Install module if not present
    if (-not $moduleInstalled) {
        try {
            Write-Host "Installing module: $ModuleName"
            if ($MinimumVersion -eq "") {
                Install-Module -Name $ModuleName -Force -Scope CurrentUser -AllowClobber
            } else {
                Install-Module -Name $ModuleName -Force -MinimumVersion $MinimumVersion -Scope CurrentUser -AllowClobber
            }
            Write-Host "$ModuleName module installed successfully"
        }
        catch {
            Write-Error "Failed to install $ModuleName module: $_"
            throw
        }
    } else {
        Write-Host "$ModuleName module already installed"
    }
}

# Function to create and deploy the package
function Create-FunctionPackage {
    param (
        [string]$ScriptPath,
        [string]$OutputPath
    )
    
    Write-Host "Creating function package from $ScriptPath..."
    
    # Ensure output directory exists
    if (-not (Test-Path -Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath | Out-Null
    }
    
    # Copy required files to package directory
    Copy-Item -Path "$ScriptPath\sharepoint-storage-monitor.ps1" -Destination "$OutputPath\sharepoint-storage-monitor.ps1"
    Copy-Item -Path "$ScriptPath\function.json" -Destination "$OutputPath\function.json"
    
    # Create profile.ps1 for module initialization
    @"
# Azure Functions profile.ps1
# This file runs when the Function App starts and initializes the environment
# It's a good place to set up any prerequisites that the function needs

# Install or import modules here
if (`$env:MSI_ENDPOINT) {
    # Running in Azure
    Write-Output "Function App starting up in Azure environment..."
} else {
    # Running locally
    Write-Output "Function App starting up in local environment..."
}
"@ | Out-File -FilePath "$OutputPath\profile.ps1"
    
    # If configuration file exists, copy it
    if (Test-Path -Path "$ScriptPath\config.json") {
        Copy-Item -Path "$ScriptPath\config.json" -Destination "$OutputPath\config.json"
    } elseif ($TenantsJsonPath -and (Test-Path -Path $TenantsJsonPath)) {
        Copy-Item -Path $TenantsJsonPath -Destination "$OutputPath\config.json"
    }
    
    # Create zip package
    $zipPath = "$ScriptPath\function-package.zip"
    if (Test-Path -Path $zipPath) {
        Remove-Item -Path $zipPath -Force
    }
    
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($OutputPath, $zipPath)
    
    return $zipPath
}

# Function to deploy the dashboard
function Deploy-Dashboard {
    param (
        [string]$ResourceGroupName,
        [string]$DashboardName,
        [string]$Location,
        [string]$WorkspaceId,
        [string]$DashboardTemplatePath
    )
    
    Write-Host "Deploying dashboard $DashboardName..."
    
    if (Test-Path $DashboardTemplatePath) {
        $dashboardJson = Get-Content -Path $DashboardTemplatePath -Raw
        
        # Replace placeholder values
        $dashboardJson = $dashboardJson -replace '__WORKSPACE_ID__', $WorkspaceId
        
        # Create temporary file with processed JSON
        $tempFile = [System.IO.Path]::GetTempFileName()
        $dashboardJson | Out-File -FilePath $tempFile -Encoding utf8
        
        # Deploy the dashboard
        New-AzPortalDashboard -ResourceGroupName $ResourceGroupName -DashboardPath $tempFile -Name $DashboardName -Location $Location
        
        # Cleanup
        Remove-Item -Path $tempFile -Force
        
        Write-Host "Dashboard deployed successfully"
    } else {
        Write-Warning "Dashboard template file not found: $DashboardTemplatePath"
    }
}

# Start deployment process
Write-Host "Starting SharePoint Storage Monitor deployment..."

# Verify script is running in correct location
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

# Check for required modules
if (-not $SkipModuleCheck) {
    Write-Host "Checking for required modules..."
    Ensure-Module -ModuleName "Az.Accounts"
    Ensure-Module -ModuleName "Az.Resources" 
    Ensure-Module -ModuleName "Az.Functions"
    Ensure-Module -ModuleName "Az.KeyVault"
    Ensure-Module -ModuleName "Az.OperationalInsights"
    Ensure-Module -ModuleName "Az.Storage"
    Ensure-Module -ModuleName "Az.Portal"
    Ensure-Module -ModuleName "Az.Websites"
}

# Login to Azure if not already logged in
$context = Get-AzContext
if (-not $context) {
    Write-Host "Please log in to Azure..."
    Connect-AzAccount
    $context = Get-AzContext
}

Write-Host "Connected to Azure subscription: $($context.Subscription.Name)"

# Create or ensure resource group exists
Write-Host "Creating resource group $ResourceGroupName if it doesn't exist..."
New-AzResourceGroup -Name $ResourceGroupName -Location $Location -Force | Out-Null

# Create Log Analytics workspace if specified but not provided
$workspaceName = "$FunctionAppName-workspace"
if (-not $WorkspaceId -or -not $WorkspaceKey) {
    Write-Host "Creating Log Analytics workspace $workspaceName..."
    $workspace = New-AzOperationalInsightsWorkspace -ResourceGroupName $ResourceGroupName -Name $workspaceName -Location $Location -Force
    $WorkspaceId = $workspace.CustomerId.ToString()
    $WorkspaceKey = (Get-AzOperationalInsightsWorkspaceSharedKey -ResourceGroupName $ResourceGroupName -Name $workspaceName).PrimarySharedKey
    Write-Host "Log Analytics workspace created with ID: $WorkspaceId"
}

# Create Key Vault if it doesn't exist
Write-Host "Creating Key Vault $KeyVaultName if it doesn't exist..."
$keyVault = Get-AzKeyVault -VaultName $KeyVaultName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
if (-not $keyVault) {
    $keyVault = New-AzKeyVault -VaultName $KeyVaultName -ResourceGroupName $ResourceGroupName -Location $Location
    Write-Host "Key Vault created: $($keyVault.VaultName)"
} else {
    Write-Host "Key Vault already exists: $($keyVault.VaultName)"
}

# Create storage account for Function App
$storageName = ($FunctionAppName -replace '-', '' -replace '_', '').ToLower()
if ($storageName.Length -gt 24) {
    $storageName = $storageName.Substring(0, 24)
}

Write-Host "Creating storage account $storageName..."
$storage = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $storageName -ErrorAction SilentlyContinue
if (-not $storage) {
    $storage = New-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $storageName -Location $Location -SkuName Standard_LRS
    Write-Host "Storage account created: $($storage.StorageAccountName)"
} else {
    Write-Host "Storage account already exists: $($storage.StorageAccountName)"
}

# Create App Service Plan for Function App
$appServicePlanName = "$FunctionAppName-plan"
Write-Host "Creating App Service Plan $appServicePlanName..."
$appServicePlan = Get-AzAppServicePlan -ResourceGroupName $ResourceGroupName -Name $appServicePlanName -ErrorAction SilentlyContinue
if (-not $appServicePlan) {
    $appServicePlan = New-AzAppServicePlan -ResourceGroupName $ResourceGroupName -Name $appServicePlanName -Location $Location -Tier Consumption -WorkerType Dynamic
    Write-Host "App Service Plan created: $($appServicePlan.Name)"
} else {
    Write-Host "App Service Plan already exists: $($appServicePlan.Name)"
}

# Create Function App
Write-Host "Creating Function App $FunctionAppName..."
$functionApp = Get-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -ErrorAction SilentlyContinue
if (-not $functionApp) {
    $functionApp = New-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -Location $Location `
        -StorageAccountName $storageName -Runtime PowerShell -RuntimeVersion 7.2 -OSType Windows -FunctionsVersion 4 `
        -AppServicePlan $appServicePlanName
    Write-Host "Function App created: $($functionApp.Name)"
} else {
    Write-Host "Function App already exists: $($functionApp.Name)"
}

# Set Function App settings
Write-Host "Setting Function App settings..."
$appSettings = @{
    "WEBSITE_RUN_FROM_PACKAGE" = "1"
    "KeyVaultName" = $KeyVaultName
    "WorkspaceId" = $WorkspaceId
    "WorkspaceKey" = $WorkspaceKey
    "LogName" = $LogName
    "MaxRetries" = "3"
    "RetryDelaySeconds" = "5" 
    "LogErrors" = "true"
    "ErrorLogName" = "SharePointStorageErrors"
}

# Process tenants
if ($TenantsJsonPath -and (Test-Path -Path $TenantsJsonPath)) {
    # Load the tenants configuration from the file
    Write-Host "Loading tenants configuration from $TenantsJsonPath..."
    $tenantsConfig = Get-Content -Path $TenantsJsonPath -Raw | ConvertFrom-Json
    
    # Store the full tenants configuration in the function app
    $appSettings["TenantsList"] = $tenantsConfig | ConvertTo-Json -Depth 10 -Compress
    
    # Process each tenant
    foreach ($tenant in $tenantsConfig.Tenants) {
        $secretName = $tenant.SecretName
        $secretValue = Read-Host -Prompt "Enter client secret for tenant $($tenant.TenantName)" -AsSecureString
        
        Write-Host "Storing secret $secretName in Key Vault for tenant $($tenant.TenantName)..."
        Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name $secretName -SecretValue $secretValue | Out-Null
    }
}

# Update Function App settings
Write-Host "Updating Function App settings..."
Update-AzFunctionAppSetting -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -AppSetting $appSettings

# Ensure Function App has a system-assigned identity
Write-Host "Ensuring Function App has a system-assigned identity..."
Update-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -IdentityType SystemAssigned

# Wait for identity to propagate
Write-Host "Waiting for identity to propagate..."
Start-Sleep -Seconds 10

# Grant Function App access to the Key Vault
Write-Host "Granting Function App access to the Key Vault..."
$functionAppIdentity = Get-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName | Select-Object -ExpandProperty IdentityPrincipalId
Set-AzKeyVaultAccessPolicy -VaultName $KeyVaultName -ObjectId $functionAppIdentity -PermissionsToSecrets get,list

# Create and deploy the function package
$tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.Guid]::NewGuid().ToString())
$packagePath = Create-FunctionPackage -ScriptPath $scriptRoot -OutputPath $tempDir

# Deploy the package to the function app
Write-Host "Deploying function package to the Function App..."
Publish-AzWebapp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -ArchivePath $packagePath -Force

# Clean up temporary files
if (Test-Path -Path $tempDir) {
    Remove-Item -Path $tempDir -Recurse -Force
}
if (Test-Path -Path $packagePath) {
    Remove-Item -Path $packagePath -Force
}

# Deploy dashboard if requested
if ($CreateDashboard) {
    $dashboardName = "$FunctionAppName-dashboard"
    $dashboardTemplatePath = Join-Path -Path $scriptRoot -ChildPath "dashboard-template.json"
    
    if (Test-Path -Path $dashboardTemplatePath) {
        Deploy-Dashboard -ResourceGroupName $ResourceGroupName -DashboardName $dashboardName -Location $Location -WorkspaceId $WorkspaceId -DashboardTemplatePath $dashboardTemplatePath
    } else {
        Write-Warning "Dashboard template not found. Skipping dashboard creation."
    }
}

# Check if function is deployed and running
Write-Host "Checking if function is deployed and running..."
$functionKey = $null
$retryCount = 0
$maxRetries = 10

while ($retryCount -lt $maxRetries -and -not $functionKey) {
    try {
        $functionKey = Get-AzFunctionAppHostKey -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -ErrorAction SilentlyContinue
        if ($functionKey) {
            Write-Host "Function app is up and running."
            break
        }
    } catch {
        Write-Host "Waiting for function app to become ready..."
    }
    
    $retryCount++
    Start-Sleep -Seconds 10
}

# Display summary information
Write-Host "`n===== DEPLOYMENT COMPLETE =====`n"
Write-Host "SharePoint Storage Monitor has been successfully deployed to Azure!"
Write-Host ""
Write-Host "Resource Group:       $ResourceGroupName"
Write-Host "Function App:         $FunctionAppName"
Write-Host "Function App URL:     https://$FunctionAppName.azurewebsites.net"
Write-Host "Log Analytics:        $workspaceName"
Write-Host "Key Vault:            $KeyVaultName"
if ($CreateDashboard) {
    Write-Host "Dashboard:            $dashboardName"
}
Write-Host ""
Write-Host "Next Steps:"
Write-Host "1. Access your Function App in the Azure Portal to view logs and monitor execution."
Write-Host "2. The function will run automatically based on the timer trigger (default: daily at midnight)."
Write-Host "3. View your storage data in Log Analytics under the '$LogName' log."
if ($CreateDashboard) {
    Write-Host "4. Access your dashboard in the Azure Portal to visualize the data."
} else {
    Write-Host "4. Consider creating a dashboard in the Azure Portal to visualize the data."
}
Write-Host ""
Write-Host "For troubleshooting and more information, visit the SharePoint Storage Monitor documentation."