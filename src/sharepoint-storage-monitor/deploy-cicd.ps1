# CI/CD Deployment Script for SharePoint Storage Monitor
# This script is designed to be used in CI/CD pipelines (Azure DevOps, GitHub Actions)

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
    [switch]$CreateDashboard = $false,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantsList,
    
    # Key Vault secrets can be passed as secure strings, base64-encoded strings, or loaded from external files
    [Parameter(Mandatory=$false)]
    [string]$SecretsFilePath,
    
    [Parameter(Mandatory=$false)]
    [string]$ServicePrincipalId,
    
    [Parameter(Mandatory=$false)]
    [string]$ServicePrincipalKey,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId
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

# Ensure Az modules are installed
Ensure-Module -ModuleName "Az.Accounts"
Ensure-Module -ModuleName "Az.Resources"
Ensure-Module -ModuleName "Az.Functions"
Ensure-Module -ModuleName "Az.KeyVault"
Ensure-Module -ModuleName "Az.OperationalInsights"
Ensure-Module -ModuleName "Az.Storage"
Ensure-Module -ModuleName "Az.Portal"
Ensure-Module -ModuleName "Az.Websites"

# Login to Azure with Service Principal if credentials are provided
if ($ServicePrincipalId -and $ServicePrincipalKey -and $TenantId) {
    Write-Host "Logging in to Azure with Service Principal..."
    $securePassword = ConvertTo-SecureString -String $ServicePrincipalKey -AsPlainText -Force
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ServicePrincipalId, $securePassword
    Connect-AzAccount -ServicePrincipal -Credential $credential -TenantId $TenantId
} else {
    # Assume managed identity or already logged in
    Write-Host "Using existing Azure authentication context or managed identity..."
}

# Process the TenantsList parameter into a temporary config file if provided
$tempConfigPath = $null
if ($TenantsList) {
    Write-Host "Processing tenant list..."
    $tempConfigPath = [System.IO.Path]::GetTempFileName()
    $TenantsList | Out-File -FilePath $tempConfigPath
    $TenantsJsonPath = $tempConfigPath
}

# Process secrets file if provided
$secretsObject = $null
if ($SecretsFilePath -and (Test-Path $SecretsFilePath)) {
    Write-Host "Loading secrets from file..."
    $secretsContent = Get-Content -Path $SecretsFilePath -Raw
    $secretsObject = ConvertFrom-Json -InputObject $secretsContent
}

# Create temporary deployment directory
$tempDeploymentDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.Guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $tempDeploymentDir -Force | Out-Null

# Copy all necessary files to the deployment directory
Copy-Item -Path "$PSScriptRoot\sharepoint-storage-monitor.ps1" -Destination "$tempDeploymentDir\sharepoint-storage-monitor.ps1"
Copy-Item -Path "$PSScriptRoot\function.json" -Destination "$tempDeploymentDir\function.json"
if ($TenantsJsonPath -and (Test-Path $TenantsJsonPath)) {
    Copy-Item -Path $TenantsJsonPath -Destination "$tempDeploymentDir\config.json"
}

# Create the resource group if it doesn't exist
Get-AzResourceGroup -Name $ResourceGroupName -ErrorVariable notPresent -ErrorAction SilentlyContinue
if ($notPresent) {
    Write-Host "Creating resource group $ResourceGroupName..."
    New-AzResourceGroup -Name $ResourceGroupName -Location $Location
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

# Create Key Vault if it doesn't exist
Write-Host "Creating Key Vault $KeyVaultName if it doesn't exist..."
$keyVault = Get-AzKeyVault -VaultName $KeyVaultName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
if (-not $keyVault) {
    $keyVault = New-AzKeyVault -VaultName $KeyVaultName -ResourceGroupName $ResourceGroupName -Location $Location
    Write-Host "Key Vault created: $($keyVault.VaultName)"
} else {
    Write-Host "Key Vault already exists: $($keyVault.VaultName)"
}

# Create Function App
Write-Host "Creating Function App $FunctionAppName..."
$functionApp = Get-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -ErrorAction SilentlyContinue
if (-not $functionApp) {
    Write-Host "Creating Function App..."
    $functionApp = New-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -Location $Location `
        -StorageAccountName $storageName -Runtime PowerShell -RuntimeVersion 7.2 -OSType Windows -FunctionsVersion 4
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

# If we have a config file, load and process tenants
if (Test-Path "$tempDeploymentDir\config.json") {
    $config = Get-Content -Path "$tempDeploymentDir\config.json" -Raw | ConvertFrom-Json
    $appSettings["TenantsList"] = $config | ConvertTo-Json -Depth 10 -Compress
    
    # Process secrets if we have them and tenant configuration
    if ($secretsObject -and $config.Tenants) {
        foreach ($tenant in $config.Tenants) {
            $secretName = $tenant.SecretName
            $secretProperty = $secretName -replace "-", "_"
            
            if ($secretsObject.$secretProperty) {
                Write-Host "Setting secret $secretName for tenant $($tenant.TenantName)..."
                $secretValue = ConvertTo-SecureString -String $secretsObject.$secretProperty -AsPlainText -Force
                Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name $secretName -SecretValue $secretValue | Out-Null
            } else {
                Write-Warning "Secret not found for $($tenant.TenantName) tenant!"
            }
        }
    }
}

# Update Function App settings
Update-AzFunctionAppSetting -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -AppSetting $appSettings

# Ensure Function App has a system-assigned identity
Update-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -IdentityType SystemAssigned

# Wait for identity to propagate
Start-Sleep -Seconds 10

# Grant Function App access to the Key Vault
$functionAppIdentity = Get-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName | Select-Object -ExpandProperty IdentityPrincipalId
Set-AzKeyVaultAccessPolicy -VaultName $KeyVaultName -ObjectId $functionAppIdentity -PermissionsToSecrets get,list

# Package the function app
$zipPath = Join-Path -Path $tempDeploymentDir -ChildPath "function-package.zip"
$currentLocation = Get-Location
Set-Location -Path $tempDeploymentDir
Compress-Archive -Path * -DestinationPath $zipPath -Force
Set-Location -Path $currentLocation

# Deploy the package to the function app
Write-Host "Deploying function package to the Function App..."
Publish-AzWebapp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -ArchivePath $zipPath -Force

# Deploy dashboard if requested
if ($CreateDashboard) {
    $dashboardName = "$FunctionAppName-dashboard"
    $dashboardTemplatePath = Join-Path -Path $PSScriptRoot -ChildPath "dashboard-template.json"
    
    if (Test-Path -Path $dashboardTemplatePath) {
        Write-Host "Deploying dashboard $dashboardName..."
        $dashboardJson = Get-Content -Path $dashboardTemplatePath -Raw
        
        # Replace placeholder values
        $dashboardJson = $dashboardJson -replace '__WORKSPACE_ID__', $WorkspaceId
        
        # Create temporary file with processed JSON
        $tempFile = Join-Path -Path $tempDeploymentDir -ChildPath "dashboard.json"
        $dashboardJson | Out-File -FilePath $tempFile -Encoding utf8
        
        # Deploy the dashboard
        New-AzPortalDashboard -ResourceGroupName $ResourceGroupName -DashboardPath $tempFile -Name $dashboardName -Location $Location
        
        Write-Host "Dashboard deployed successfully"
    } else {
        Write-Warning "Dashboard template file not found: $dashboardTemplatePath"
    }
}

# Clean up temporary files
if (Test-Path -Path $tempDeploymentDir) {
    Remove-Item -Path $tempDeploymentDir -Recurse -Force
}
if ($tempConfigPath -and (Test-Path -Path $tempConfigPath)) {
    Remove-Item -Path $tempConfigPath -Force
}

Write-Host "CI/CD Deployment completed successfully!"