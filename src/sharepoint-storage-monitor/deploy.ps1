# SharePoint Storage Monitor - Multi-tenant Deployment Script
# This script deploys the SharePoint Storage Monitor to Azure

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
    [string]$TenantsJsonPath
)

# Ensure Az module is installed
if (-not (Get-Module -ListAvailable -Name Az)) {
    Write-Host "Installing Az module..."
    Install-Module -Name Az -Scope CurrentUser -Repository PSGallery -Force
}

# Login to Azure if not already logged in
$context = Get-AzContext
if (-not $context) {
    Connect-AzAccount
}

# Create or ensure resource group exists
Write-Host "Creating resource group $ResourceGroupName if it doesn't exist..."
New-AzResourceGroup -Name $ResourceGroupName -Location $Location -Force | Out-Null

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
}

# Process tenants
if ($TenantsJsonPath) {
    # Load the tenants configuration from the file
    $tenantsConfig = Get-Content -Path $TenantsJsonPath -Raw | ConvertFrom-Json
    
    # Store the full tenants configuration in the function app
    $appSettings["TenantsList"] = $tenantsConfig | ConvertTo-Json -Depth 10 -Compress
    
    # Process each tenant
    foreach ($tenant in $tenantsConfig) {
        $secretName = $tenant.SecretName
        $secretValue = Read-Host -Prompt "Enter client secret for tenant $($tenant.TenantName)" -AsSecureString
        
        Write-Host "Storing secret $secretName in Key Vault for tenant $($tenant.TenantName)..."
        Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name $secretName -SecretValue $secretValue | Out-Null
        
        # Grant Function App access to the Key Vault
        Write-Host "Granting Function App access to the Key Vault for tenant $($tenant.TenantName)..."
        $functionAppIdentity = Get-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName | Select-Object -ExpandProperty IdentityPrincipalId
        Set-AzKeyVaultAccessPolicy -VaultName $KeyVaultName -ObjectId $functionAppIdentity -PermissionsToSecrets get,list
    }
} else {
    Write-Host "No tenants configuration file provided. You'll need to manually set up the tenant configuration."
}

# Update Function App settings
Update-AzFunctionAppSetting -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -AppSetting $appSettings

# Ensure Function App has a system-assigned identity
Write-Host "Ensuring Function App has a system-assigned identity..."
Update-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -IdentityType SystemAssigned

# Grant Function App access to the Key Vault
Write-Host "Granting Function App access to the Key Vault..."
$functionAppIdentity = Get-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName | Select-Object -ExpandProperty IdentityPrincipalId
Set-AzKeyVaultAccessPolicy -VaultName $KeyVaultName -ObjectId $functionAppIdentity -PermissionsToSecrets get,list

Write-Host "Deployment completed successfully!"
Write-Host "Next steps:"
Write-Host "1. Compress the PowerShell scripts and function configuration files"
Write-Host "2. Deploy the package to the function app using 'Publish-AzWebapp' or through the Azure Portal"
Write-Host "3. Verify that the function app is running correctly by checking the logs"