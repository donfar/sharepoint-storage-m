# SharePoint Storage Monitor - Azure Deployment Script
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
    [string]$SharePointClientSecret,
    
    [Parameter(Mandatory=$true)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$true)]
    [string]$SharePointSites,
    
    [Parameter(Mandatory=$false)]
    [string]$LogAnalyticsWorkspaceName = "$FunctionAppName-workspace"
)

# Function to ensure required modules are installed
function Ensure-Module {
    param (
        [string]$ModuleName
    )
    
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Installing module: $ModuleName"
        Install-Module -Name $ModuleName -Force -Scope CurrentUser
    } else {
        Write-Host "Module $ModuleName is already installed"
    }
}

# Ensure Azure modules are installed
Ensure-Module -ModuleName "Az.Accounts"
Ensure-Module -ModuleName "Az.Resources"
Ensure-Module -ModuleName "Az.Functions"
Ensure-Module -ModuleName "Az.KeyVault"
Ensure-Module -ModuleName "Az.OperationalInsights"
Ensure-Module -ModuleName "Az.Storage"

# Login to Azure if not already logged in
$context = Get-AzContext
if (-not $context) {
    Connect-AzAccount
    $context = Get-AzContext
}

Write-Host "Connected to Azure: $($context.Subscription.Name)"

# Create Resource Group if it doesn't exist
$rg = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
if (-not $rg) {
    Write-Host "Creating resource group: $ResourceGroupName"
    New-AzResourceGroup -Name $ResourceGroupName -Location $Location
} else {
    Write-Host "Resource group $ResourceGroupName already exists"
}

# Create Log Analytics workspace
Write-Host "Creating Log Analytics workspace: $LogAnalyticsWorkspaceName"
$workspace = New-AzOperationalInsightsWorkspace -ResourceGroupName $ResourceGroupName -Name $LogAnalyticsWorkspaceName -Location $Location -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
if (-not $workspace) {
    $workspace = Get-AzOperationalInsightsWorkspace -ResourceGroupName $ResourceGroupName -Name $LogAnalyticsWorkspaceName
}
$workspaceId = $workspace.CustomerId
$workspaceKey = (Get-AzOperationalInsightsWorkspaceSharedKey -ResourceGroupName $ResourceGroupName -Name $LogAnalyticsWorkspaceName).PrimarySharedKey

# Create Key Vault
Write-Host "Creating Key Vault: $KeyVaultName"
$keyVault = New-AzKeyVault -Name $KeyVaultName -ResourceGroupName $ResourceGroupName -Location $Location -EnabledForTemplateDeployment -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
if (-not $keyVault) {
    $keyVault = Get-AzKeyVault -ResourceGroupName $ResourceGroupName -VaultName $KeyVaultName
}

# Store the SharePoint client secret in Key Vault
$secretName = "SharePointClientSecret"
Write-Host "Adding secret to Key Vault: $secretName"
$secret = ConvertTo-SecureString -String $SharePointClientSecret -AsPlainText -Force
Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name $secretName -SecretValue $secret

# Create storage account for Function App
$storageAccountName = $FunctionAppName.ToLower() -replace '[^a-z0-9]', ''
if ($storageAccountName.Length -gt 24) {
    $storageAccountName = $storageAccountName.Substring(0, 24)
}
Write-Host "Creating storage account: $storageAccountName"
$storageAccount = New-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $storageAccountName -Location $Location -SkuName Standard_LRS -Kind StorageV2 -ErrorAction SilentlyContinue
if (-not $storageAccount) {
    $storageAccount = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $storageAccountName
}

# Create Function App
Write-Host "Creating Function App: $FunctionAppName"
$functionApp = New-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -Location $Location -StorageAccountName $storageAccountName -Runtime PowerShell -RuntimeVersion 7.2 -FunctionsVersion 4 -ErrorAction SilentlyContinue
if (-not $functionApp) {
    $functionApp = Get-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName
}

# Set Function App settings
Write-Host "Setting Function App settings"
Update-AzFunctionAppSetting -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -AppSetting @{
    "TenantId" = $TenantId
    "ClientId" = $ClientId
    "KeyVaultName" = $KeyVaultName
    "SecretName" = $secretName
    "WorkspaceId" = $workspaceId
    "WorkspaceKey" = $workspaceKey
    "SharePointSites" = $SharePointSites
} | Out-Null

# Assign Managed Identity to Function App and grant permissions to Key Vault
Write-Host "Setting up Managed Identity for Function App"
Update-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -IdentityType SystemAssigned | Out-Null
Start-Sleep -Seconds 10  # Wait for the identity to propagate

# Get the Function App's managed identity object ID
$functionAppIdentity = (Get-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName).IdentityPrincipalId

# Assign Key Vault permissions to Function App's managed identity
Write-Host "Granting Key Vault access to Function App's managed identity"
Set-AzKeyVaultAccessPolicy -VaultName $KeyVaultName -ObjectId $functionAppIdentity -PermissionsToSecrets Get,List

Write-Host "Deployment completed successfully!"
Write-Host ""
Write-Host "=== NEXT STEPS ==="
Write-Host "1. Deploy the PowerShell script to your Azure Function App"
Write-Host "2. Create a dashboard in Azure Monitor to visualize the data"
Write-Host ""
Write-Host "Function App URL: https://$FunctionAppName.azurewebsites.net"
Write-Host "Log Analytics Workspace ID: $workspaceId"
Write-Host ""