# SharePoint Storage Monitor - CI/CD Prerequisites Setup Script
# This script creates the required resources and service principals for CI/CD deployment

param (
    [Parameter(Mandatory=$true)]
    [string]$SubscriptionId,
    
    [Parameter(Mandatory=$true)]
    [string]$ResourceGroupName,
    
    [Parameter(Mandatory=$false)]
    [string]$ServicePrincipalName = "sp-sharepoint-monitor-cicd",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputJsonPath = "./sp-credentials.json"
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

# Select the subscription
Write-Host "Setting subscription context to: $SubscriptionId"
Select-AzSubscription -SubscriptionId $SubscriptionId

# Create or ensure resource group exists
Write-Host "Creating resource group $ResourceGroupName if it doesn't exist..."
$rg = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
if (-not $rg) {
    Write-Host "Resource group doesn't exist. Please specify a location to create it:"
    $location = Read-Host "Enter Azure region (e.g., eastus)"
    $rg = New-AzResourceGroup -Name $ResourceGroupName -Location $location
    Write-Host "Resource group created: $($rg.ResourceGroupName) in $($rg.Location)"
} else {
    Write-Host "Using existing resource group: $($rg.ResourceGroupName) in $($rg.Location)"
}

# Create service principal with Contributor access to the resource group
Write-Host "Creating service principal for CI/CD: $ServicePrincipalName"
$sp = Get-AzADServicePrincipal -DisplayName $ServicePrincipalName -ErrorAction SilentlyContinue
if (-not $sp) {
    $sp = New-AzADServicePrincipal -DisplayName $ServicePrincipalName
    Write-Host "Service principal created with ID: $($sp.Id)"
    
    # Create a password/secret for the service principal
    $credentials = New-AzADSpCredential -ObjectId $sp.Id
    $clientSecret = $credentials.SecretText
    
    # Store the original password for JSON output
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($clientSecret)
    $clearClientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
} else {
    Write-Host "Service principal already exists with ID: $($sp.Id)"
    # Create a new password/secret for the service principal
    $credentials = New-AzADSpCredential -ObjectId $sp.Id
    $clientSecret = $credentials.SecretText
    
    # Store the password for JSON output
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($clientSecret)
    $clearClientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    Write-Host "Created new secret for existing service principal"
}

# Assign Contributor role to the service principal for the resource group
Write-Host "Assigning Contributor role to service principal for resource group $ResourceGroupName"
$assignment = Get-AzRoleAssignment -ObjectId $sp.Id -RoleDefinitionName Contributor -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
if (-not $assignment) {
    New-AzRoleAssignment -ObjectId $sp.Id -RoleDefinitionName Contributor -ResourceGroupName $ResourceGroupName
    Write-Host "Role assignment created"
} else {
    Write-Host "Role assignment already exists"
}

# Create the output JSON for GitHub Actions
$outputJson = @{
    clientId = $sp.AppId
    clientSecret = $clearClientSecret
    tenantId = (Get-AzContext).Tenant.Id
    subscriptionId = $SubscriptionId
} | ConvertTo-Json

# Save the JSON to a file
$outputJson | Out-File -FilePath $OutputJsonPath -Force
Write-Host "Service principal credentials saved to: $OutputJsonPath"

# Format the output for Azure DevOps service connection
Write-Host "`n`nFor Azure DevOps Service Connection, use these values:"
Write-Host "=============================================="
Write-Host "Subscription ID: $SubscriptionId"
Write-Host "Subscription Name: $((Get-AzSubscription -SubscriptionId $SubscriptionId).Name)"
Write-Host "Service Principal ID: $($sp.AppId)"
Write-Host "Service Principal Key: [Saved in $OutputJsonPath]"
Write-Host "Tenant ID: $((Get-AzContext).Tenant.Id)"
Write-Host "Service Connection Name: serviceconnection-dev"
Write-Host "Resource Group: $ResourceGroupName"
Write-Host "=============================================="

# Instructions for GitHub Actions
Write-Host "`n`nFor GitHub Actions, create a secret named AZURE_CREDENTIALS with the following value:"
Write-Host "=============================================="
Write-Host $outputJson
Write-Host "=============================================="

Write-Host "`n`nSetup completed successfully!"
Write-Host "IMPORTANT: Store the generated credentials securely and delete the credentials file after use."
Write-Host "The service principal has Contributor access to resource group: $ResourceGroupName"