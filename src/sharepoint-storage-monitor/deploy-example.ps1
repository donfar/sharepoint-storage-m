# Example deployment script for SharePoint Storage Monitor
# Modify the parameters below to match your environment

# Parameters for deployment
$params = @{
    ResourceGroupName = "SharePointMonitor-RG"      # Name of the resource group to create or use
    Location = "eastus"                            # Azure region for resource deployment
    FunctionAppName = "sp-storage-monitor"          # Name of the Azure Function App
    KeyVaultName = "sp-monitor-kv"                 # Name of the Azure Key Vault
    WorkspaceId = ""                               # Log Analytics workspace ID (leave empty to create new)
    WorkspaceKey = ""                              # Log Analytics workspace key (leave empty to create new)
    LogName = "SharePointStorageStats"              # Name of the log in Log Analytics
    TenantsJsonPath = "./config.json"              # Path to the tenants configuration file
    CreateDashboard = $true                        # Create an Azure dashboard
}

# Run the deployment script with the parameters
$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "deploy-complete.ps1"
& $scriptPath @params