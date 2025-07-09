# SharePoint Storage Monitor Deployment Guide

This guide explains how to deploy the SharePoint Storage Monitor to Azure using the automated deployment script.

## Prerequisites

Before you begin, ensure you have the following:

1. **PowerShell 7.2 or higher** installed on your machine
2. **Azure PowerShell modules** installed:
   ```powershell
   Install-Module -Name Az -Scope CurrentUser -Repository PSGallery -Force
   ```
3. **Azure subscription** with appropriate permissions to create resources
4. **App registrations** in each tenant you want to monitor, with `Sites.Read.All` permissions
5. **Client secrets** for each app registration
6. **Configuration file** for your tenants (see below)

## Configuration File

Create a `config.json` file with the following structure:

```json
{
  "WorkspaceId": "your-log-analytics-workspace-id",
  "WorkspaceKey": "your-log-analytics-workspace-key",
  "LogName": "SharePointStorageStats",
  "KeyVaultName": "your-keyvault-name",
  "MaxRetries": 3,
  "RetryDelaySeconds": 5,
  "LogErrors": true,
  "ErrorLogName": "SharePointStorageErrors",
  "Tenants": [
    {
      "TenantId": "tenant1-id",
      "TenantName": "Contoso",
      "ClientId": "app-registration-client-id-for-tenant1",
      "SecretName": "contoso-sharepoint-app-secret",
      "SharePointSites": [
        "contoso.sharepoint.com/sites/site1",
        "contoso.sharepoint.com/sites/site2"
      ]
    },
    {
      "TenantId": "tenant2-id",
      "TenantName": "Fabrikam",
      "ClientId": "app-registration-client-id-for-tenant2",
      "SecretName": "fabrikam-sharepoint-app-secret",
      "SharePointSites": [
        "fabrikam.sharepoint.com/sites/marketing",
        "fabrikam.sharepoint.com/sites/sales"
      ]
    }
  ]
}
```

## Deployment Options

### Option 1: Full Automated Deployment

The `deploy-complete.ps1` script handles the entire deployment process:

1. Creates Azure resources (Resource Group, Function App, Key Vault, etc.)
2. Securely stores client secrets in Key Vault
3. Packages and deploys the function code
4. Optionally creates a dashboard for visualization

```powershell
# Example usage with all required parameters
./deploy-complete.ps1 `
  -ResourceGroupName "SharePointMonitor-RG" `
  -Location "eastus" `
  -FunctionAppName "sharepoint-storage-monitor" `
  -KeyVaultName "sp-monitor-kv" `
  -WorkspaceId "your-log-analytics-workspace-id" `
  -WorkspaceKey "your-log-analytics-workspace-key" `
  -TenantsJsonPath "./config.json" `
  -CreateDashboard
```

### Option 2: Step-by-Step Manual Deployment

If you prefer to deploy each component separately:

#### 1. Create Infrastructure

```powershell
./deploy.ps1 `
  -ResourceGroupName "SharePointMonitor-RG" `
  -Location "eastus" `
  -FunctionAppName "sharepoint-storage-monitor" `
  -KeyVaultName "sp-monitor-kv" `
  -WorkspaceId "your-log-analytics-workspace-id" `
  -WorkspaceKey "your-log-analytics-workspace-key" `
  -TenantsJsonPath "./config.json"
```

#### 2. Package and Deploy Function Code

```powershell
# Create a ZIP package
Compress-Archive -Path ./sharepoint-storage-monitor/* -DestinationPath function.zip -Force

# Deploy to Azure Function App
Publish-AzWebapp -ResourceGroupName "SharePointMonitor-RG" -Name "sharepoint-storage-monitor" -ArchivePath ./function.zip
```

## Parameters

| Parameter | Description | Required |
|-----------|-------------|----------|
| ResourceGroupName | Name of the resource group to create or use | Yes |
| Location | Azure region for resource deployment | Yes |
| FunctionAppName | Name of the Azure Function App | Yes |
| KeyVaultName | Name of the Azure Key Vault | Yes |
| WorkspaceId | Log Analytics workspace ID | Yes |
| WorkspaceKey | Log Analytics workspace key | Yes |
| LogName | Name of the log in Log Analytics | No (default: "SharePointStorageStats") |
| TenantsJsonPath | Path to the tenants configuration file | No, but highly recommended |
| CreateDashboard | Switch to create an Azure dashboard | No (default: false) |
| SkipModuleCheck | Skip checking and installing required PowerShell modules | No (default: false) |

## Post-Deployment Steps

1. **Verify Function App Deployment**:
   - Check the Function App in the Azure Portal
   - View the logs to ensure it's running correctly

2. **Test the Function**:
   - Manually trigger the function to verify data collection
   - Check Log Analytics for the collected data

3. **Set Up Monitoring and Alerts**:
   - Create alerts for storage thresholds
   - Configure notification emails for alerts

4. **Review Dashboard**:
   - If you created a dashboard, customize it as needed
   - Add additional visualizations or queries

## Troubleshooting

If you encounter issues during deployment:

1. **Function App Not Running**:
   - Check the function logs in the Azure Portal
   - Verify the system-assigned identity has access to Key Vault

2. **Authentication Errors**:
   - Verify client secrets are correct and not expired
   - Ensure app registrations have the required permissions

3. **Missing Data**:
   - Check if SharePoint sites are accessible
   - Verify the correct tenant IDs and site URLs in the configuration

4. **Deployment Script Errors**:
   - Ensure you have the latest Az PowerShell modules
   - Check if you have sufficient permissions in your Azure subscription

For more detailed troubleshooting steps, see the main documentation.

## Security Considerations

- Client secrets are stored securely in Azure Key Vault
- The Function App uses managed identity to access Key Vault
- No secrets are stored in code or configuration files
- Only required permissions are granted to the Function App
- Log Analytics data is encrypted at rest

## Additional Resources

- [Azure Functions Documentation](https://docs.microsoft.com/en-us/azure/azure-functions/)
- [Azure Key Vault Documentation](https://docs.microsoft.com/en-us/azure/key-vault/)
- [Log Analytics Documentation](https://docs.microsoft.com/en-us/azure/azure-monitor/logs/log-analytics-overview)
- [SharePoint REST API Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)