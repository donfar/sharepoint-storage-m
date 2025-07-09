# SharePoint Storage Monitor

A PowerShell application that runs as an Azure Function to track and visualize SharePoint storage utilization over time using Azure Monitor and Log Analytics.

## Features

- Automated daily collection of SharePoint site storage metrics
- Secure credential management using Azure Key Vault
- Comprehensive error handling and logging
- Visualization through Azure Monitor dashboards
- Automatic PowerShell module management

## Prerequisites

- Azure subscription
- SharePoint Online tenant with appropriate permissions
- SharePoint App Registration with Sites.Read.All permissions
- PowerShell 7.2 or higher (for local testing)

## Deployment Instructions

### 1. Register an Azure AD Application for SharePoint Access

1. Go to Azure Active Directory > App registrations
2. Create a new registration
3. Grant API permissions: SharePoint > Sites.Read.All
4. Create a client secret and save the value (you'll need it for deployment)

### 2. Deploy Azure Resources

Run the deployment script with the following parameters:

```powershell
./deploy.ps1 `
  -ResourceGroupName "SharePointMonitor-RG" `
  -Location "eastus" `
  -FunctionAppName "sharepoint-storage-monitor" `
  -KeyVaultName "sp-monitor-kv" `
  -SharePointClientSecret "your-client-secret" `
  -TenantId "your-tenant-id" `
  -ClientId "your-app-registration-client-id" `
  -SharePointSites "tenant.sharepoint.com/sites/site1,tenant.sharepoint.com/sites/site2"
```

### 3. Deploy the Function App Code

1. Clone this repository
2. Navigate to the `src` directory
3. Deploy to Azure Function App:

```powershell
Compress-Archive -Path *.ps1,*.json -DestinationPath function.zip
Publish-AzWebapp -ResourceGroupName "SharePointMonitor-RG" -Name "sharepoint-storage-monitor" -ArchivePath function.zip
```

## Local Testing

To test locally before deployment:

1. Install Azure Functions Core Tools
2. Update `local.settings.json` with your values
3. Run the function locally:

```powershell
func start
```

## Creating a Dashboard in Azure Monitor

1. Go to Azure Monitor > Dashboards
2. Create a new dashboard
3. Add a Log Analytics query widget
4. Use this sample query:

```kusto
SharePointStorageStats_CL
| project TimeGenerated, SiteUrl_s, StorageUsed_d, StorageQuota_d, PercentageUsed_d
| render timechart
```

## Troubleshooting

Check Azure Function logs for detailed error information:

1. Go to your Function App in Azure Portal
2. Navigate to Functions > SharePointStorageMonitor > Monitor
3. Review the logs for any error messages

Common issues:
- Missing permissions for SharePoint access
- Incorrect client secret or expired credentials
- Network connectivity issues between Azure and SharePoint

## Maintenance

- Periodically check for expired client secrets in Azure AD
- Review Azure Function execution logs
- Update PowerShell modules when newer versions are available

## License

This project is licensed under the MIT License - see the LICENSE file for details.