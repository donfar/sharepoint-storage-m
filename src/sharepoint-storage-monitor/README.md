# SharePoint Storage Monitor - Multi-tenant Edition

This PowerShell application runs as an Azure Function to collect and visualize SharePoint storage utilization across multiple tenants. The solution sends all data to Azure Log Analytics for comprehensive monitoring and visualization.

## Features

- **Multi-tenant Support**: Monitor SharePoint sites across multiple Microsoft 365 tenants
- **Centralized Monitoring**: Aggregate data from all tenants into a single Azure Log Analytics workspace
- **Secure Credential Management**: Safely store and access credentials for each tenant in Azure Key Vault
- **Comprehensive Visualization**: Pre-built queries for cross-tenant reporting and analysis
- **Automated Error Handling**: Robust error management to ensure reliability across tenants
- **Modular Configuration**: Easy-to-maintain tenant configuration via JSON

## Prerequisites

- Azure subscription
- PowerShell 7.2 or higher (for local testing)
- Azure Function App with PowerShell runtime
- Azure Key Vault for secure credential storage
- Azure Log Analytics workspace
- SharePoint App Registration in each tenant with Sites.Read.All permissions

## Getting Started

### 1. Set Up App Registrations in Each Tenant

For each tenant you want to monitor:

1. Sign in to the Azure portal for that tenant
2. Navigate to Azure Active Directory > App registrations
3. Create a new registration
4. Grant API permissions: SharePoint > Sites.Read.All
5. Create a client secret and save the value
6. Note the Application (client) ID and Directory (tenant) ID

### 2. Prepare Configuration

Create a `config.json` file based on the provided `config-sample.json`:

```json
{
  "WorkspaceId": "your-log-analytics-workspace-id",
  "WorkspaceKey": "your-log-analytics-workspace-key",
  "LogName": "SharePointStorageStats",
  "KeyVaultName": "your-keyvault-name",
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

### 3. Deploy to Azure

Use the provided deployment script:

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

### 4. Deploy the Function App Code

```powershell
# Compress the files
Compress-Archive -Path *.ps1,*.json -DestinationPath function.zip

# Deploy to Azure Function App
Publish-AzWebapp -ResourceGroupName "SharePointMonitor-RG" `
  -Name "sharepoint-storage-monitor" `
  -ArchivePath function.zip
```

### 5. Test the Function

You can trigger the function manually in the Azure portal or wait for the scheduled execution.

## Local Testing

1. Update `local.settings.json` with your values
2. Install Azure Functions Core Tools
3. Run the function locally:

```powershell
func start
```

## Dashboard Queries

The `dashboard-queries.kql` file contains sample queries for visualizing data across multiple tenants in Log Analytics. Key visualizations include:

- Storage usage by tenant (pie chart)
- Storage utilization over time by tenant (line chart)
- Top 10 sites by storage across all tenants (table)
- Growth rate by tenant (30 days)
- Site count by tenant
- Storage utilization heatmap by tenant and site

## Security Considerations

- The solution uses a managed identity for the Function App to access Key Vault
- Tenant credentials are stored as secrets in Key Vault
- Each tenant requires a separate app registration to limit scope of permissions
- Log Analytics secure configuration is maintained through connection keys

## Troubleshooting

### Authentication Issues

- Verify that client secrets haven't expired
- Check that app registrations have the necessary permissions
- Ensure admin consent was granted for the Sites.Read.All permission in each tenant

### Missing Data

- Check Function App logs for errors
- Verify Log Analytics workspace ID and key
- Ensure SharePoint sites are accessible to the respective app identities

### Function Execution Problems

- Verify that the PowerShell modules can be installed
- Check for network connectivity issues
- Ensure the Function App has internet access

## Contributing

Feel free to submit pull requests to enhance the functionality of this solution.

## License

This project is licensed under the MIT License - see the LICENSE file for details.