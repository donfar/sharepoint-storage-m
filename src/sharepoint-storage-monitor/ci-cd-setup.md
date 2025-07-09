# CI/CD Setup for SharePoint Storage Monitor

This guide provides instructions for setting up continuous integration and continuous deployment (CI/CD) pipelines for the SharePoint Storage Monitor using either GitHub Actions or Azure DevOps.

## Prerequisites

- A SharePoint Storage Monitor codebase in a Git repository
- Azure subscription with appropriate permissions
- Azure Log Analytics workspace
- Azure Key Vault for secure credential storage

## GitHub Actions Setup

### 1. Configure GitHub Secrets

Add the following secrets to your GitHub repository:

- `AZURE_CREDENTIALS`: JSON output from az ad sp create-for-rbac command with Contributor access to your Azure subscription
- `RESOURCE_GROUP`: Target Azure resource group name
- `LOCATION`: Azure region (e.g., eastus)
- `KEY_VAULT_NAME`: Name of your Azure Key Vault
- `WORKSPACE_ID`: Log Analytics workspace ID
- `WORKSPACE_KEY`: Log Analytics workspace key
- `TENANT_CONFIG`: JSON string containing tenant configuration (example below)

Example TENANT_CONFIG format:
```json
[
  {
    "TenantId": "tenant1-id",
    "TenantName": "Contoso",
    "ClientId": "app-registration-client-id-for-tenant1",
    "SecretName": "contoso-sharepoint-app-secret",
    "SecretValue": "app-client-secret-value",
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
    "SecretValue": "app-client-secret-value",
    "SharePointSites": [
      "fabrikam.sharepoint.com/sites/marketing",
      "fabrikam.sharepoint.com/sites/sales"
    ]
  }
]
```

### 2. Set Up Repository Structure

Ensure your repository has the following structure:
```
src/
  sharepoint-storage-monitor/
    .github/
      workflows/
        azure-deploy.yml
    deploy.ps1
    sharepoint-storage-monitor.ps1
    function.json
    ...other files
```

### 3. Run the Workflow

The workflow will run automatically when you push changes to the main branch that affect files in the `src/sharepoint-storage-monitor/` directory.

You can also manually trigger the workflow from the Actions tab in your GitHub repository.

## Azure DevOps Setup

### 1. Create Azure DevOps Pipeline

1. Create a new pipeline in Azure DevOps
2. Select your repository
3. Choose "Existing Azure Pipelines YAML file"
4. Select the path to `src/sharepoint-storage-monitor/azure-pipelines.yml`

### 2. Configure Pipeline Variables

Create the following pipeline variables:

- `keyVaultName`: Name of your Azure Key Vault
- `workspaceId`: Log Analytics workspace ID
- `workspaceKey`: Log Analytics workspace key (mark as secret)
- `tenantConfig`: JSON string containing tenant configuration (mark as secret, same format as GitHub example above)

### 3. Set Up Service Connections

Create service connections in Azure DevOps for each environment:

1. Go to Project Settings > Service connections
2. Create a new Azure Resource Manager service connection
3. Name it `serviceconnection-dev` (also create test/prod as needed)
4. Select your Azure subscription and resource group

### 4. Create Environments

Create environments in Azure DevOps for each deployment target:

1. Go to Pipelines > Environments
2. Create environments named `dev`, `test`, and `prod`
3. Configure approval requirements for test and prod environments if needed

### 5. Run the Pipeline

The pipeline will run automatically when you push changes to the main branch that affect files in the `src/sharepoint-storage-monitor/` directory.

You can also manually trigger the pipeline from the Pipelines section in Azure DevOps.

## Security Considerations

- **Secrets Management**: All sensitive information is stored in Azure Key Vault or pipeline secrets
- **Service Principal Permissions**: Use the principle of least privilege when creating service principals
- **Environment Isolation**: Use separate resource groups and service principals for each environment
- **Secret Rotation**: Implement a process for rotating secrets regularly

## Validation Steps

After deployment, verify:

1. The Function App is deployed and running
2. Key Vault contains all the required secrets
3. Function App has access to Key Vault
4. Test the function by triggering it manually
5. Check Log Analytics for data collection