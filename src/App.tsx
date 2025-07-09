import { useState } from "react";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Separator } from "@/components/ui/separator";
import { CodeBlock } from "./components/code-block";
import { FileIcon, Database, LineChart, Terminal, Code, FileCode, Copy, Info, Buildings } from "@phosphor-icons/react";

function App() {
  const [activeTab, setActiveTab] = useState("overview");
  
  const handleCopyFile = async (content: string) => {
    try {
      await navigator.clipboard.writeText(content);
      // Could use toast here if integrated
      console.log("Copied to clipboard");
    } catch (err) {
      console.error("Failed to copy:", err);
    }
  };

  return (
    <div className="min-h-screen bg-background p-4 sm:p-6 md:p-8">
      <header className="mb-8 space-y-2">
        <h1 className="text-3xl font-bold tracking-tight text-foreground">SharePoint Storage Monitor</h1>
        <p className="text-muted-foreground text-lg">
          Multi-tenant PowerShell application that runs as an Azure Function to track and visualize SharePoint storage utilization across environments
        </p>
      </header>

      <Tabs defaultValue="overview" value={activeTab} onValueChange={setActiveTab} className="space-y-6">
        <TabsList className="grid w-full grid-cols-3 md:grid-cols-7">
          <TabsTrigger value="overview" className="flex items-center gap-2">
            <Info weight="duotone" /> Overview
          </TabsTrigger>
          <TabsTrigger value="multi-tenant" className="flex items-center gap-2">
            <Buildings weight="duotone" /> Multi-tenant
          </TabsTrigger>
          <TabsTrigger value="code" className="flex items-center gap-2">
            <FileCode weight="duotone" /> Code
          </TabsTrigger>
          <TabsTrigger value="deployment" className="flex items-center gap-2">
            <Terminal weight="duotone" /> Deployment
          </TabsTrigger>
          <TabsTrigger value="cicd" className="flex items-center gap-2">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"></circle><path d="M12 16v-4"></path><path d="M12 8h.01"></path></svg>
            CI/CD
          </TabsTrigger>
          <TabsTrigger value="dashboard" className="hidden md:flex items-center gap-2">
            <LineChart weight="duotone" /> Dashboard
          </TabsTrigger>
          <TabsTrigger value="troubleshooting" className="hidden md:flex items-center gap-2">
            <Code weight="duotone" /> Troubleshooting
          </TabsTrigger>
        </TabsList>
        
        <TabsContent value="overview">
          <div className="grid gap-6 md:grid-cols-2">
            <Card>
              <CardHeader>
                <CardTitle>Features</CardTitle>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="flex items-start gap-3">
                  <div className="mt-1 bg-primary/10 text-primary rounded-md p-1.5">
                    <Buildings size={20} weight="duotone" />
                  </div>
                  <div>
                    <h3 className="font-medium text-lg">Multi-tenant Support</h3>
                    <p className="text-muted-foreground">Monitor multiple SharePoint environments in a single solution</p>
                  </div>
                </div>
                
                <div className="flex items-start gap-3">
                  <div className="mt-1 bg-primary/10 text-primary rounded-md p-1.5">
                    <Terminal size={20} weight="duotone" />
                  </div>
                  <div>
                    <h3 className="font-medium text-lg">Automated Collection</h3>
                    <p className="text-muted-foreground">Daily collection of SharePoint site storage metrics</p>
                  </div>
                </div>
                
                <div className="flex items-start gap-3">
                  <div className="mt-1 bg-primary/10 text-primary rounded-md p-1.5">
                    <Database size={20} weight="duotone" />
                  </div>
                  <div>
                    <h3 className="font-medium text-lg">Secure Credentials</h3>
                    <p className="text-muted-foreground">Azure Key Vault integration for secure credential storage</p>
                  </div>
                </div>
                
                <div className="flex items-start gap-3">
                  <div className="mt-1 bg-primary/10 text-primary rounded-md p-1.5">
                    <LineChart size={20} weight="duotone" />
                  </div>
                  <div>
                    <h3 className="font-medium text-lg">Cross-tenant Visualization</h3>
                    <p className="text-muted-foreground">Unified dashboards across all monitored environments</p>
                  </div>
                </div>
                
                <div className="flex items-start gap-3">
                  <div className="mt-1 bg-primary/10 text-primary rounded-md p-1.5">
                    <Code size={20} weight="duotone" />
                  </div>
                  <div>
                    <h3 className="font-medium text-lg">Module Management</h3>
                    <p className="text-muted-foreground">Automatic PowerShell module installation and management</p>
                  </div>
                </div>
              </CardContent>
            </Card>
            
            <Card>
              <CardHeader>
                <CardTitle>Architecture</CardTitle>
                <CardDescription>How the components fit together</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="bg-muted p-4 rounded-lg">
                  <div className="flex flex-col gap-2 text-center">
                    <div className="grid grid-cols-3 gap-2">
                      <div className="border border-border bg-card rounded-md p-2">
                        Tenant 1
                        <div className="text-xs text-muted-foreground">SharePoint Online</div>
                      </div>
                      <div className="border border-border bg-card rounded-md p-2">
                        Tenant 2
                        <div className="text-xs text-muted-foreground">SharePoint Online</div>
                      </div>
                      <div className="border border-border bg-card rounded-md p-2">
                        Tenant 3
                        <div className="text-xs text-muted-foreground">SharePoint Online</div>
                      </div>
                    </div>
                    <div className="text-muted-foreground text-2xl">↓</div>
                    <div className="border border-border bg-primary text-primary-foreground rounded-md p-2">
                      Azure Function
                      <div className="text-xs opacity-90">Multi-tenant PowerShell script</div>
                    </div>
                    <div className="text-muted-foreground text-2xl">↓</div>
                    <div className="border border-border bg-card rounded-md p-2">
                      Log Analytics
                      <div className="text-xs text-muted-foreground">Unified data storage</div>
                    </div>
                    <div className="text-muted-foreground text-2xl">↓</div>
                    <div className="border border-border bg-accent text-accent-foreground rounded-md p-2">
                      Azure Monitor
                      <div className="text-xs opacity-90">Cross-tenant visualization</div>
                    </div>
                  </div>
                </div>
              </CardContent>
              <CardFooter>
                <Button variant="secondary" className="w-full" onClick={() => setActiveTab("multi-tenant")}>
                  Multi-tenant Features
                </Button>
              </CardFooter>
            </Card>
          </div>

          <div className="mt-6">
            <Card>
              <CardHeader>
                <CardTitle>Prerequisites</CardTitle>
                <CardDescription>What you'll need before getting started</CardDescription>
              </CardHeader>
              <CardContent>
                <ul className="list-disc pl-6 space-y-2">
                  <li>Azure subscription</li>
                  <li>Multiple SharePoint Online tenants with appropriate permissions</li>
                  <li>SharePoint App Registration in each tenant with Sites.Read.All permissions</li>
                  <li>PowerShell 7.2 or higher (for local testing)</li>
                  <li>Azure Key Vault for secure credential storage</li>
                  <li>Azure Log Analytics workspace for centralized monitoring</li>
                </ul>
              </CardContent>
            </Card>
          </div>
        </TabsContent>
        
        <TabsContent value="multi-tenant">
          <div className="grid gap-6">
            <Card>
              <CardHeader>
                <CardTitle>Multi-tenant Capabilities</CardTitle>
                <CardDescription>Monitor multiple SharePoint environments from a single solution</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <p>
                  The SharePoint Storage Monitor now supports collecting and analyzing storage data across multiple SharePoint tenants 
                  in a single unified solution. This enables organizations to:
                </p>
                
                <div className="grid md:grid-cols-2 gap-6 mt-4">
                  <div className="space-y-2">
                    <h3 className="font-medium text-lg flex items-center gap-2">
                      <div className="bg-primary/10 text-primary rounded-md p-1.5">
                        <Buildings size={18} weight="duotone" />
                      </div>
                      Centralized Management
                    </h3>
                    <p className="text-muted-foreground">
                      Monitor all SharePoint environments through a single pane of glass, eliminating the need for separate monitoring solutions.
                    </p>
                  </div>
                  
                  <div className="space-y-2">
                    <h3 className="font-medium text-lg flex items-center gap-2">
                      <div className="bg-primary/10 text-primary rounded-md p-1.5">
                        <Database size={18} weight="duotone" />
                      </div>
                      Tenant Isolation
                    </h3>
                    <p className="text-muted-foreground">
                      Maintain security boundaries between tenants while still enabling unified reporting and analysis.
                    </p>
                  </div>
                  
                  <div className="space-y-2">
                    <h3 className="font-medium text-lg flex items-center gap-2">
                      <div className="bg-primary/10 text-primary rounded-md p-1.5">
                        <LineChart size={18} weight="duotone" />
                      </div>
                      Comparative Analytics
                    </h3>
                    <p className="text-muted-foreground">
                      Compare storage utilization across tenants to identify trends and optimize resource allocation.
                    </p>
                  </div>
                  
                  <div className="space-y-2">
                    <h3 className="font-medium text-lg flex items-center gap-2">
                      <div className="bg-primary/10 text-primary rounded-md p-1.5">
                        <Code size={18} weight="duotone" />
                      </div>
                      Resilient Operation
                    </h3>
                    <p className="text-muted-foreground">
                      Gracefully handle failures in one tenant without affecting data collection from others.
                    </p>
                  </div>
                </div>
              </CardContent>
            </Card>
            
            <Card>
              <CardHeader>
                <CardTitle>Configuration</CardTitle>
                <CardDescription>How to set up multiple tenants in your config file</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="space-y-4">
                  <p>
                    Configure multiple tenants using the JSON configuration format below:
                  </p>
                  <CodeBlock 
                    code={`{
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
}`}
                  />
                </div>
              </CardContent>
              <CardFooter>
                <p className="text-sm text-muted-foreground">
                  Each tenant requires its own app registration in the respective Azure AD tenant with Sites.Read.All permissions.
                </p>
              </CardFooter>
            </Card>
            
            <Card>
              <CardHeader>
                <CardTitle>Cross-Tenant Analytics</CardTitle>
                <CardDescription>Example queries for multi-tenant reporting</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-2">
                  <h3 className="font-medium">Storage By Tenant</h3>
                  <CodeBlock
                    code={`SharePointStorageStats_CL
| summarize TotalStorageGB = sum(StorageUsed_d) / 1024 by TenantName_s
| sort by TotalStorageGB desc
| render piechart`}
                  />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Growth Rate By Tenant</h3>
                  <CodeBlock
                    code={`let startDate = ago(30d);
let endDate = now();
SharePointStorageStats_CL
| where TimeGenerated >= startDate and TimeGenerated <= endDate
| summarize 
    StartStorage = min(StorageUsed_d),
    EndStorage = max(StorageUsed_d)
    by TenantName_s
| extend GrowthMB = EndStorage - StartStorage
| extend GrowthPercent = iff(StartStorage > 0, (GrowthMB / StartStorage) * 100, 0)
| project 
    TenantName = TenantName_s,
    StartStorageGB = StartStorage / 1024,
    EndStorageGB = EndStorage / 1024,
    GrowthGB = GrowthMB / 1024,
    GrowthPercent
| sort by GrowthPercent desc`}
                  />
                </div>
                
                <div className="grid md:grid-cols-2 gap-4 mt-4">
                  <div className="border rounded-md p-4">
                    <h4 className="font-medium mb-2">Site Count by Tenant</h4>
                    <div className="bg-card rounded-md p-3 h-32 flex flex-col">
                      <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                        [Chart visualization]
                      </div>
                    </div>
                  </div>
                  
                  <div className="border rounded-md p-4">
                    <h4 className="font-medium mb-2">Storage Usage Over Time</h4>
                    <div className="bg-card rounded-md p-3 h-32 flex flex-col">
                      <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                        [Chart visualization]
                      </div>
                    </div>
                  </div>
                </div>
              </CardContent>
            </Card>
          </div>
        </TabsContent>
        
        <TabsContent value="code">
          <div className="grid gap-6">
            <Card>
              <CardHeader>
                <div className="flex items-center justify-between">
                  <div>
                    <CardTitle className="flex items-center gap-2">
                      <FileIcon weight="duotone" className="text-primary" /> sharepoint-storage-monitor.ps1
                    </CardTitle>
                    <CardDescription>Multi-tenant PowerShell script for SharePoint data collection</CardDescription>
                  </div>
                  <Button 
                    variant="ghost" 
                    size="sm"
                    onClick={() => handleCopyFile(`# View the complete code in the src/sharepoint-storage-monitor/sharepoint-storage-monitor.ps1 file`)}
                  >
                    <Copy className="mr-2 h-4 w-4" /> Copy
                  </Button>
                </div>
              </CardHeader>
              <CardContent>
                <CodeBlock 
                  code={`# SharePoint Storage Monitor - Multi-tenant Edition
# This script collects SharePoint storage utilization from multiple tenants and sends data to Azure Log Analytics

# Get the current script path
$scriptPath = $PSScriptRoot

# Import configuration if exists
$configPath = Join-Path -Path $scriptPath -ChildPath "config.json"
if (Test-Path $configPath) {
    $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
} else {
    # Default configuration (environment variables are used as fallback)
    $config = @{
        WorkspaceId = $env:WorkspaceId
        WorkspaceKey = $env:WorkspaceKey
        LogName = $env:LogName ?? "SharePointStorageStats"
        KeyVaultName = $env:KeyVaultName
        Tenants = @()
    }

    # If tenants are provided via environment variable
    if ($env:TenantsList) {
        try {
            $config.Tenants = $env:TenantsList | ConvertFrom-Json
        } catch {
            Write-Error "Failed to parse TenantsList environment variable: $_"
        }
    }
}

# Main execution
try {
    Write-Output "Starting SharePoint Storage Monitor for multiple tenants"
    $totalResults = @()
    
    # Process each tenant
    foreach ($tenant in $config.Tenants) {
        try {
            Write-Output "Processing tenant: $($tenant.TenantName)"
            
            # Get credentials from Key Vault
            $credential = Get-SecureCredentials -TenantId $tenant.TenantId -ClientId $tenant.ClientId -KeyVaultName $config.KeyVaultName -SecretName $tenant.SecretName
            
            # Get SharePoint storage statistics
            $storageData = Get-SharePointStorageStats -TenantId $tenant.TenantId -TenantName $tenant.TenantName -Credential $credential -SiteUrls $tenant.SharePointSites
            
            if ($storageData.Count -gt 0) {
                $totalResults += $storageData
                Write-Output "Collected data from $($storageData.Count) sites in tenant $($tenant.TenantName)"
            }
        }
        catch {
            Write-Error "Error processing tenant $($tenant.TenantName): $_"
            # Continue with next tenant
        }
    }
    
    # Send all data to Log Analytics
    if ($totalResults.Count -gt 0) {
        Send-LogAnalyticsData -WorkspaceId $config.WorkspaceId -WorkspaceKey $config.WorkspaceKey -LogName $config.LogName -Data $totalResults
        Write-Output "Successfully sent $($totalResults.Count) total records to Log Analytics"
    }
}`}
                />
                <div className="mt-4 text-sm text-muted-foreground">
                  This is an abbreviated version. The complete script includes comprehensive error handling, secure credential management, and detailed SharePoint data collection across multiple tenants.
                </div>
              </CardContent>
              <CardFooter>
                <Button variant="outline" className="w-full" onClick={() => setActiveTab("deployment")}>
                  See Deployment Instructions
                </Button>
              </CardFooter>
            </Card>
            
            <div className="grid gap-6 md:grid-cols-2">
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center gap-2">
                    <FileIcon weight="duotone" className="text-primary" /> config-sample.json
                  </CardTitle>
                  <CardDescription>Multi-tenant configuration file</CardDescription>
                </CardHeader>
                <CardContent>
                  <CodeBlock 
                    code={`{
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
}`}
                  />
                </CardContent>
              </Card>
              
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center gap-2">
                    <FileIcon weight="duotone" className="text-primary" /> function.json
                  </CardTitle>
                  <CardDescription>Azure Function timer trigger configuration</CardDescription>
                </CardHeader>
                <CardContent>
                  <CodeBlock 
                    code={`{
  "bindings": [
    {
      "name": "Timer",
      "type": "timerTrigger",
      "direction": "in",
      "schedule": "0 0 0 * * *"
    }
  ],
  "disabled": false,
  "scriptFile": "sharepoint-storage-monitor.ps1"
}`}
                  />
                </CardContent>
              </Card>
            </div>
          </div>
        </TabsContent>
        
        <TabsContent value="deployment">
          <div className="grid gap-6">
            <Card>
              <CardHeader>
                <CardTitle>Multi-tenant Deployment Steps</CardTitle>
                <CardDescription>Follow these steps to deploy the solution</CardDescription>
              </CardHeader>
              <CardContent className="space-y-6">
                <div className="space-y-2">
                  <h3 className="font-semibold text-lg flex items-center gap-2">
                    <span className="bg-primary text-primary-foreground rounded-full w-6 h-6 inline-flex items-center justify-center text-sm">1</span> 
                    Register App in Each Azure AD Tenant
                  </h3>
                  <div className="ml-8 space-y-2">
                    <p>Create an app registration in each Azure Active Directory tenant:</p>
                    <ul className="list-disc pl-6 space-y-1">
                      <li>Go to Azure Active Directory {">"}  App registrations</li>
                      <li>Create a new registration in each tenant</li>
                      <li>Grant API permissions: SharePoint {">"}  Sites.Read.All</li>
                      <li>Create a client secret for each app registration and save the values</li>
                      <li>Note the client IDs and tenant IDs for the configuration file</li>
                    </ul>
                  </div>
                </div>
                
                <Separator />
                
                <div className="space-y-2">
                  <h3 className="font-semibold text-lg flex items-center gap-2">
                    <span className="bg-primary text-primary-foreground rounded-full w-6 h-6 inline-flex items-center justify-center text-sm">2</span> 
                    Prepare Multi-tenant Configuration
                  </h3>
                  <div className="ml-8 space-y-2">
                    <p>Create a config.json file with tenant details:</p>
                    <CodeBlock 
                      code={`{
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
}`}
                    />
                  </div>
                </div>
                
                <Separator />
                
                <div className="space-y-2">
                  <h3 className="font-semibold text-lg flex items-center gap-2">
                    <span className="bg-primary text-primary-foreground rounded-full w-6 h-6 inline-flex items-center justify-center text-sm">3</span> 
                    Deploy Azure Resources
                  </h3>
                  <div className="ml-8 space-y-2">
                    <p>Run the enhanced deployment script with your parameters:</p>
                    <CodeBlock 
                      code={`./deploy.ps1 \
  -ResourceGroupName "SharePointMonitor-RG" \
  -Location "eastus" \
  -FunctionAppName "sharepoint-storage-monitor" \
  -KeyVaultName "sp-monitor-kv" \
  -WorkspaceId "your-log-analytics-workspace-id" \
  -WorkspaceKey "your-log-analytics-workspace-key" \
  -TenantsJsonPath "./config.json"`}
                    />
                  </div>
                </div>
                
                <Separator />
                
                <div className="space-y-2">
                  <h3 className="font-semibold text-lg flex items-center gap-2">
                    <span className="bg-primary text-primary-foreground rounded-full w-6 h-6 inline-flex items-center justify-center text-sm">4</span> 
                    Deploy the Function App Code
                  </h3>
                  <div className="ml-8 space-y-2">
                    <p>Package and deploy the code to Azure:</p>
                    <CodeBlock 
                      code={`# Compress the files
Compress-Archive -Path ./sharepoint-storage-monitor/* -DestinationPath function.zip

# Deploy to Azure Function App
Publish-AzWebapp -ResourceGroupName "SharePointMonitor-RG" \
  -Name "sharepoint-storage-monitor" \
  -ArchivePath function.zip`}
                    />
                  </div>
                </div>
                
                <Separator />
                
                <div className="space-y-2">
                  <h3 className="font-semibold text-lg flex items-center gap-2">
                    <span className="bg-primary text-primary-foreground rounded-full w-6 h-6 inline-flex items-center justify-center text-sm">5</span> 
                    Verify Multi-tenant Operation
                  </h3>
                  <div className="ml-8 space-y-2">
                    <p>Check the function logs to verify successful data collection from all tenants:</p>
                    <CodeBlock 
                      code={`# Sample log output
Starting SharePoint Storage Monitor for multiple tenants
Processing tenant: Contoso
Collected data from 2 sites in tenant Contoso
Processing tenant: Fabrikam
Collected data from 2 sites in tenant Fabrikam
Successfully sent 4 total records to Log Analytics`}
                    />
                  </div>
                </div>
              </CardContent>
            </Card>
            
            <Card>
              <CardHeader>
                <CardTitle>Local Testing</CardTitle>
                <CardDescription>Test the multi-tenant function locally before deployment</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="space-y-4">
                  <p>To test the multi-tenant functionality locally:</p>
                  <ol className="list-decimal pl-6 space-y-2">
                    <li>Install Azure Functions Core Tools</li>
                    <li>Update <code>local.settings.json</code> with your multi-tenant configuration:</li>
                  </ol>
                  <CodeBlock code={`{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "powershell",
    "WorkspaceId": "your-workspace-id",
    "WorkspaceKey": "your-workspace-key",
    "LogName": "SharePointStorageStats",
    "KeyVaultName": "your-keyvault-name",
    "TenantsList": "[{\"TenantId\":\"tenant1-id\",\"TenantName\":\"Contoso\",\"ClientId\":\"client-id-1\",\"SecretName\":\"secret-name-1\",\"SharePointSites\":[\"contoso.sharepoint.com/sites/site1\"]},{\"TenantId\":\"tenant2-id\",\"TenantName\":\"Fabrikam\",\"ClientId\":\"client-id-2\",\"SecretName\":\"secret-name-2\",\"SharePointSites\":[\"fabrikam.sharepoint.com/sites/marketing\"]}]"
  }
}`} />
                  <p>Run the function locally with the Azure Functions Core Tools:</p>
                  <CodeBlock code={`func start`} />
                </div>
              </CardContent>
            </Card>
          </div>
        </TabsContent>
        
        <TabsContent value="cicd">
          <div className="grid gap-6">
            <Card>
              <CardHeader>
                <CardTitle>CI/CD Pipeline Options</CardTitle>
                <CardDescription>Automated deployment options for SharePoint Storage Monitor</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <p>
                  SharePoint Storage Monitor supports automated deployment through industry-standard CI/CD platforms.
                  Choose the option that best fits your development workflow:
                </p>
                
                <div className="grid md:grid-cols-2 gap-6 mt-4">
                  <div className="border rounded-md p-4 flex flex-col">
                    <div className="flex items-center gap-3 mb-4">
                      <div className="bg-accent/10 text-accent rounded-md p-2">
                        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M9 19c-4.3 1.4-4.3-2.5-6-3m12 5v-3.5c0-1 .1-1.4-.5-2 2.8-.3 5.5-1.4 5.5-6a4.6 4.6 0 0 0-1.3-3.2 4.2 4.2 0 0 0-.1-3.2s-1.1-.3-3.5 1.3a12.3 12.3 0 0 0-6.2 0C6.5 2.8 5.4 3.1 5.4 3.1a4.2 4.2 0 0 0-.1 3.2A4.6 4.6 0 0 0 4 9.5c0 4.6 2.7 5.7 5.5 6-.6.6-.6 1.2-.5 2V21"></path></svg>
                      </div>
                      <h3 className="font-medium text-lg">GitHub Actions</h3>
                    </div>
                    <p className="text-muted-foreground mb-4">
                      Seamlessly integrate with GitHub repositories for automated deployment pipelines that run directly from your GitHub repository.
                    </p>
                    <ul className="list-disc pl-5 space-y-1 mb-4 flex-grow">
                      <li>Built-in integration with GitHub repositories</li>
                      <li>Easy setup with predefined workflow templates</li>
                      <li>Environment-specific deployments (dev/test/prod)</li>
                      <li>Secret management through GitHub Secrets</li>
                    </ul>
                    <Button variant="outline" className="mt-auto" onClick={() => document.getElementById('github-actions-setup')?.scrollIntoView({ behavior: 'smooth' })}>
                      GitHub Actions Setup
                    </Button>
                  </div>
                  
                  <div className="border rounded-md p-4 flex flex-col">
                    <div className="flex items-center gap-3 mb-4">
                      <div className="bg-primary/10 text-primary rounded-md p-2">
                        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="m18 16 4-4-4-4"></path><path d="m6 8-4 4 4 4"></path><path d="m14.5 4-5 16"></path></svg>
                      </div>
                      <h3 className="font-medium text-lg">Azure DevOps</h3>
                    </div>
                    <p className="text-muted-foreground mb-4">
                      Enterprise-grade CI/CD pipelines with advanced capabilities for complex deployment scenarios and integration with other Azure DevOps services.
                    </p>
                    <ul className="list-disc pl-5 space-y-1 mb-4 flex-grow">
                      <li>Comprehensive enterprise-grade pipelines</li>
                      <li>Advanced approval workflows and gates</li>
                      <li>Tight integration with Azure services</li>
                      <li>Rich environment configuration options</li>
                    </ul>
                    <Button variant="outline" className="mt-auto" onClick={() => document.getElementById('azure-devops-setup')?.scrollIntoView({ behavior: 'smooth' })}>
                      Azure DevOps Setup
                    </Button>
                  </div>
                </div>
              </CardContent>
            </Card>
            
            <div id="github-actions-setup">
              <Card>
                <CardHeader>
                  <CardTitle>GitHub Actions CI/CD Setup</CardTitle>
                  <CardDescription>Automate deployment with GitHub Actions workflows</CardDescription>
                </CardHeader>
                <CardContent className="space-y-6">
                  <div className="space-y-4">
                    <h3 className="font-medium text-lg">Prerequisites</h3>
                    <ul className="list-disc pl-6 space-y-1">
                      <li>GitHub repository with SharePoint Storage Monitor code</li>
                      <li>Azure subscription with appropriate permissions</li>
                      <li>Service principal with access to target resources</li>
                    </ul>
                    
                    <div className="bg-muted p-4 rounded-md">
                      <p className="text-sm mb-2 font-medium">Create a service principal for GitHub Actions:</p>
                      <CodeBlock 
                        code={`# Run this script to generate credentials for GitHub Actions
./setup-cicd-prerequisites.ps1 \\
  -SubscriptionId "your-subscription-id" \\
  -ResourceGroupName "SharePointMonitor-RG" \\
  -ServicePrincipalName "sp-sharepoint-monitor-cicd" \\
  -OutputJsonPath "./sp-credentials.json"`}
                      />
                    </div>
                  </div>
                  
                  <Separator />
                  
                  <div className="space-y-4">
                    <h3 className="font-medium text-lg">Repository Setup</h3>
                    <p>
                      Add the following GitHub Actions workflow file to your repository at <code>.github/workflows/azure-deploy.yml</code>:
                    </p>
                    <CodeBlock 
                      code={`name: Deploy SharePoint Storage Monitor

on:
  push:
    branches: [ main ]
    paths:
      - 'src/sharepoint-storage-monitor/**'
  workflow_dispatch:
    inputs:
      environment:
        description: 'Environment to deploy to'
        required: true
        default: 'dev'
        type: choice
        options:
          - dev
          - test
          - prod

env:
  AZURE_FUNCTIONAPP_NAME: "sharepoint-storage-monitor-$\\{{ github.event.inputs.environment || 'dev' }}"
  AZURE_FUNCTIONAPP_PACKAGE_PATH: 'src/sharepoint-storage-monitor'

jobs:
  build-and-deploy:
    runs-on: windows-latest
    environment: $\\{{ github.event.inputs.environment || 'dev' }}
    
    steps:
    - name: Checkout repository
      uses: actions/checkout@v3

    - name: Setup PowerShell Core
      uses: actions/setup-dotnet@v3
      with:
        dotnet-version: '7.0.x'

    - name: Azure Login
      uses: azure/login@v1
      with:
        creds: $\\{{ secrets.AZURE_CREDENTIALS }}
    
    # Additional deployment steps...`}
                    />
                  </div>
                  
                  <Separator />
                  
                  <div className="space-y-4">
                    <h3 className="font-medium text-lg">GitHub Secrets Configuration</h3>
                    <p>
                      Add these secrets to your GitHub repository settings:
                    </p>
                    <div className="grid md:grid-cols-2 gap-4">
                      <div className="border rounded-md p-3">
                        <h4 className="font-medium mb-1">AZURE_CREDENTIALS</h4>
                        <p className="text-sm text-muted-foreground">
                          JSON content from the service principal credentials file (from setup script)
                        </p>
                      </div>
                      <div className="border rounded-md p-3">
                        <h4 className="font-medium mb-1">RESOURCE_GROUP</h4>
                        <p className="text-sm text-muted-foreground">
                          Name of the target Azure resource group
                        </p>
                      </div>
                      <div className="border rounded-md p-3">
                        <h4 className="font-medium mb-1">LOCATION</h4>
                        <p className="text-sm text-muted-foreground">
                          Azure region (e.g., eastus)
                        </p>
                      </div>
                      <div className="border rounded-md p-3">
                        <h4 className="font-medium mb-1">KEY_VAULT_NAME</h4>
                        <p className="text-sm text-muted-foreground">
                          Name of your Azure Key Vault
                        </p>
                      </div>
                      <div className="border rounded-md p-3">
                        <h4 className="font-medium mb-1">WORKSPACE_ID</h4>
                        <p className="text-sm text-muted-foreground">
                          Log Analytics workspace ID
                        </p>
                      </div>
                      <div className="border rounded-md p-3">
                        <h4 className="font-medium mb-1">WORKSPACE_KEY</h4>
                        <p className="text-sm text-muted-foreground">
                          Log Analytics workspace key
                        </p>
                      </div>
                      <div className="border rounded-md p-3 md:col-span-2">
                        <h4 className="font-medium mb-1">TENANT_CONFIG</h4>
                        <p className="text-sm text-muted-foreground">
                          JSON array containing tenant configurations (without secrets)
                        </p>
                      </div>
                    </div>
                  </div>
                </CardContent>
                <CardFooter>
                  <Button variant="outline" className="w-full" onClick={() => handleCopyFile("View the complete GitHub Actions workflow in src/sharepoint-storage-monitor/.github/workflows/azure-deploy.yml")}>
                    <Copy className="mr-2 h-4 w-4" /> Copy Full GitHub Actions Workflow
                  </Button>
                </CardFooter>
              </Card>
            </div>
            
            <div id="azure-devops-setup">
              <Card>
                <CardHeader>
                  <CardTitle>Azure DevOps CI/CD Setup</CardTitle>
                  <CardDescription>Enterprise pipeline for automated deployment</CardDescription>
                </CardHeader>
                <CardContent className="space-y-6">
                  <div className="space-y-4">
                    <h3 className="font-medium text-lg">Create Azure DevOps Pipeline</h3>
                    <p>
                      Add an <code>azure-pipelines.yml</code> file to your repository:
                    </p>
                    <CodeBlock 
                      code={`trigger:
  branches:
    include:
    - main
  paths:
    include:
    - src/sharepoint-storage-monitor/**

parameters:
- name: environment
  displayName: Environment
  type: string
  default: dev
  values:
  - dev
  - test
  - prod

variables:
  functionAppName: 'sharepoint-storage-monitor-$\\{{ parameters.environment }}'
  functionAppPath: '$(System.DefaultWorkingDirectory)/src/sharepoint-storage-monitor'
  vmImageName: 'windows-latest'
  resourceGroupName: 'SharePointMonitor-$\\{{ parameters.environment }}-RG'

stages:
- stage: Build
  displayName: Build and Package
  jobs:
  - job: BuildPackage
    displayName: Build and Package Function App
    pool:
      vmImage: $(vmImageName)
    steps:
    - task: PowerShell@2
      displayName: Install Required Modules
      # Additional steps...

- stage: Deploy
  displayName: Deploy to $\\{{ parameters.environment }}
  # Deployment steps...`}
                    />
                  </div>
                  
                  <Separator />
                  
                  <div className="space-y-4">
                    <h3 className="font-medium text-lg">Service Connections</h3>
                    <p>
                      Configure service connections in your Azure DevOps project:
                    </p>
                    <ol className="list-decimal pl-6 space-y-2">
                      <li>Go to Project Settings > Service Connections</li>
                      <li>Create an Azure Resource Manager connection</li>
                      <li>Name it <code>serviceconnection-dev</code></li>
                      <li>Create additional connections for other environments (<code>serviceconnection-test</code>, <code>serviceconnection-prod</code>)</li>
                    </ol>
                    <div className="bg-muted p-4 rounded-md mt-2">
                      <p className="text-sm font-medium">Service Connection Naming Convention:</p>
                      <p className="text-xs text-muted-foreground mt-1">
                        The pipeline expects service connections named <code>serviceconnection-dev</code>, <code>serviceconnection-test</code>, and <code>serviceconnection-prod</code>. 
                        This naming convention allows the pipeline to dynamically select the appropriate service connection based on the target environment.
                      </p>
                    </div>
                  </div>
                  
                  <Separator />
                  
                  <div className="space-y-4">
                    <h3 className="font-medium text-lg">Pipeline Variables</h3>
                    <p>
                      Set up these variables in your Azure DevOps pipeline:
                    </p>
                    <div className="grid md:grid-cols-2 gap-4">
                      <div className="border rounded-md p-3">
                        <h4 className="font-medium mb-1">keyVaultName</h4>
                        <p className="text-sm text-muted-foreground">
                          Name of your Azure Key Vault
                        </p>
                      </div>
                      <div className="border rounded-md p-3">
                        <h4 className="font-medium mb-1">workspaceId</h4>
                        <p className="text-sm text-muted-foreground">
                          Log Analytics workspace ID
                        </p>
                      </div>
                      <div className="border rounded-md p-3">
                        <h4 className="font-medium mb-1">workspaceKey</h4>
                        <p className="text-sm text-muted-foreground">
                          Log Analytics workspace key (mark as secret)
                        </p>
                      </div>
                      <div className="border rounded-md p-3">
                        <h4 className="font-medium mb-1">tenantConfig</h4>
                        <p className="text-sm text-muted-foreground">
                          JSON string with tenant configurations (mark as secret)
                        </p>
                      </div>
                    </div>
                  </div>
                  
                  <Separator />
                  
                  <div className="space-y-4">
                    <h3 className="font-medium text-lg">Environment Configuration</h3>
                    <p>
                      Create and configure environments in Azure DevOps:
                    </p>
                    <ol className="list-decimal pl-6 space-y-2">
                      <li>Go to Pipelines > Environments</li>
                      <li>Create environments named <code>dev</code>, <code>test</code>, and <code>prod</code></li>
                      <li>For <code>test</code> and <code>prod</code>, add approval requirements:</li>
                    </ol>
                    <div className="bg-muted p-4 rounded-md mt-2">
                      <p className="text-sm font-medium">Environment Approval Configuration:</p>
                      <ul className="list-disc pl-4 text-xs text-muted-foreground mt-1 space-y-1">
                        <li>Select the environment and click "Approvals and checks"</li>
                        <li>Add an "Approvals" check</li>
                        <li>Specify required approvers</li>
                        <li>Set timeout (e.g., 7 days)</li>
                        <li>Optional: Add instructions for approvers</li>
                      </ul>
                    </div>
                  </div>
                </CardContent>
                <CardFooter>
                  <Button variant="outline" className="w-full" onClick={() => handleCopyFile("View the complete Azure DevOps pipeline in src/sharepoint-storage-monitor/azure-pipelines.yml")}>
                    <Copy className="mr-2 h-4 w-4" /> Copy Full Azure DevOps Pipeline
                  </Button>
                </CardFooter>
              </Card>
            </div>
            
            <Card>
              <CardHeader>
                <CardTitle>Security Best Practices</CardTitle>
                <CardDescription>Secure CI/CD implementation guidelines</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="space-y-4">
                  <div className="flex items-start gap-3">
                    <div className="mt-1 bg-primary/10 text-primary rounded-md p-1.5">
                      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect width="18" height="11" x="3" y="11" rx="2" ry="2"></rect><path d="M7 11V7a5 5 0 0 1 10 0v4"></path></svg>
                    </div>
                    <div>
                      <h3 className="font-medium">Secure Secret Management</h3>
                      <p className="text-muted-foreground">
                        Store all sensitive information (client secrets, connection strings) in secure secret stores like GitHub Secrets or Azure Key Vault.
                        Never include secrets directly in code or configuration files.
                      </p>
                    </div>
                  </div>
                  
                  <div className="flex items-start gap-3">
                    <div className="mt-1 bg-primary/10 text-primary rounded-md p-1.5">
                      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"></circle><path d="m9 9 6 6"></path><path d="m9 15 6-6"></path></svg>
                    </div>
                    <div>
                      <h3 className="font-medium">Principle of Least Privilege</h3>
                      <p className="text-muted-foreground">
                        Service principals and identities should have only the minimum permissions needed. 
                        Use scoped access to specific resource groups rather than subscription-level permissions.
                      </p>
                    </div>
                  </div>
                  
                  <div className="flex items-start gap-3">
                    <div className="mt-1 bg-primary/10 text-primary rounded-md p-1.5">
                      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10"></path></svg>
                    </div>
                    <div>
                      <h3 className="font-medium">Environment Isolation</h3>
                      <p className="text-muted-foreground">
                        Maintain strict isolation between environments (dev/test/prod) using separate resource groups,
                        service principals, and approval processes for production deployments.
                      </p>
                    </div>
                  </div>
                  
                  <div className="flex items-start gap-3">
                    <div className="mt-1 bg-primary/10 text-primary rounded-md p-1.5">
                      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect width="20" height="14" x="2" y="3" rx="2"></rect><path d="M12 17v4"></path><path d="M8 21h8"></path><path d="M15 8a3 3 0 1 0-3 3"></path></svg>
                    </div>
                    <div>
                      <h3 className="font-medium">Secret Rotation</h3>
                      <p className="text-muted-foreground">
                        Implement a regular schedule for rotating client secrets, service principal credentials, and other sensitive information.
                        Automate the rotation process where possible.
                      </p>
                    </div>
                  </div>
                  
                  <div className="flex items-start gap-3">
                    <div className="mt-1 bg-primary/10 text-primary rounded-md p-1.5">
                      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M15 3h6v6"></path><path d="M10 14 21 3"></path><path d="M8 14H3v7h7v-5h5V9h-7z"></path></svg>
                    </div>
                    <div>
                      <h3 className="font-medium">Pipeline Security Scanning</h3>
                      <p className="text-muted-foreground">
                        Integrate security scanning tools into your pipelines to detect vulnerabilities, secrets in code,
                        and insecure configurations before deployment.
                      </p>
                    </div>
                  </div>
                </div>
              </CardContent>
            </Card>
          </div>
        </TabsContent>

        <TabsContent value="dashboard">
          <div className="grid gap-6">
            <Card>
              <CardHeader>
                <CardTitle>Multi-tenant Azure Monitor Dashboard</CardTitle>
                <CardDescription>Visualize SharePoint storage data across all tenants</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <p>
                  After collecting data from multiple tenants, you can create a unified dashboard in Azure Monitor to visualize storage trends across your entire SharePoint estate.
                </p>
                
                <div className="rounded-lg border overflow-hidden">
                  <div className="bg-muted px-4 py-3 border-b">
                    <h3 className="font-semibold">Multi-tenant Dashboard Layout</h3>
                  </div>
                  <div className="p-4 grid gap-3">
                    <div className="grid md:grid-cols-2 gap-3">
                      <div className="bg-card border rounded-md p-3 h-32 flex flex-col">
                        <div className="text-sm font-medium mb-1">Storage by Tenant</div>
                        <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                          [Pie chart visualization]
                        </div>
                      </div>
                      <div className="bg-card border rounded-md p-3 h-32 flex flex-col">
                        <div className="text-sm font-medium mb-1">Total Storage Usage Over Time</div>
                        <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                          [Line chart visualization with tenant breakdown]
                        </div>
                      </div>
                    </div>
                    <div className="grid md:grid-cols-2 gap-3">
                      <div className="bg-card border rounded-md p-3 h-32 flex flex-col">
                        <div className="text-sm font-medium mb-1">Tenant Connection Status</div>
                        <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                          [Status visualization with success/failure metrics]
                        </div>
                      </div>
                      <div className="bg-card border rounded-md p-3 h-32 flex flex-col">
                        <div className="text-sm font-medium mb-1">Error Trends by Category</div>
                        <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                          [Error category breakdown chart]
                        </div>
                      </div>
                    </div>
                    <div className="bg-card border rounded-md p-3 h-32 flex flex-col">
                      <div className="text-sm font-medium mb-1">Top Sites Across All Tenants</div>
                      <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                        [Table visualization]
                      </div>
                    </div>
                    <div className="grid md:grid-cols-2 gap-3">
                      <div className="bg-card border rounded-md p-3 h-32 flex flex-col">
                        <div className="text-sm font-medium mb-1">Growth Rate by Tenant (30 Days)</div>
                        <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                          [Bar chart visualization]
                        </div>
                      </div>
                      <div className="bg-card border rounded-md p-3 h-32 flex flex-col">
                        <div className="text-sm font-medium mb-1">Site Count by Tenant</div>
                        <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                          [Bar chart visualization]
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </CardContent>
              <CardFooter className="flex-col items-start gap-3">
                <p className="text-sm text-muted-foreground">
                  A full dashboard template JSON file is included in the project files for easy import.
                </p>
                <Button variant="outline" className="w-full">
                  Download Multi-tenant Dashboard Template
                </Button>
              </CardFooter>
            </Card>
            
            <Card>
              <CardHeader>
                <CardTitle>Multi-tenant Sample Queries</CardTitle>
                <CardDescription>Log Analytics queries to visualize data across all tenants</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-2">
                  <h3 className="font-medium">Storage by Tenant</h3>
                  <CodeBlock
                    code={`SharePointStorageStats_CL
| summarize TotalStorageGB = sum(StorageUsed_d) / 1024 by TenantName_s
| sort by TotalStorageGB desc
| render piechart`}
                  />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Storage Usage Over Time by Tenant</h3>
                  <CodeBlock
                    code={`SharePointStorageStats_CL
| summarize AvgStorageUsedGB = sum(StorageUsed_d) / 1024 by TenantName_s, bin(TimeGenerated, 1d)
| render timechart`}
                  />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Top 10 Sites Across All Tenants</h3>
                  <CodeBlock
                    code={`SharePointStorageStats_CL
| summarize arg_max(TimeGenerated, *) by SiteUrl_s
| project TenantName_s, SiteTitle_s, SiteUrl_s, StorageUsedGB = StorageUsed_d / 1024, StorageLimitGB = StorageLimit_d / 1024, PercentageUsed_d
| sort by StorageUsedGB desc
| take 10`}
                  />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Connection Error Monitoring</h3>
                  <CodeBlock
                    code={`SharePointStorageErrors_CL
| where TimeGenerated > ago(7d)
| where Status_s == "Failed" or Status_s == "TenantFailed"
| extend ErrorType = tostring(extract("(AccessDenied|Timeout|NotFound|NetworkIssue|Credentials)", 1, ErrorMessage_s))
| extend ErrorType = iif(isempty(ErrorType), "Other", ErrorType)
| summarize ErrorCount=count() by ErrorType, TenantName_s, bin(TimeGenerated, 1d)
| render timechart`}
                  />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Tenant Health Summary</h3>
                  <CodeBlock
                    code={`// Sites by connection status
let data = SharePointStorageStats_CL
| where TimeGenerated > ago(24h)
| extend Status = "Success";
let errors = SharePointStorageErrors_CL
| where TimeGenerated > ago(24h)
| extend Status = "Failed";
union data, errors
| summarize SuccessCount=countif(Status=="Success"), FailCount=countif(Status=="Failed") by TenantName_s
| extend SuccessRate = 100.0 * SuccessCount / (SuccessCount + FailCount)
| project TenantName_s, SuccessCount, FailCount, SuccessRate
| sort by SuccessRate asc`}
                  />
                </div>

                <div className="space-y-2">
                  <h3 className="font-medium">Growth Rate Analysis by Tenant</h3>
                  <CodeBlock
                    code={`let startDate = ago(30d);
let endDate = now();
SharePointStorageStats_CL
| where TimeGenerated >= startDate and TimeGenerated <= endDate
| summarize 
    StartStorage = min(StorageUsed_d),
    EndStorage = max(StorageUsed_d)
    by TenantName_s
| extend GrowthMB = EndStorage - StartStorage
| extend GrowthPercent = iff(StartStorage > 0, (GrowthMB / StartStorage) * 100, 0)
| project 
    TenantName = TenantName_s,
    StartStorageGB = StartStorage / 1024,
    EndStorageGB = EndStorage / 1024,
    GrowthGB = GrowthMB / 1024,
    GrowthPercent
| sort by GrowthPercent desc`}
                  />
                </div>
              </CardContent>
            </Card>
          </div>
        </TabsContent>
        
        <TabsContent value="troubleshooting">
          <div className="grid gap-6 md:grid-cols-2">
            <Card>
              <CardHeader>
                <CardTitle>Enhanced Error Handling</CardTitle>
                <CardDescription>New resilient multi-tenant features</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-2">
                  <h3 className="font-medium">Tenant Connection Resilience</h3>
                  <p className="text-muted-foreground text-sm">
                    The script now includes improved error handling for tenant connections:
                  </p>
                  <ul className="list-disc pl-5 text-sm space-y-1">
                    <li>Automatic retry attempts for failed tenant connections</li>
                    <li>Independent processing for each tenant to prevent cascading failures</li>
                    <li>Detailed error logging with error categorization</li>
                    <li>Configurable retry parameters in the config file</li>
                  </ul>
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Configuration Options</h3>
                  <p className="text-muted-foreground text-sm">
                    New configuration settings for error handling:
                  </p>
                  <CodeBlock
                    code={`{
  "MaxRetries": 3,         // Number of retry attempts for failed operations
  "RetryDelaySeconds": 5,  // Delay between retry attempts
  "LogErrors": true,       // Enable error logging to Log Analytics
  "ErrorLogName": "SharePointStorageErrors"  // Custom log name for errors
}`}
                  />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Error Categorization</h3>
                  <p className="text-muted-foreground text-sm">
                    Errors are now categorized for better troubleshooting:
                  </p>
                  <ul className="list-disc pl-5 text-sm space-y-1">
                    <li><strong>AccessDenied</strong>: Permission issues</li>
                    <li><strong>Timeout</strong>: Connection or operation timeouts</li>
                    <li><strong>NotFound</strong>: Resources not found</li>
                    <li><strong>NetworkIssue</strong>: Connectivity problems</li>
                    <li><strong>Credentials</strong>: Authentication failures</li>
                  </ul>
                </div>
              </CardContent>
            </Card>
            
            <Card>
              <CardHeader>
                <CardTitle>Common Multi-tenant Issues</CardTitle>
                <CardDescription>Problems and their solutions</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-2">
                  <h3 className="font-medium">Authentication Failures</h3>
                  <p className="text-muted-foreground text-sm">
                    If you're seeing authentication errors for specific tenants:
                  </p>
                  <ul className="list-disc pl-5 text-sm space-y-1">
                    <li>Verify client secret hasn't expired for that tenant</li>
                    <li>Check that the app in that tenant has Sites.Read.All permissions</li>
                    <li>Ensure admin consent was granted for the permissions in each tenant</li>
                    <li>Verify tenant ID is correct in the config file</li>
                  </ul>
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Missing Data from Specific Tenant</h3>
                  <p className="text-muted-foreground text-sm">
                    If data appears from some tenants but not others:
                  </p>
                  <ul className="list-disc pl-5 text-sm space-y-1">
                    <li>Check function execution logs for tenant-specific errors</li>
                    <li>Verify the tenant's SharePoint sites are accessible</li>
                    <li>Test connection to that tenant's SharePoint sites manually</li>
                    <li>Check if the secret for that tenant exists in Key Vault</li>
                  </ul>
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Tenant List Not Processing</h3>
                  <p className="text-muted-foreground text-sm">
                    If the tenant list isn't being processed correctly:
                  </p>
                  <ul className="list-disc pl-5 text-sm space-y-1">
                    <li>Verify TenantsList JSON format is valid</li>
                    <li>Check for malformed JSON in config file or environment variable</li>
                    <li>Ensure all required properties are present for each tenant</li>
                  </ul>
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Function Timeout Issues</h3>
                  <p className="text-muted-foreground text-sm">
                    If the function times out with multiple tenants:
                  </p>
                  <ul className="list-disc pl-5 text-sm space-y-1">
                    <li>Consider increasing the function timeout setting</li>
                    <li>Split tenants into separate function instances if needed</li>
                    <li>Optimize site collection list per tenant</li>
                  </ul>
                </div>
              </CardContent>
            </Card>
            
            <Card>
              <CardHeader>
                <CardTitle>Multi-tenant Diagnostic Steps</CardTitle>
                <CardDescription>How to diagnose issues across tenants</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-2">
                  <h3 className="font-medium">Check Tenant Processing</h3>
                  <p className="text-muted-foreground text-sm">
                    Verify which tenants are being processed:
                  </p>
                  <CodeBlock
                    code={`# Check function logs for tenant processing
Starting SharePoint Storage Monitor for multiple tenants
Processing tenant: Contoso
...
Processing tenant: Fabrikam
...`}
                  />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Test Individual Tenant Access</h3>
                  <p className="text-muted-foreground text-sm">
                    Isolate and test a single tenant:
                  </p>
                  <CodeBlock
                    code={`# Create a test script for a single tenant
$tenantId = "tenant-id"
$clientId = "client-id"
$clientSecret = "client-secret"
$siteUrl = "https://tenant.sharepoint.com/sites/site1"

# Create credential
$secPassword = ConvertTo-SecureString $clientSecret -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($clientId, $secPassword)

# Test connection
Connect-PnPOnline -Url $siteUrl -Credentials $cred
Get-PnPWeb`}
                  />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Verify Data by Tenant in Log Analytics</h3>
                  <p className="text-muted-foreground text-sm">
                    Check which tenants are reporting data:
                  </p>
                  <CodeBlock
                    code={`SharePointStorageStats_CL
| where TimeGenerated > ago(24h)
| summarize count() by TenantName_s`}
                  />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Key Vault Secret Access</h3>
                  <p className="text-muted-foreground text-sm">
                    Verify Function App can access secrets for all tenants:
                  </p>
                  <CodeBlock
                    code={`# Test Key Vault access in Azure Cloud Shell
$keyVaultName = "your-keyvault-name"
$functionAppIdentity = (Get-AzFunctionApp -ResourceGroupName "YourRG" -Name "YourFunctionApp").IdentityPrincipalId
Get-AzKeyVaultAccessPolicy -VaultName $keyVaultName | Where-Object {$_.ObjectId -eq $functionAppIdentity}`}
                  />
                </div>
              </CardContent>
              <CardFooter>
                <p className="text-sm text-muted-foreground">
                  If issues persist with a specific tenant, try removing it from the configuration temporarily to ensure other tenants continue working while you troubleshoot.
                </p>
              </CardFooter>
            </Card>
          </div>
        </TabsContent>
      </Tabs>
    </div>
  );
}

export default App;