import { useState } from "react";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Separator } from "@/components/ui/separator";
import { CodeBlock } from "./components/code-block";
import { FileIcon, Database, LineChart, Terminal, Code, FileCode, Copy, Info } from "@phosphor-icons/react";

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
          PowerShell application that runs as an Azure Function to track and visualize SharePoint storage utilization
        </p>
      </header>

      <Tabs defaultValue="overview" value={activeTab} onValueChange={setActiveTab} className="space-y-6">
        <TabsList className="grid w-full grid-cols-3 md:grid-cols-5">
          <TabsTrigger value="overview" className="flex items-center gap-2">
            <Info weight="duotone" /> Overview
          </TabsTrigger>
          <TabsTrigger value="code" className="flex items-center gap-2">
            <FileCode weight="duotone" /> Code
          </TabsTrigger>
          <TabsTrigger value="deployment" className="flex items-center gap-2">
            <Terminal weight="duotone" /> Deployment
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
                    <h3 className="font-medium text-lg">Visualization</h3>
                    <p className="text-muted-foreground">Data visualization through Azure Monitor dashboards</p>
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
                    <div className="border border-border bg-card rounded-md p-2">
                      PowerShell Script
                      <div className="text-xs text-muted-foreground">SharePoint data collection</div>
                    </div>
                    <div className="text-muted-foreground text-2xl">↓</div>
                    <div className="border border-border bg-primary text-primary-foreground rounded-md p-2">
                      Azure Function
                      <div className="text-xs opacity-90">Daily scheduled execution</div>
                    </div>
                    <div className="text-muted-foreground text-2xl">↓</div>
                    <div className="border border-border bg-card rounded-md p-2">
                      Log Analytics
                      <div className="text-xs text-muted-foreground">Data storage</div>
                    </div>
                    <div className="text-muted-foreground text-2xl">↓</div>
                    <div className="border border-border bg-accent text-accent-foreground rounded-md p-2">
                      Azure Monitor
                      <div className="text-xs opacity-90">Data visualization</div>
                    </div>
                  </div>
                </div>
              </CardContent>
              <CardFooter>
                <Button variant="secondary" className="w-full" onClick={() => setActiveTab("deployment")}>
                  Deployment Guide
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
                  <li>SharePoint Online tenant with appropriate permissions</li>
                  <li>SharePoint App Registration with Sites.Read.All permissions</li>
                  <li>PowerShell 7.2 or higher (for local testing)</li>
                </ul>
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
                    <CardDescription>Main PowerShell script for SharePoint data collection</CardDescription>
                  </div>
                  <Button 
                    variant="ghost" 
                    size="sm"
                    onClick={() => handleCopyFile(`# View the complete code in the src/sharepoint-storage-monitor.ps1 file`)}
                  >
                    <Copy className="mr-2 h-4 w-4" /> Copy
                  </Button>
                </div>
              </CardHeader>
              <CardContent>
                <CodeBlock 
                  code={`# SharePoint Storage Monitor
# This script collects SharePoint storage utilization and sends data to Azure Log Analytics

# Get the current script path
$scriptPath = $PSScriptRoot

# Import configuration if exists
$configPath = Join-Path -Path $scriptPath -ChildPath "config.json"
if (Test-Path $configPath) {
    $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
} else {
    # Default configuration
    $config = @{
        TenantId = $env:TenantId
        ClientId = $env:ClientId
        KeyVaultName = $env:KeyVaultName
        SecretName = $env:SecretName
        # Additional configuration...
    }
}

# Function to check and install modules
function Ensure-Module {
    param (
        [string]$ModuleName,
        [string]$MinimumVersion = ""
    )
    
    # Function implementation...
}

# Main execution
try {
    # Get credentials from Key Vault
    $credential = Get-SecureCredentials -TenantId $config.TenantId -ClientId $config.ClientId -KeyVaultName $config.KeyVaultName -SecretName $config.SecretName
    
    # Get SharePoint storage statistics
    $storageData = Get-SharePointStorageStats -TenantId $config.TenantId -Credential $credential -SiteUrls $config.SharePointSites
    
    # Send data to Log Analytics
    if ($storageData.Count -gt 0) {
        Send-LogAnalyticsData -WorkspaceId $config.WorkspaceId -WorkspaceKey $config.WorkspaceKey -LogName $config.LogName -Data $storageData
    }
}
catch {
    Write-Error "Error in main execution: $_"
    throw
}`}
                />
                <div className="mt-4 text-sm text-muted-foreground">
                  This is an abbreviated version. The complete script includes comprehensive error handling, secure credential management, and detailed SharePoint data collection.
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
  "disabled": false
}`}
                  />
                </CardContent>
              </Card>
              
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center gap-2">
                    <FileIcon weight="duotone" className="text-primary" /> config-sample.json
                  </CardTitle>
                  <CardDescription>Sample configuration file</CardDescription>
                </CardHeader>
                <CardContent>
                  <CodeBlock 
                    code={`{
  "TenantId": "your-tenant-id",
  "ClientId": "your-client-id",
  "KeyVaultName": "your-keyvault-name",
  "SecretName": "your-secret-name",
  "WorkspaceId": "your-log-analytics-workspace-id",
  "WorkspaceKey": "your-log-analytics-workspace-key",
  "SharePointSites": [
    "tenant.sharepoint.com/sites/site1",
    "tenant.sharepoint.com/sites/site2"
  ],
  "LogName": "SharePointStorageStats"
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
                <CardTitle>Deployment Steps</CardTitle>
                <CardDescription>Follow these steps to deploy the solution</CardDescription>
              </CardHeader>
              <CardContent className="space-y-6">
                <div className="space-y-2">
                  <h3 className="font-semibold text-lg flex items-center gap-2">
                    <span className="bg-primary text-primary-foreground rounded-full w-6 h-6 inline-flex items-center justify-center text-sm">1</span> 
                    Register an Azure AD Application
                  </h3>
                  <div className="ml-8 space-y-2">
                    <p>Create an app registration in Azure Active Directory:</p>
                    <ul className="list-disc pl-6 space-y-1">
                      <li>Go to Azure Active Directory > App registrations</li>
                      <li>Create a new registration</li>
                      <li>Grant API permissions: SharePoint > Sites.Read.All</li>
                      <li>Create a client secret and save the value</li>
                    </ul>
                  </div>
                </div>
                
                <Separator />
                
                <div className="space-y-2">
                  <h3 className="font-semibold text-lg flex items-center gap-2">
                    <span className="bg-primary text-primary-foreground rounded-full w-6 h-6 inline-flex items-center justify-center text-sm">2</span> 
                    Deploy Azure Resources
                  </h3>
                  <div className="ml-8 space-y-2">
                    <p>Run the deployment script with your parameters:</p>
                    <CodeBlock 
                      code={`./deploy.ps1 \\
  -ResourceGroupName "SharePointMonitor-RG" \\
  -Location "eastus" \\
  -FunctionAppName "sharepoint-storage-monitor" \\
  -KeyVaultName "sp-monitor-kv" \\
  -SharePointClientSecret "your-client-secret" \\
  -TenantId "your-tenant-id" \\
  -ClientId "your-app-registration-client-id" \\
  -SharePointSites "tenant.sharepoint.com/sites/site1,tenant.sharepoint.com/sites/site2"`}
                    />
                  </div>
                </div>
                
                <Separator />
                
                <div className="space-y-2">
                  <h3 className="font-semibold text-lg flex items-center gap-2">
                    <span className="bg-primary text-primary-foreground rounded-full w-6 h-6 inline-flex items-center justify-center text-sm">3</span> 
                    Deploy the Function App Code
                  </h3>
                  <div className="ml-8 space-y-2">
                    <p>Package and deploy the code to Azure:</p>
                    <CodeBlock 
                      code={`# Compress the files
Compress-Archive -Path *.ps1,*.json -DestinationPath function.zip

# Deploy to Azure Function App
Publish-AzWebapp -ResourceGroupName "SharePointMonitor-RG" \\
  -Name "sharepoint-storage-monitor" \\
  -ArchivePath function.zip`}
                    />
                  </div>
                </div>
              </CardContent>
            </Card>
            
            <Card>
              <CardHeader>
                <CardTitle>Local Testing</CardTitle>
                <CardDescription>Test the function locally before deployment</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="space-y-4">
                  <p>To test locally before deployment:</p>
                  <ol className="list-decimal pl-6 space-y-2">
                    <li>Install Azure Functions Core Tools</li>
                    <li>Update <code>local.settings.json</code> with your values</li>
                    <li>Run the function locally:</li>
                  </ol>
                  <CodeBlock code={`func start`} />
                </div>
              </CardContent>
            </Card>
          </div>
        </TabsContent>

        <TabsContent value="dashboard">
          <div className="grid gap-6">
            <Card>
              <CardHeader>
                <CardTitle>Azure Monitor Dashboard</CardTitle>
                <CardDescription>Visualize SharePoint storage data</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <p>
                  Once data is collected, you can create a custom dashboard in Azure Monitor to visualize storage utilization trends.
                </p>
                
                <div className="rounded-lg border overflow-hidden">
                  <div className="bg-muted px-4 py-3 border-b">
                    <h3 className="font-semibold">Sample Dashboard Layout</h3>
                  </div>
                  <div className="p-4 grid gap-3">
                    <div className="grid md:grid-cols-2 gap-3">
                      <div className="bg-card border rounded-md p-3 h-32 flex flex-col">
                        <div className="text-sm font-medium mb-1">Storage Usage Over Time</div>
                        <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                          [Line chart visualization]
                        </div>
                      </div>
                      <div className="bg-card border rounded-md p-3 h-32 flex flex-col">
                        <div className="text-sm font-medium mb-1">Storage Percentage Used</div>
                        <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                          [Line chart visualization]
                        </div>
                      </div>
                    </div>
                    <div className="bg-card border rounded-md p-3 h-32 flex flex-col">
                      <div className="text-sm font-medium mb-1">Latest Storage Data</div>
                      <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                        [Table visualization]
                      </div>
                    </div>
                    <div className="grid md:grid-cols-2 gap-3">
                      <div className="bg-card border rounded-md p-3 h-32 flex flex-col">
                        <div className="text-sm font-medium mb-1">Storage by Site</div>
                        <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                          [Pie chart visualization]
                        </div>
                      </div>
                      <div className="bg-card border rounded-md p-3 h-32 flex flex-col">
                        <div className="text-sm font-medium mb-1">30-Day Growth Rate</div>
                        <div className="flex-1 flex items-center justify-center text-muted-foreground text-xs">
                          [Table visualization]
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
                  Download Dashboard Template
                </Button>
              </CardFooter>
            </Card>
            
            <Card>
              <CardHeader>
                <CardTitle>Sample Queries</CardTitle>
                <CardDescription>Log Analytics queries to visualize your data</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-2">
                  <h3 className="font-medium">Storage Usage Over Time</h3>
                  <CodeBlock
                    code={`SharePointStorageStats_CL
| project TimeGenerated, SiteUrl_s, StorageUsed_d
| render timechart`}
                  />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Storage Usage by Site</h3>
                  <CodeBlock
                    code={`SharePointStorageStats_CL
| summarize arg_max(TimeGenerated, *) by SiteUrl_s
| project SiteTitle_s, StorageUsed_d
| sort by StorageUsed_d desc
| render piechart`}
                  />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Growth Rate Analysis</h3>
                  <CodeBlock
                    code={`let startDate = ago(30d);
let endDate = now();
SharePointStorageStats_CL
| where TimeGenerated >= startDate and TimeGenerated <= endDate
| summarize StartStorageUsed = min(StorageUsed_d), EndStorageUsed = max(StorageUsed_d) by SiteUrl_s
| extend GrowthMB = EndStorageUsed - StartStorageUsed
| extend GrowthPercent = iff(StartStorageUsed > 0, (GrowthMB / StartStorageUsed) * 100, 0)
| project SiteUrl_s, StartStorageUsed, EndStorageUsed, GrowthMB, GrowthPercent
| order by GrowthMB desc`}
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
                <CardTitle>Common Issues</CardTitle>
                <CardDescription>Problems and their solutions</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-2">
                  <h3 className="font-medium">Authentication Failures</h3>
                  <p className="text-muted-foreground text-sm">
                    If you're seeing authentication errors:
                  </p>
                  <ul className="list-disc pl-5 text-sm space-y-1">
                    <li>Verify client secret hasn't expired</li>
                    <li>Check that app has Sites.Read.All permissions</li>
                    <li>Ensure admin consent was granted for the permissions</li>
                  </ul>
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Missing Data</h3>
                  <p className="text-muted-foreground text-sm">
                    If no data appears in Log Analytics:
                  </p>
                  <ul className="list-disc pl-5 text-sm space-y-1">
                    <li>Check function execution logs for errors</li>
                    <li>Verify workspace ID and key are correct</li>
                    <li>Ensure sites are accessible to the app identity</li>
                  </ul>
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Module Installation Failures</h3>
                  <p className="text-muted-foreground text-sm">
                    If modules fail to install:
                  </p>
                  <ul className="list-disc pl-5 text-sm space-y-1">
                    <li>Verify PowerShell execution policy</li>
                    <li>Check for network connectivity</li>
                    <li>Ensure function app has internet access</li>
                  </ul>
                </div>
              </CardContent>
            </Card>
            
            <Card>
              <CardHeader>
                <CardTitle>Diagnostic Steps</CardTitle>
                <CardDescription>How to diagnose issues</CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-2">
                  <h3 className="font-medium">Check Function Logs</h3>
                  <p className="text-muted-foreground text-sm">
                    Access logs via the Azure Portal:
                  </p>
                  <ol className="list-decimal pl-5 text-sm space-y-1">
                    <li>Go to your Function App</li>
                    <li>Navigate to Functions > SharePointStorageMonitor</li>
                    <li>Select Monitor tab</li>
                    <li>Review invocation logs</li>
                  </ol>
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Test Script Locally</h3>
                  <p className="text-muted-foreground text-sm">
                    Run the script in your local environment:
                  </p>
                  <CodeBlock
                    code={`# Set required environment variables
$env:TenantId = "your-tenant-id"
$env:ClientId = "your-client-id"
# ... other variables

# Run script with detailed logging
./sharepoint-storage-monitor.ps1 -Verbose`}
                  />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-medium">Verify Log Analytics Data</h3>
                  <p className="text-muted-foreground text-sm">
                    Check if data is reaching Log Analytics:
                  </p>
                  <CodeBlock
                    code={`SharePointStorageStats_CL
| where TimeGenerated > ago(24h)
| summarize count()`}
                  />
                </div>
              </CardContent>
              <CardFooter>
                <p className="text-sm text-muted-foreground">
                  If issues persist, check the function's Application Insights telemetry for detailed error information.
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