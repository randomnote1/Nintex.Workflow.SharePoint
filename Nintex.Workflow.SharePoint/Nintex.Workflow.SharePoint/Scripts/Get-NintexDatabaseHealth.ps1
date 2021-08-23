
try
{
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction Stop
}
catch
{
    $message = 'Unable to load the SharePoint PowerShell snapin.'
    Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 1 -Description $message -DebugLogging
    exit
}

$farm = Get-SPFarm

try
{
    $nwAdminExe = Get-Command -Name NWAdmin.exe -ErrorAction Stop
}
catch
{
    $message = 'Could not find NWAdmin.exe.'
    Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 1 -Description $message -DebugLogging
    exit
}

$message = "`nNWAdmin.exe Path: $($nwAdminExe.Source)"
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug

# Get the Nintex Workflow databases
$databasesRaw = ( & $nwAdminExe.Source -o CheckDatabaseVersion ) -join "`n"
$message = "`nNintex Workflow Databases:`n$databasesRaw"
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug

if ( $databasesRaw -match 'Command line execution error: Failed to open a connection to the Nintex Workflow configuration database' )
{
    # TODO: Add the database and instance or connection string info to the alert
    $message = "NWAdmin was unable to open a connection to the Nintex Workflow configuration database. Ensure the RunAs account ($whoami) is a member of the 'db_datareader' database role."
    Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 1 -Description $message -DebugLogging:$debug
    exit
}

# Parse the string into an object
$databasesString = $databasesRaw -split "`n`n"
foreach ( $databaseString in $databasesString )
{
    $databaseProperties = $databaseString -split "`n"
    $databaseConnectionString = New-Object -TypeName System.Data.SqlClient.SqlConnectionStringBuilder -ArgumentList $databaseProperties[0]
    $databaseType = $databaseProperties[1].Replace('Type: ','')
    $databaseVersion = [System.Version]::new($databaseProperties[2].Replace('Version: ',''))
    $databaseStatus = $databaseProperties[3]

     switch ( $databaseType )
     {
         Configuration { $discoveryType = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Configuration.Class"]$' }
         Content { $discoveryType = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Content.Class"]$' }
     }

     $discoveryData += @{
         DiscoveryType = $discoveryType
         Properties = @(
            @{
                ClassName = 'Nintex.Workflow.SharePoint.Class'
                PropertyInstance = $propertyInstance.Farm
                PropertyName = 'Farm'
                PropertyValue = $farm.Name
             }    
            @{
                ClassName = 'Nintex.Workflow.SharePoint.Database.Class'
                PropertyInstance = $propertyInstance.InstanceName
                PropertyName = 'InstanceName'
                PropertyValue = $databaseConnectionString.DataSource
             }
             @{
                ClassName = 'Nintex.Workflow.SharePoint.Database.Class'
                PropertyInstance = $propertyInstance.DatabaseName
                PropertyName = 'DatabaseName'
                PropertyValue = $databaseConnectionString.InitialCatalog
             }
             @{
                ClassName = 'Nintex.Workflow.SharePoint.Database.Class'
                PropertyInstance = $propertyInstance.Version
                PropertyName = 'Version'
                PropertyValue = $databaseVersion.ToString()
             }
         )
     }
}

### Version check

### Web App Exists check
# Get the installed web apps for the farm
$waps = Get-SPWebApplication -IncludeCentralAdministration

# Build a string of web app IDs to pass into the query
$wapIdsString = "'$(( $waps | Select-Object -ExpandProperty Id ) -join "', '")'"

@'
SELECT DISTINCT [w].[WebApplicaitonId]
FROM [dbo].[Workflows] AS [w]
WHERE [WebApplicationId] NOT IN ({0})
'@ -f $wapIdsString

### Site Exists check

### Web Exists check
