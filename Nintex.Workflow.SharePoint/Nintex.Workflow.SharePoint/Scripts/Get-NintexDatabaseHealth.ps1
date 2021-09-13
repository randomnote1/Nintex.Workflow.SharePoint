param
(
    [Parameter()]
	[System.String]
	$DebugLogging = 'false',

    [Parameter()]
    [Switch]
    $TestRun
)

#region initialize script

$debug = [System.Boolean]::Parse($DebugLogging)
$parameterString = $PSBoundParameters.GetEnumerator() | ForEach-Object -Process { "`n$($_.Key) => $($_.Value)" }

# Enable Write-Debug without inquiry when debug is enabled
if ($debug -or $DebugPreference -ne 'SilentlyContinue')
{
    $DebugPreference = 'Continue'
}

# Import the helper functions
if ( -not $TestRun )
{
    . '$FileResource[Name="Nintex.Workflow.SharePoint.HelperFunctions"]/Path$'
}
else
{
    $helperFunctionsPath = Join-Path -Path $PSScriptRoot -ChildPath HelperFunctions.ps1
    . $helperFunctionsPath
}

$scriptName = 'Get-NintexDatabaseHealth.ps1'
$scriptEventID = 17083 # randomly generated for this script

# Gather the start time of the script
$startTime = Get-Date

# If TestRun is specified, skip loading MOM API
if (-not $TestRun)
{
    # Load MOMScript API
    $momapi = New-Object -comObject MOM.ScriptAPI
}

# Set up parameters to use for all logging in this script
$writeOperationsManagerEventParams = @{
    ScriptName = $scriptName
    EventID = $scriptEventID
    TestRun = $TestRun.IsPresent
}

trap
{
    $message = "`n $parameterString `n $($_ | Format-List -Force | Out-String)"
    Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 1 -Description $message -DebugLogging

    throw $message
}

# Log script event that we are starting task
$whoami = whoami
$message = "`nScript is starting.`nRunning As: $whoami`n$parameterString"
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug

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

Import-NintexWorkflowAssembly

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

#endregion initialize script

# Get the Nintex configuration database
$configurationDatabase = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase()

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
$results = @()
foreach ( $databaseString in $databasesString )
{
    $databaseProperties = $databaseString -split "`n"
    $databaseConnectionString = New-Object -TypeName System.Data.SqlClient.SqlConnectionStringBuilder -ArgumentList $databaseProperties[0]
    $databaseType = $databaseProperties[1].Replace('Type: ','')
    $databaseVersion = [System.Version]::new($databaseProperties[2].Replace('Version: ',''))
    $databaseStatus = $databaseProperties[3]
    $health = 'Healthy'

    if ( $databaseStatus -match '^Requires update' )
    {
        $health = 'Critical'
    }

    switch ( $databaseType )
    {
        Configuration
        {
            $databaseId = 0
        }
        
        Content
        {
            $contentDatabase = $configurationDatabase.ContentDatabases |
                Where-Object -FilterScript { $_.SqlConnectionString.DataSource -eq $databaseConnectionString.DataSource -and $_.SqlConnectionString.InitialCatalog -eq $databaseConnectionString.InitialCatalog }
            $databaseId = $contentDatabase.DatabaseId
        }
    }

    $results += @{
        InstanceName = $databaseConnectionString.DataSource
        DatabaseName = $databaseConnectionString.InitialCatalog
        DatabaseVersionStatus = $databaseStatus
        DatabaseId = $databaseId
        Health = $health
        Monitor = 'Nintex.Workflow.SharePoint.Database.Version'
    }
}

<#
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
#>

if ( $debug )
{
	$i = 0
	$bagsString = $results | ForEach-Object -Process { $i++; $_.GetEnumerator() } | ForEach-Object -Process { "`nBag $i : $($_.Key) => $($_.Value)" }
	$message = "`nProperty bag values: $bagsString"
    Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug
}

if ( -not $TestRun.IsPresent )
{
    foreach ( $result in $results )
    {
	    # Create and fill the property bag
	    $bag = $momapi.CreatePropertyBag()
	    foreach ( $returnValue in $result.GetEnumerator() )
	    {
		    $bag.AddValue($returnValue.Key,$returnValue.Value)
	    }

	    # Return the property bag
	    #$momapi.Return($bag)
	    $bag
    }
}

# Log an event for script ending and total execution time.
$endTime = Get-Date
$scriptTime = ($endTime - $startTime).TotalSeconds

$message = "`n Script Completed. `n Script Runtime: ($scriptTime) seconds."
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug
