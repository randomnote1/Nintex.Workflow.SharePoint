param
(
    [Parameter(Mandatory = $true)]
    [System.String]
    $SourceId,

    [Parameter(Mandatory = $true)]
    [System.String]
    $ManagedEntityId,

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

$scriptName = 'Invoke-NintexWorkflowSharePointDiscovery.ps1'
$scriptEventID = 17081 # randomly generated for this script

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
    $message = "`n $parameterString `n $($_.ToString())"
    Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 1 -Description $message -DebugLogging

    throw $message
}

# Log script event that we are starting task
$whoami = whoami
$message = "`nScript is starting.`nRunning As: $whoami`n$parameterString"
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug

#endregion initialize script

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
    $message = "NWAdmin was unable to open a connection to the Nintex Workflow configuration database. Ensure the RunAs account ($whoami) is a member of the 'WSS_Content_Application_Pools' database role."
    Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 1 -Description $message -DebugLogging:$debug
    exit
}

# Parse the string into an object
$databasesString = $databasesRaw -split "`n`n"
$databases = @()
foreach ( $databaseString in $databasesString )
{
    $databaseProperties = $databaseString -split "`n"
    $databases += New-Object -TypeName PSCustomObject -Property @{
        ConnectionString = $databaseProperties[0]
        Type = $databaseProperties[1].Replace('Type: ','')
        Version = [System.Version]::new($databaseProperties[2].Replace('Version: ',''))
        Comment = $databaseProperties[3]
    }
}

# Log an event for script ending and total execution time.
$endTime = Get-Date
$scriptTime = ($endTime - $startTime).TotalSeconds

$message = "`n Script Completed. `n Script Runtime: ($scriptTime) seconds."
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug
