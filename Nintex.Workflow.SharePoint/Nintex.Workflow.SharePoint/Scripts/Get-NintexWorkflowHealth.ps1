param
(
    [Parameter(Mandatory = $true)]
    [System.Int32]
    $WarningThreshold,

    [Parameter(Mandatory = $true)]
    [System.Int32]
    $CriticalThreshold,

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

$scriptName = 'Get-NintexWorkflowHealth.ps1'
$scriptEventID = 17082 # randomly generated for this script

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

#endregion initialize script

#region get workflow log entries

#$defaultZone = [Microsoft.SharePoint.Administration.SPUrlZone]::Default

# Get the installed web apps for the farm
#$waps = Get-SPWebApplication -IncludeCentralAdministration

# Build a string of web app IDs to pass into the query
#$wapIdsString = "'$(( $waps | Select-Object -ExpandProperty Id ) -join "', '")'"

$getLogEntriesPerWorkflowCountQuery = @'
SELECT
	[w].[WebApplicationId],
	[wi].[SiteId],
	[wi].[WebId],
	[wi].[ListID],
    [wi].[WorkflowID],
	[wi].[WorkflowName],
	COUNT(0) AS [LogEntriesPerWorkflowCount]
FROM [dbo].[Workflows] AS [w] WITH (NOLOCK)
INNER JOIN [dbo].[WorkflowInstance] AS [wi] WITH (NOLOCK) ON [w].[WorkflowId] = [wi].[WorkflowID]
INNER JOIN [dbo].[WorkflowProgress] AS [wp] WITH (NOLOCK) ON [wi].[InstanceID] = [wp].[InstanceID]
--WHERE [w].[WebApplicationId] IN ({2})
GROUP BY [w].[WebApplicationID], [wi].[SiteId], [wi].[WebID], [wi].[ListID], [wi].[WorkflowID], [wi].[WorkflowName]
HAVING COUNT(0) > {0} OR COUNT(0) > {1}
ORDER BY [LogEntriesPerWorkflowCount] DESC
'@ -f $WarningThreshold,$CriticalThreshold,$wapIdsString
$message = "`nQuery to get the in count of log entries per workflow:`n$getLogEntriesPerWorkflowCountQuery"
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug

$workflowLogEntryStats = @()
$configurationDatabase = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase()
foreach ( $contentDatabase in $configurationDatabase.ContentDatabases )
{
    $sqlConnection = Connect-SQL -ConnectionString $contentDatabase.SQLConnectionString.ToString()
    $workflowLogEntryStats += Invoke-SqlQuery -Connection $sqlConnection -QueryString $getLogEntriesPerWorkflowCountQuery -WithResults
}

# Lookup the web information once per site/web
$webs = $workflowLogEntryStats |
    Select-Object -Property SiteId,WebId -Unique |
    ForEach-Object -Process {
        Get-SPWeb -Identity $_.WebId -Site $_.SiteId -ErrorAction SilentlyContinue
    }

foreach ( $workflowLogEntryStat  in $workflowLogEntryStats )
{
    $web = $webs | Where-Object -FilterScript {
        $_.Id -eq $workflowLogEntryStat.WebId -and
        $_.Site.Id -eq $workflowLogEntryStat.SiteId -and
        $_.Site.WebApplication.Id -eq $workflowLogEntryStat.WebApplicationId
    }

    $healthState = 'Healthy'
    if ( $workflowLogEntryStat.LogEntriesPerWorkflowCount -gt $CriticalThreshold )
    {
        $healthState = 'Critical'
    }
    elseif ( $workflowLogEntryStat.LogEntriesPerWorkflowCount -gt $WarningThreshold )
    {
        $healthState = 'Warning'
    }

    $propertiesToAdd = @{
        HealthState = $healthState
    }

    $site = $web.Site
    $webApplication = $web.Site.WebApplication

    $propertiesToAdd.Add('Web Application Name', $webApplication.Name)
    $propertiesToAdd.Add('Web Application Farm ID', $webApplication.Farm.Id.Guid)
    $propertiesToAdd.Add('Site URL', $site.Url)
    $propertiesToAdd.Add('Web URL', $web.Url)

    # If this workflow targets a List
    if ( $workflowLogEntryStat.ListID -ne '00000000-0000-0000-0000-000000000000' )
    {
        $list = $web.Lists | Where-Object -Property ID -eq $workflowLogEntryStat.ListID
        $propertiesToAdd.Add('List Title',$list.Title)

    }

    foreach ( $propertyToAdd in $propertiesToAdd.GetEnumerator() )
    {
        Add-Member -InputObject $workflowLogEntryStat -MemberType NoteProperty -Name $propertyToAdd.Key -Value $propertyToAdd.Value -Force
    }

    Remove-Variable -Name web,site,webApplication -ErrorAction SilentlyContinue
}

#endregion get workflow log entries

#region return results

$sortProperties = @(
    'Web Application Name'
    'Site URL'
    'Web URL'
    'List Title'
    'WorkflowName'
)

$selectProperties = @(
    'WebApplicationId'
    'Web Application Name'
    'Web Application Farm ID'
    'SiteId'
    'Site URL'
    'WebId'
    'Web URL'
    'ListID'
    'List Title'
    'WorkflowID'
    'WorkflowName'
    'LogEntriesPerWorkflowCount'
)

$criticalStats = $workflowLogEntryStats | Sort-Object -Property $sortProperties | Where-Object -Property HealthState -EQ -Value 'Critical'
$warningStats = $workflowLogEntryStats | Sort-Object -Property $sortProperties | Where-Object -Property HealthState -EQ -Value 'Warning'

# Determine the health state
if ( $criticalStats.Count -gt 0 )
{
    $healthState = 'Critical'
}
elseif ( $warningStats.Count -gt 0 )
{
    $healthState = 'Warning'
}
else
{
    $healthState = 'Healthy'
}

$returnValues = @{
    DatabaseName = $configurationDatabase.SQLConnectionString.InitialCatalog
    Monitor = 'WorkflowLogEntries'
    HealthState = $healthState
}
        
#region build the alert details string

if ( $healthState -ne 'Healthy' )
{
    $alertDetails = [System.Text.StringBuilder]::new()

    if ( $criticalStats.Count -gt 0 )
    {
        $alertDetails.AppendLine('Critical alerts:') > $null
                
        foreach ( $criticalStat in $criticalStats )
        {
            $alertDetails.AppendLine("    - $($criticalStat.HealthState)") > $null
                    
            foreach ( $property in $selectProperties )
            {
                $alertDetails.AppendLine("        > $($criticalStat.$property)") > $null
            }
        }
    }

    $alertDetails.AppendLine('') > $null

    if ( $warningStats.Count -gt 0 )
    {
        $alertDetails.AppendLine('Warning alerts:') > $null
                
        foreach ( $warningStat in $warningStats )
        {
            $alertDetails.AppendLine("    - $($warningStat.HealthState)") > $null
                    
            foreach ( $property in $selectProperties )
            {
                $alertDetails.AppendLine("        > $($property): $($warningStat.$property)") > $null
            }
        }
    }

    $returnValues.Add('AlertDetails', $alertDetails.ToString())
}

#endregion build the alert details string

$i = 0
$bagsString = $returnValues | ForEach-Object -Process { $i++; $_.GetEnumerator() } | ForEach-Object -Process { "`nBag $i : $($_.Key) => $($_.Value)" }
$message = "`nProperty bag values: $bagsString"
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug
        
if (-not $TestRun )
{
    # Create and fill the property bag
    $bag = $momapi.CreatePropertyBag()
    foreach ( $returnValue in $returnValues.GetEnumerator() )
    {
        $bag.AddValue($returnValue.Key,$returnValue.Value)
    }

    # Return the property bag
    #$momapi.Return($bag)
    $bag
}

#endregion return results

# Log an event for script ending and total execution time.
$endTime = Get-Date
$scriptTime = ($endTime - $startTime).TotalSeconds

$message = "`n Script Completed. `n Script Runtime: ($scriptTime) seconds."
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug
