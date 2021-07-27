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

$scriptName = 'Set-NintexWorkflowForSharePointRunAsProfile.ps1'
$scriptEventID = 17082 # randomly generated for this script

# Gather the start time of the script
$startTime = Get-Date

# If TestRun is specified, skip loading MOM API
if (-not $TestRun)
{
    # Load MOMScript API
    $momapi = New-Object -comObject MOM.ScriptAPI
}

trap
{
    $message = "`n $parameterString `n $($_.ToString())"

    if (-not $TestRun)
    {
        $momapi.LogScriptEvent($scriptName, $scriptEventID, 1, $message)
    }

    Write-Debug -Message $message

    throw $message
}

# Log script event that we are starting task
if ($debug)
{
    $user = "Running As: $($env:USERDOMAIN)\$($env:USERNAME)"
    $message = "`nScript is starting.`n$user`n$parameterString"

    if (-not $TestRun)
    {
        $momapi.LogScriptEvent($scriptName, $scriptEventID, 0, $message)
    }

    Write-Debug -Message $message
}

#endregion initialize script

#region Main Script

# Get the RunAs Profile
$nintexProfile = Get-SCOMRunAsProfile -Name Nintex.Workflow.SharePoint.RunasProfile

# Get the MOSSInstallation group
$mossInstallationGroup = Get-SCOMGroup -DisplayName Nintex.Workflow.SharePoint.MOSSInstallation.Group

# Get the SharePoint RunAs Profile
$sharePointProfile = Get-SCOMRunAsProfile -Name Microsoft.SharePoint.Library.AdminAccount

# Get the accounts associated with the SharePoint profile and group by name. Selecting unique objects didn't work
$managementGroup = Get-SCOMManagementGroup
$sharePointRunAsAccounts = $managementGroup.GetMonitoringSecureDataHealthServiceReferenceBySecureReferenceId($sharePointProfile.Id) | Group-Object -Property Name
if ($debug)
{
    $sharePointRunAsAccountsString = ( $sharePointRunAsAccounts | Select-Object -ExpandProperty Name | Sort-Object ) -join "`n"
    $message = "`nDiscovered SharePoint RunAs Accounts:`n$sharePointRunAsAccountsString"

    if (-not $TestRun)
    {
        $momapi.LogScriptEvent($scriptName, $scriptEventID, 0, $message)
    }

    Write-Debug -Message $message
}

foreach ( $sharePointRunAsAccount in $sharePointRunAsAccounts )
{
	# Add the account to the Nintex Workflow profile
	Set-SCOMRunAsProfile -Account $sharePointRunAsAccount.Group[0] -Profile $nintexProfile -Group $mossInstallationGroup -Action Add
}

#endregion Main Script

# Log an event for script ending and total execution time.
$endTime = Get-Date
$scriptTime = ($endTime - $startTime).TotalSeconds

if ($debug)
{
    $message = "`n Script Completed. `n Script Runtime: ($scriptTime) seconds."

    if (-not $TestRun)
    {
        $momapi.LogScriptEvent($scriptName, $scriptEventID, 0, $message)
    }

    Write-Debug -Message $message
}
