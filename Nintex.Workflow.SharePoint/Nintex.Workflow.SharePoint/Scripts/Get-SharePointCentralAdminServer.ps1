param
(
    [Parameter(Mandatory = $true)]
    [System.String]
    $SourceId,

    [Parameter(Mandatory = $true)]
    [System.String]
    $ManagedEntityId,

    [Parameter(Mandatory = $true)]
	[System.String]
    $HostServer,
    
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

$scriptName = 'Get-SharePointCentralAdminServer.ps1'
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

if ( Get-Module -Name IISAdministration -ListAvailable )
{
    Import-Module -Name IISAdministration -ErrorAction Stop
}
else
{
    Import-Module -Name WebAdministration -ErrorAction Stop
}

#endregion initialize script

#region Determine if this is the CA server

$discoveryData = @()

$caName = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local.DisplayName
$caUrl = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local.URL
$caFarmName = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local.Farm.Name

if ( Get-Command -Name Get-IISSite )
{
    $caIisSite = Get-IISSite -Name $caName
}
else
{
    $caIisSite = Get-WebSite -Name $caName
}

if ( $caIisSite )
{
    if ( $caUrl -match $caIisSite.Bindings.Host )
    {
        $defaultZone = [Microsoft.SharePoint.Administration.SPUrlZone]::Default
        
        $discoveryData += @{
            DiscoveryType = '$MPElement[Name="Nintex.Workflow.SharePoint.CentralAdministration.Class"]$'
            Properties = @(
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.CentralAdministration.Class'
                    PropertyInstance = '$MPElement[Name="System!System.Entity"]/DisplayName$'
                    PropertyName = 'DisplayName'
                    PropertyValue = "$caFarmName Central Administration"
                }
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.CentralAdministration.Class'
                    PropertyInstance = '$MPElement[Name="Windows!Microsoft.Windows.Computer"]/PrincipalName$'
                    PropertyName = 'PrincipalName'
                    PropertyValue = $HostServer
                }
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.CentralAdministration.Class'
                    PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.CentralAdministration.Class"]/HostServer$'
                    PropertyName = 'HostServer'
                    PropertyValue = $HostServer
                }
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.CentralAdministration.Class'
                    PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.CentralAdministration.Class"]/FarmID$'
                    PropertyName = 'FarmID'
                    PropertyValue = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local.Farm.ID.Guid
                }
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.CentralAdministration.Class'
                    PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.CentralAdministration.Class"]/FarmName$'
                    PropertyName = 'FarmName'
                    PropertyValue = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local.Farm.Name
                }
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.CentralAdministration.Class'
                    PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.CentralAdministration.Class"]/AppPath$'
                    PropertyName = 'AppPath'
                    PropertyValue = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local.IisSettings[$defaultZone].Path.ToString()
                }
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.CentralAdministration.Class'
                    PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.CentralAdministration.Class"]/ResponseUri$'
                    PropertyName = 'ResponseUri'
                    PropertyValue = $caUrl
                }
            )
        }
    }
}

if ( $discoveryData.Count -eq 0 )
{
    $message = "`nThe computer $HostServer is not running the Central Administration role."
    Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug
}

#endregion Determine if this is the CA server

#region Create Discovery Data

$discoveredString = New-Object -TypeName System.Text.StringBuilder
$d = 1
foreach ( $discovered in $discoveryData )
{
    $discoveredString.AppendLine("Discovered Item $d") > $null
    $discoveredString.AppendLine("    - Discovery Type: $($discovered.DiscoveryType)") > $null
    $discoveredString.AppendLine("    - Properties:") > $null

    $p = 1
    foreach ( $property in $discovered.Properties )
    {
        $discoveredString.AppendLine("        + Property $p") > $null
        $discoveredString.AppendLine("            > ClassName: $($property.ClassName)") > $null
        $discoveredString.AppendLine("            > PropertyInstance: $($property.PropertyInstance)") > $null
        $discoveredString.AppendLine("            > PropertyName: $($property.PropertyName)") > $null
        $discoveredString.AppendLine("            > PropertyValue: $($property.PropertyValue)") > $null
        $discoveredString.AppendLine('') > $null
        $p++
    }

    $discoveredString.AppendLine('') > $null
    $d++
}
$message = "`n$($discoveredString.ToString())"
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug

if (-not $TestRun)
{
    Write-Debug -Message 'Generating discovery object to return to SCOM'

    # Initialize SCOM discovery data object
    $scomDiscoveryData = $momapi.CreateDiscoveryData(0, $SourceId, $ManagedEntityId)

    foreach ( $currentDiscoveryData in $discoveryData )
    {
        # Only return mapped discovery types
        if ($currentDiscoveryData.DiscoveryType -ne 'Unknown')
        {
            Write-Debug -Message "Discovery class --> $($currentDiscoveryData.DiscoveryType)"

            # Initialize the appropriate class
            $discoveredObject = $scomDiscoveryData.CreateClassInstance($currentDiscoveryData.DiscoveryType)

            # Add all additional properties
            foreach ($property in $currentDiscoveryData.Properties)
            {
                Write-Debug -Message "Adding property '$($property.PropertyName) with value $($property.PropertyValue)"
                $discoveredObject.AddProperty($property.PropertyInstance, $property.PropertyValue)
            }

            $scomDiscoveryData.AddInstance($discoveredObject)
        }
    }
}
#endregion Create Discovery Data

#region Log an event for script ending and total execution time.
$endTime = Get-Date
$scriptTime = ($endTime - $startTime).TotalSeconds

$message = "`n Script Completed. `n Script Runtime: ($scriptTime) seconds."
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug

if ($TestRun)
{
    # Return the object containing the discovered information
    return $discoveryData
}
else
{
    $scomDiscoveryData
}
#endregion Log an event for script ending and total execution time.
