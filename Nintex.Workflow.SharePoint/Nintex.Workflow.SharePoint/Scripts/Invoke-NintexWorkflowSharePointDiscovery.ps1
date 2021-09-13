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
    $PrincipalName,

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
    $message = "`n $parameterString `n $($_ | Format-List -Force | Out-String)"
    Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 1 -Description $message -DebugLogging

    throw $message
}

# Log script event that we are starting task
$whoami = whoami
$message = "`nScript is starting.`nRunning As: $whoami`n$parameterString"
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug

# Initialize the DiscoveryData variable
$discoveryData = @()

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

#region Get Nintex Installation Info


$farm = Get-SPFarm
$message = "`nFarm Name: $($farm.Name)"
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug

$nintexVersion = [Nintex.Workflow.Version]::GetVersion()
$message = "`nNintex Version: $($nintexVersion.ToString())"
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug

$discoveryData += @{
    DiscoveryType = '$MPElement[Name="Nintex.Workflow.SharePoint.Class"]$'
    Properties = @(
        @{
            ClassName = 'Nintex.Workflow.SharePoint.CentralAdministration.Class'
            PropertyInstance = '$MPElement[Name="Windows!Microsoft.Windows.Computer"]/PrincipalName$'
            PropertyName = 'PrincipalName'
            PropertyValue = $PrincipalName
        }
        @{
            ClassName = 'System.Entity'
            PropertyInstance = '$MPElement[Name="System!System.Entity"]/DisplayName$'
            PropertyName = 'DisplayName'
            PropertyValue = $farm.Name
        }
        @{
            ClassName = 'Nintex.Workflow.SharePoint.Class'
            PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.Class"]/Farm$'
            PropertyName = 'Farm'
            PropertyValue = $farm.Name
        }
        @{
            ClassName = 'Nintex.Workflow.SharePoint.Class'
            PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.Class"]/Version$'
            PropertyName = 'Version'
            PropertyValue = $nintexVersion.ToString()
        }
    )
}

#endregion Get Nintex Installation Info

$configurationDatabase = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase()
if ( [System.String]::IsNullOrEmpty($configurationDatabase.SQLConnectionString.ToString()) )
{
    $message = "`nNo Nintex Configuration database is configured in the SharePoint farm '$($farm.Name)'."
    Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 2 -Description $message -DebugLogging
}
else
{
    #region Get Nintex Database Info

    $configurationDatabaseVersion = [Nintex.Workflow.Administration.ConfigurationDatabase]::DatabaseVersion
    if ( $configurationDatabaseVersion -eq '0.0.0.0' )
    {
        $message = "`nThe configuration database version was detected as $configurationDatabaseVersion. This indicates the run-as account does not have permissions to the configuration database. Add the run-as account to the 'WSS_Content_Application_Pools' role in the configuration database."
        Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 2 -Description $message -DebugLogging:$debug
    }

    $discoveryData += @{
        DiscoveryType = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Configuration.Class"]$'
        Properties = @(
            @{
                ClassName = 'Nintex.Workflow.SharePoint.CentralAdministration.Class'
                PropertyInstance = '$MPElement[Name="Windows!Microsoft.Windows.Computer"]/PrincipalName$'
                PropertyName = 'PrincipalName'
                PropertyValue = $PrincipalName
            }
            @{
                ClassName = 'Nintex.Workflow.SharePoint.Class'
                PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.Class"]/Farm$'
                PropertyName = 'Farm'
                PropertyValue = $farm.Name
            }   
            @{
                ClassName = 'Nintex.Workflow.SharePoint.Database.Class'
                PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Class"]/InstanceName$'
                PropertyName = 'InstanceName'
                PropertyValue = $configurationDatabase.SQLConnectionString.DataSource
            }
            @{
                ClassName = 'System.Entity'
                PropertyInstance = '$MPElement[Name="System!System.Entity"]/DisplayName$'
                PropertyName = 'DisplayName'
                PropertyValue = $configurationDatabase.SQLConnectionString.InitialCatalog
            }
            @{
                ClassName = 'Nintex.Workflow.SharePoint.Database.Class'
                PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Class"]/DatabaseName$'
                PropertyName = 'DatabaseName'
                PropertyValue = $configurationDatabase.SQLConnectionString.InitialCatalog
            }
            @{
                ClassName = 'Nintex.Workflow.SharePoint.Database.Class'
                PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Class"]/Version$'
                PropertyName = 'Version'
                PropertyValue = $configurationDatabaseVersion
            }
            @{
                ClassName = 'Nintex.Workflow.SharePoint.Database.Class'
                PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Class"]/DatabaseId$'
                PropertyName = 'DatabaseId'
                PropertyValue = 0
            }
        )
    }

    foreach ( $contentDatabase in $configurationDatabase.ContentDatabases )
    {
        if ( [System.String]::IsNullOrEmpty($contentDatabase.DatabaseVersion) )
        {
            $message = "`nThe content database '$($contentDatabase.SQLConnectionString.InitialCatalog)' was not discovered. This indicates the run-as account does not have permissions to the content database. Add the run-as account to the 'WSS_Content_Application_Pools' role in the content database."
            Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 2 -Description $message -DebugLogging:$debug
        }
        
        $discoveryData += @{
            DiscoveryType = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Content.Class"]$'
            Properties = @(
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.CentralAdministration.Class'
                    PropertyInstance = '$MPElement[Name="Windows!Microsoft.Windows.Computer"]/PrincipalName$'
                    PropertyName = 'PrincipalName'
                    PropertyValue = $PrincipalName
                }
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.Class'
                    PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.Class"]/Farm$'
                    PropertyName = 'Farm'
                    PropertyValue = $farm.Name
                }
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.Database.Class'
                    PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Class"]/InstanceName$'
                    PropertyName = 'InstanceName'
                    PropertyValue = $contentDatabase.SQLConnectionString.DataSource
                }
                @{
                    ClassName = 'System.Entity'
                    PropertyInstance = '$MPElement[Name="System!System.Entity"]/DisplayName$'
                    PropertyName = 'DisplayName'
                    PropertyValue = $contentDatabase.SQLConnectionString.InitialCatalog
                }
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.Database.Class'
                    PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Class"]/DatabaseName$'
                    PropertyName = 'DatabaseName'
                    PropertyValue = $contentDatabase.SQLConnectionString.InitialCatalog
                }
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.Database.Class'
                    PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Class"]/Version$'
                    PropertyName = 'Version'
                    PropertyValue = $contentDatabase.DatabaseVersion
                }
                @{
                    ClassName = 'Nintex.Workflow.SharePoint.Database.Class'
                    PropertyInstance = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Class"]/DatabaseId$'
                    PropertyName = 'DatabaseId'
                    PropertyValue = $contentDatabase.DatabaseId
                }
            )
        }
    }

    #endregion Get Nintex Database Info
}


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
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $discoveredString.ToString() -DebugLogging:$debug

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

# Log an event for script ending and total execution time.
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
