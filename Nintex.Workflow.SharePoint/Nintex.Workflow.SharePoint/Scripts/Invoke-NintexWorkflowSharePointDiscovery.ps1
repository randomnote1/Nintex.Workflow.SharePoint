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

# Define a lookup table to relate property names to the class instance
$propertyInstance = @{
    Farm = '$MPElement[Name="Nintex.Workflow.SharePoint.Class"]/Farm$'
    NintexVersion = '$MPElement[Name="Nintex.Workflow.SharePoint.Class"]/Version$'
    InstanceName = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Class"]/InstanceName$'
    DatabaseName = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Class"]/DatabaseName$'
    Version = '$MPElement[Name="Nintex.Workflow.SharePoint.Database.Class"]/Version$'
}

# Initialize the DiscoveryData variable
$discoveryData = @()
#endregion initialize script

#region Get Nintex Installation Info
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

[System.Reflection.Assembly]::Load('Nintex.Workflow, Version=1.0.0.0, Culture=neutral, PublicKeyToken=913f6bae0ca5ae12') > $null
$nintexVersion = [Nintex.Workflow.Version]::GetVersion()
$message = "`nNintex Version: $($nintexVersion.ToString())"
Write-OperationsManagerEvent @writeOperationsManagerEventParams -Severity 0 -Description $message -DebugLogging:$debug

$discoveryData += @{
    DiscoveryType = '$MPElement[Name="Nintex.Workflow.SharePoint.Class"]$'
    Properties = @(
        @{
            ClassName = 'Nintex.Workflow.SharePoint.Class'
            PropertyInstance = $propertyInstance.Farm
            PropertyName = 'Farm'
            PropertyValue = $farm.Name
        }
        @{
            ClassName = 'Nintex.Workflow.SharePoint.Class'
            PropertyInstance = $propertyInstance.NintexVersion
            PropertyName = 'Version'
            PropertyValue = $nintexVersion.ToString()
        }
    )
}

#endregion Get Nintex Installation Info

#region Get Nintex Database Info

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

#endregion Get Nintex Database Info

#region Create Discovery Data

if (-not $TestRun)
{
    Write-Debug -Message 'Generating discovery object to return to SCOM'

    # Initialize SCOM discovery data object
    $scomDiscoveryData = $momapi.CreateDiscoveryData(0, $SourceId, $ManagedEntityId)

    # Create an instance of the Nintex Workflow for SharePoint installation to create a relationship with the databases
    #$installationInstance = $scomDiscoveryData.CreateClassInstance('$MPElement[Name="Nintex.Workflow.SharePoint.Class"]$')

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

if ($TestRun)
{
    # Return the object containing the discovered information
    return $discoveryData
}
else
{
    $scomDiscoveryData
}
