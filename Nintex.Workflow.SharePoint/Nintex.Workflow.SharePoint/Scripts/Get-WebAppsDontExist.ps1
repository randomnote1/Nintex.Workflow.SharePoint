# Get the installed web apps for the farm
$waps = Get-SPWebApplication -IncludeCentralAdministration

# Build a string of web app IDs to pass into the query
$wapIdsString = "'$(( $waps | Select-Object -ExpandProperty Id ) -join "', '")'"

@'
SELECT DISTINCT [w].[WebApplicaitonId]
FROM [dbo].[Workflows] AS [w]
WHERE [WebApplicationId] NOT IN ({0})
'@ -f $wapIdsString