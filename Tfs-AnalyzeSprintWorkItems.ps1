﻿param([string]$titleText="build failure in build") #Must be the first statement in your script

function Import-TFSAssemblies_2013 {
    Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.Client.dll";
    Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.Common.dll";
    Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.dll";
    Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.VersionControl.Client.dll";
    Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.WorkItemTracking.Client.dll";
}

function Get-TfsTeamProjectCollectionIds ($configServer) {
    # Get a list of TeamProjectCollections
    [guid[]]$types = [guid][Microsoft.TeamFoundation.Framework.Common.CatalogResourceTypes]::ProjectCollection
    $options = [Microsoft.TeamFoundation.Framework.Common.CatalogQueryOptions]::None
    $configServer.CatalogNode.QueryChildren( $types, $false, $options) | % { $_.Resource.Properties["InstanceId"]}
}

Import-TFSAssemblies_2013

Clear-Host

$configServer = [Microsoft.TeamFoundation.Client.TfsConfigurationServerFactory]::GetConfigurationServer("http://divcd83:8080/tfs")
[void]$configServer.Authenticate()
if(!$configServer.HasAuthenticated)
{
    Write-Host "Not Authenticated"
    exit
}
else
{
    Write-Host "Authenticated"
    
    $tpcIds = Get-TfsTeamProjectCollectionIds($configServer)

    foreach($tpcId in $tpcIds){
        #Get TPC instance
        $tpc = $configServer.GetTeamProjectCollection($tpcId) 
        Write-Host "Looking for Sprint Goal values in Sprint WI" -foregroundcolor Red
        #Get WorkItemStore
        $wiService = New-Object "Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore" -ArgumentList $tpc
        
        #Get a list of TeamProjects
        $tps = $wiService.Projects

        $totalWIDestroyed = 0;
        #iterate through the TeamProjects
        foreach ($tp in $tps)
        { 
            #most recent work item change
            $wiql = "SELECT [System.Id], [Changed Date] FROM WorkItems WHERE [System.WorkItemType] = 'Sprint' AND [System.TeamProject] = '$($tp.name)' ORDER BY [Changed Date] DESC"
            $wiQuery = New-Object -TypeName "Microsoft.TeamFoundation.WorkItemTracking.Client.Query" -ArgumentList $wiService, $wiql
            $results = $wiQuery.RunQuery()
            $matches = $results | ? {
                if ($_.Fields.Name.Contains("Retrospective")){
                    $_.Fields["Retrospective"].Value -ne "<h5>What worked?</h5><h5>What didn't work?</h5><h5>What will we do differently?</h5>"
                } elseif ($_.Fields.Name.Contains("Description")){
                    $_.Fields["Description"].Value -ne "<h5>What worked?</h5><h5>What didn't work?</h5><h5>What will we do differently?</h5>"
                }
            }
            Write-Host "Found Sprint Work Items $($results.Count) - Found non-standard notes $($matches.Count)" -NoNewLine

            if (($tp.Name -eq "CAAF") -or ($tp.Name.Contains("PTWeb"))){
                Write-Host "CAAF - STOP HERE"
            }
            if ($matches.Count -gt 0) {
                Write-Host " Found retrospective notes - $($tp.Name)" -foregroundcolor Green
                $matches | % {Write-Host $_.Fields["Retrospective"].Value}
            }else {
                Write-Host ""
            }
        }
    }
}