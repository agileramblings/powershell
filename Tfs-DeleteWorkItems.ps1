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

        #Get WorkItemStore
        $wiService = New-Object "Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore" -ArgumentList $tpc
        
        #Get a list of TeamProjects
        $tps = $wiService.Projects

        #iterate through the TeamProjects
        foreach ($tp in $tps)
        { 
            #most recent work item change
            $wiql = "SELECT [System.Id], [Changed Date] FROM WorkItems WHERE [System.TeamProject] = '$($tp.name)' ORDER BY [Changed Date] DESC"
            $wiQuery = New-Object -TypeName "Microsoft.TeamFoundation.WorkItemTracking.Client.Query" -ArgumentList $wiService, $wiql
            $results = $wiQuery.RunQuery()
            
            $matches = $results | ? {$_.Title.ToLower().Contains("build failure in build")}
            $idList = $matches | % {$_.Id}
            $ids = [string]::Join(",", $idList)
            $witList = witadmin destroywi /collection:$($tpc.Name)  /id:$ids /noprompt
        }
    }
}