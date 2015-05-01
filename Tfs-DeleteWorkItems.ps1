Clear-Host

function Get-TfsTeamProjectCollectionIds() {
    [CmdLetBinding()]
    param($configServer)

    process{
        # Get a list of TeamProjectCollections
        [guid[]]$types = [guid][Microsoft.TeamFoundation.Framework.Common.CatalogResourceTypes]::ProjectCollection
        $options = [Microsoft.TeamFoundation.Framework.Common.CatalogQueryOptions]::None
        $configServer.CatalogNode.QueryChildren( $types, $false, $options) | % { $_.Resource.Properties["InstanceId"]}
    }
}

function Get-TfsConfigServer() {
<# 
.SYNOPSIS
  Describe the function here
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Give an example of how to use it
  .EXAMPLE
  Give another example of how to use it
  .PARAMETER url
  The Url of the TFS server that you'd like to access
  .PARAMETER tfsVersion
  The version of the TFS server that you'd like to load the object model of

#>

    [CmdletBinding()]
    param( $url, $tfsVersion)

    begin {
        Write-Verbose "Loading TFS OM Assemblies for $tfsVersion"

        function Import-TFSAssemblies_2013 {
            Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.Client.dll";
            Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.Common.dll";
            #Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.dll";
            Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.VersionControl.Client.dll";
            Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.WorkItemTracking.Client.dll";
        }

        Import-TFSAssemblies_2013
    }

    process {
        $retVal = [Microsoft.TeamFoundation.Client.TfsConfigurationServerFactory]::GetConfigurationServer($url)
        [void]$retVal.Authenticate()
        if(!$retVal.HasAuthenticated)
        {
            Write-Host "Not Authenticated"
            $null;
        }
        else
        {
            Write-Host "Authenticated"
            $retVal;
        }
    }

    end {
        Write-Verbose "ConfigurationServer object created."
    }

}


function Delete-TfsWorkItems(){
    [CmdLetBinding()]
    param($url, $tfsVersion, $tpcName, $tpName, $witType, $titleTextContains)

    begin{}

    process {
        $configServer = Get-TfsConfigServer $url $tfsVersion

        if($configServer -eq $null)
        {
            Write-Host "Not Authenticated"
            exit
        }
        else
        {
            Write-Host "Authenticated"
    
            $tpcIds = Get-TfsTeamProjectCollectionIds $configServer

            foreach($tpcId in $tpcIds){
                #Get TPC instance
                $tpc = $configServer.GetTeamProjectCollection($tpcId)

                if (!$tpc.Name.ToLower().Contains(([string]$($tpcName)).ToLower())) { continue }

                Write-Host "Destroying WI /w Title containing `"$titleTextContains`"" -foregroundcolor Red
                Write-host "Destroying WI in TeamProjectCollection $($tpc.Name)" -foregroundcolor Yellow
                
                #Get WorkItemStore
                $wiService = New-Object "Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore" -ArgumentList $tpc
        
                #Get a list of TeamProjects
                $tps = $wiService.Projects

                $totalWIDestroyed = 0;
                #iterate through the TeamProjects
                foreach ($tp in $tps)
                { 
                    if ($tp.Name -ne $tpName) { continue }

                    #most recent work item change
                    $wiql = "SELECT [System.Id], [Changed Date] FROM WorkItems WHERE [System.WorkItemType] = '$witType' AND [System.TeamProject] = '$($tp.name)' ORDER BY [Changed Date] DESC"
                    $wiQuery = New-Object -TypeName "Microsoft.TeamFoundation.WorkItemTracking.Client.Query" -ArgumentList $wiService, $wiql
                    $results = $wiQuery.RunQuery()
                    
                    if ([string]::IsNullOrEmpty($titleTextContains)) {
                        $matches = $results
                    } else {
                        $matches = $results | ? {$_.Title.ToLower().Contains($titleTextContains)}
                    }
                    $idList = $matches | % {$_.Id}
                    if ($idList.Count -gt 0){
                        $ids = [string]::Join(",", $idList)
                        witadmin destroywi /collection:$($tpc.Name)  /id:$ids /noprompt | Out-Null
                        $totalWIDestroyed += $idList.Count
                    }
                    Write-Host "$($tp.Name) - $witType WI Destroyed: $($idList.Count)"
                }
                Write-Host "Total WI Destroyed: $totalWIDestroyed" -foregroundcolor Yellow
            }
        }
    }

    end{}
}

Delete-TfsWorkItems "http://divcd83:8080/tfs/" "2013.4" "ProjectCollection01" "(APM_CAR) APM - Calgary Awards Rewrite" "Bug" "build failure in build"
Delete-TfsWorkItems "http://divcd83:8080/tfs/" "2013.4" "ProjectCollection01" "(APM_CAR) APM - Calgary Awards Rewrite" "Task"
