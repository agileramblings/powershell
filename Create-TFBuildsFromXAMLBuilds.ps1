# http://www.incyclesoftware.com/2014/10/change-build-controller-multiple-build-definitions/

$cs = Get-TfsConfigServer http://ditfssb02:8080/tfs
$tpcIds = Get-TfsTeamProjectCollectionIds $cs

Write-Verbose "Iterating through each TPC"
foreach($tpcId in $tpcIds){
    Write-Verbose "Get TPC instance for $tpcId"
    $tpc = $cs.GetTeamProjectCollection($tpcIds[0])

    $bs = $tpc.GetService([Microsoft.TeamFoundation.Build.Client.IBuildServer])
    $css = $tpc.GetService([Microsoft.TeamFoundation.Server.ICommonStructureService])

    $projects = $css.ListAllProjects();
    $controllers = $bs.QueryBuildControllers($true)
    $dscName = "(DSC) DevSupport Centre"
    $newControllerName = "Ditfssb03 - Controller"
    foreach($proj in $projects){
        $name = $proj.Name
        #$name = $dscName
        Write-Host "Getting definitions for $name"
        $def = $bs.QueryBuildDefinitions($name)

        if ($def -ne $null){
          
            $controller = $bs.GetBuildController($newControllerName)
            $def | % {
                Write-Host "Checking" $_.Name
                Write-Host $_.Name "is using" $_.BuildController.Name

                # get build drop location
                Write-Host $_.DefaultDropLocation

            }
        }
    }
}
