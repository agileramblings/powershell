# http://www.incyclesoftware.com/2014/10/change-build-controller-multiple-build-definitions/

$cs = Get-TfsConfigServer http://ditfssb01:8080/tfs
$tpcIds = Get-TfsTeamProjectCollectionIds $cs

Write-Verbose "Iterating through each TPC"
foreach($tpcId in $tpcIds){
    Write-Verbose "Get TPC instance for $tpcId"
    $tpc = $cs.GetTeamProjectCollection($tpcId)

    $bs = $tpc.GetService([Microsoft.TeamFoundation.Build.Client.IBuildServer])
    $css = $tpc.GetService([Microsoft.TeamFoundation.Server.ICommonStructureService])

    $projects = $css.ListAllProjects();
    $controllers = $bs.QueryBuildControllers($true)
    $dscName = "(DSC) DevSupport Centre"
    $newControllername =  if ($tpc.Name.Contains("ProjectCollection01")) { "Ditfssb03 - Controller" } else { "Ditfssb04 - Controller" }
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

#                # Update not to create work items on failure
#                # This is not possible this way - need to explore different way
#                $buildPT = $_.Process
#                [xml]$buildPTXml = $buildPT.Parameters
#                $buildPTXml.Activity.'Process.CreateWorkItem' = "[False]"
#                $buildPT.Parameters = $buildPTXml

                # update drop location
                $_.DefaultDropLocation = "\\coc\it\gis-tfs\qa\" + $name 

                if ($_.BuildController.Uri -ne $controller.Uri) {
                    Write-Host "Setting" $_.Name "to use $newControllerName"
                    $_.BuildController = $controller
                }
                else {
                    Write-Host "Build controller is already set. Taking no action."
                }
                
                if (!$WhatIf) {
                    # update controller
                    #Write-Host "Stop here"
                    $_.Save()
                }
            }
        }
    }
}
