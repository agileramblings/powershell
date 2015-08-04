$configServer = Get-TfsConfigServer "http://ditfssb01:8080/tfs"
$tpcIds = Get-TfsTeamProjectCollectionIds $configServer

$taskWITTitles = @(
 "Setup: Set Permissions" 
,"Setup: Migration of Source Code"
,"Setup: Migration of Work Items"
,"Setup: Set Check-in Policies"
,"Setup: Send mail to users for installation and getting started"
,"Setup: Create Project Structure"
,"Create Vision Statement"
,"Create Configuration Management Plan"
,"Create Personas"
,"Create Quality of Service Requirements"
,"Create Scenarios"
,"Create Project Plan"
,"Create Master Schedule"
,"Create Iteration Plan"
,"Create Test Approach Worksheet"
)

foreach($tpcId in $tpcIds){
    #Get TPC instance
    $tpc = $configServer.GetTeamProjectCollection($tpcId)

    Write-Host "Destroying WI /w Title containing `"$titleTextContains`"" -foregroundcolor Red
                
    #Get WorkItemStore
    $wiService = New-Object "Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore" -ArgumentList $tpc
        
    #Get a list of TeamProjects
    $tps = $wiService.Projects

    $totalWIDestroyed = 0;
    #iterate through the TeamProjects
    foreach ($tp in $tps)
    { 
        Write-host "Destroying WI 'Bug' in TeamProjectCollection $($tpc.Name) in TeamProject $($tp.Name)" -foregroundcolor Yellow
        Remove-TfsWorkItems -configServer $configServer -tpcName  $($tpc.Name) -tpName  $($tp.Name) -witType "Bug" -titleTextContains "build failure in build"
        foreach ($taskTitle in $taskWITTitles)
        {
            Remove-TfsWorkItems -configServer $configServer -tpcName  $($tpc.Name) -tpName  $($tp.Name) -witType "Task" -titleTextContains $taskTitle
        }
    }
}
            