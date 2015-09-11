#load DscTfs module
impo DscTfs
$cs = Get-TfsConfigServer http://pitfs02:8080/tfs


#run on VS client
#this script will destroy all of the bug workitems that are created when a build fails (noise) 
#and will destroy a set of templated task work items that were created when a team project was created in a bunch of TPs
Write-Host "$(Get-Date -Format g) ---  Start - Deleting Noise Work Items ---"
.\Clean-TfsWorkItems.ps1 -Verbose
Write-Host "$(Get-Date -Format g) ---  Complete - Deleting Noise Work Items ---"

#run anywhere
#this will backup all of the remaining work items to CSV files in the folder location below
Backup-TfsWorkItems $cs -rootFolder "\\exchange\exchange\dwhite2\TFS Production Upgrade\WorkItemBackup\"

#run on VS client
#this script will upgrade all TeamProjects to modern version of their Process Template
Write-Host "$(Get-Date -Format g) ---  Start - Upgrade Process Templates ---"
.\Update-TfsTeamProjectProcessTemplate.ps1 -Verbose
Write-Host "$(Get-Date -Format g) ---  Complete - Upgrade Process Templates ---"

#run anwyhere
Write-Host "$(Get-Date -Format g) ---  Start - Configure Team Projects ---"
.\Automate-IEFeatureConfiguration.ps1
Write-Host "$(Get-Date -Format g) ---  Complete - Configure Team Projects ---"
