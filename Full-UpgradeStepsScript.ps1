#Master TeamProject Upgrade Script
Set-Tfs2013

#run on VS client 
Write-Host "$(Get-Date -Format g) ---  Start - Deleting Noise Work Items ---"
.\Clean-TfsWorkItems.ps1 -Verbose
Write-Host "$(Get-Date -Format g) ---  Complete - Deleting Noise Work Items ---"

#run on VS client 
Write-Host "$(Get-Date -Format g) ---  Start - Upgrade Process Templates ---"
.\Duplicate-UpdateTfsTeamProjectProcessTemplate.ps1 -Verbose
Write-Host "$(Get-Date -Format g) ---  Complete - Upgrade Process Templates ---"

#must run on AppTier (server)
Write-Host "$(Get-Date -Format g) ---  Start - Configure Team Projects ---"
.\Tfs-UpdateTeamProjectFeatures.ps1
Write-Host "$(Get-Date -Format g) ---  Complete - Configure Team Projects ---"
