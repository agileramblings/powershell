# Credits to: 
# http://blogs.msdn.com/b/visualstudioalm/archive/2012/05/31/how-to-configure-features-for-dozens-of-team-projects.aspx  - Ewald Hofman
# https://features4tfs.codeplex.com/ - Oleg Mikhaylov
# https://gallery.technet.microsoft.com/scriptcenter/Invoke-Generic-Methods-bf7675af#content - Dave Wyatt

Clear-Host

function Import-TFS2015 {
    try{
        # server OM bin folder
        # "C:\Program Files\Microsoft Team Foundation Server 12.0\Application Tier\Web Services\bin"
        Add-Type -LiteralPath 'C:\Program Files\Microsoft Team Foundation Server 14.0\Application Tier\Web Services\bin\Microsoft.TeamFoundation.Framework.Server.dll'
        Add-Type -LiteralPath 'C:\Program Files\Microsoft Team Foundation Server 14.0\Application Tier\Web Services\bin\Microsoft.TeamFoundation.Server.Core.dll'
        Add-Type -LiteralPath 'C:\Program Files\Microsoft Team Foundation Server 14.0\Application Tier\Web Services\bin\Microsoft.TeamFoundation.Server.WebAccess.WorkItemTracking.Common.dll'
        Add-Type -LiteralPath 'C:\Program Files\Microsoft Team Foundation Server 14.0\Application Tier\Web Services\bin\Microsoft.VisualStudio.Services.Client.dll'
    
        #copy client OM from Visual Studio install on developer machine to Server bin folder noted above
        # "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0"
        Add-Type -LiteralPath 'C:\Program Files\Microsoft Team Foundation Server 14.0\Application Tier\Web Services\bin\Microsoft.TeamFoundation.WorkItemTracking.Client.dll'
        Add-Type -LiteralPath 'C:\Program Files\Microsoft Team Foundation Server 14.0\Application Tier\Web Services\bin\Microsoft.TeamFoundation.Client.dll'
        Add-Type -LiteralPath 'C:\Program Files\Microsoft Team Foundation Server 14.0\Application Tier\Web Services\bin\Microsoft.TeamFoundation.Common.dll'
    }
    catch [Exception]
    {
        echo $_.Exception|format-list -force
    }
}

function Get-TfsDbConnectionString(){
    [CmdLetBinding()]
    param(
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$location
    )
    begin{}
    process {
        Write-Host "------ Getting DB Connection String from Web Config located at $location ..."
        if (![System.IO.File]::Exists($location)) {
            Write-Host "Invalid web.config location: $location"
            exit;
        }
        [xml]$webConfig = Get-Content $location
        $node = $webConfig.SelectSingleNode("/configuration/appSettings/add[@key='applicationDatabase']/@value");
        $node.Value
    }
    end{}
}

function Get-TfsDeploymentServiceHost ($urlToCollection, $webConfigLocation) {
    Write-Host "------ Creating Deployment Service Host..."
    $connStrPath = "/Configuration/Database/Framework/ConnectionString"
    $tpc = New-Object 'Microsoft.TeamFoundation.Client.TfsTeamProjectCollection' -ArgumentList $urlToCollection
    $connStr = Get-TfsDbConnectionString($webConfigLocation)
    $instanceId = $tpc.InstanceId

    $deploymentHostProperties = New-Object 'Microsoft.TeamFoundation.Framework.Server.TeamFoundationServiceHostProperties'
    $deploymentHostProperties.ConnectionInfo = [Microsoft.TeamFoundation.Framework.Server.SqlConnectionInfoFactory]::Create($connStr, $null, $null);
    $deploymentHostProperties.HostType = [Microsoft.TeamFoundation.Framework.Server.TeamFoundationHostType]::Application -bor [Microsoft.TeamFoundation.Framework.Server.TeamFoundationHostType]::Deployment
    $dsh = New-Object 'Microsoft.TeamFoundation.Framework.Server.DeploymentServiceHost' -ArgumentList $deploymentHostProperties, $false
    $dsh; $instanceId;
}

function Get-TfsGetContext ($deploymentServiceHost, $instanceId){
    Write-Host "------ Creating Service Context..."
    $requestContext = $deploymentServiceHost.CreateSystemContext($true)
    $tfHost = Invoke-GenericMethod -InputObject $requestContext -MethodName GetService -GenericType 'Microsoft.TeamFoundation.Framework.Server.TeamFoundationHostManagementService'
    $tfHost.BeginRequest($requestContext, $instanceId, [Microsoft.TeamFoundation.Framework.Server.RequestContextType]::ServicingContext)
    $requestContext.Dispose() | Out-Null
}

function Invoke-TfsProvisionProjectFeatures ($context, $project){
    Write-Host "------ Provisioning Features in $project..."
    $projFeatProvServ = Invoke-GenericMethod -InputObject $context -MethodName GetService -GenericType 'Microsoft.TeamFoundation.Server.WebAccess.WorkItemTracking.Common.ProjectFeatureProvisioningService'
    $needsUpgrade = $projFeatProvServ.GetFeatures($context, $project.Uri) | ? { ($_.State -eq [Microsoft.TeamFoundation.Server.WebAccess.WorkItemTracking.Common.ProjectFeatureState]::NotConfigured) -and (!$_.IsHidden) }
    if ($needsUpgrade.Count -eq 0){
        Write-Host "$project is up to date"
        $project.Name; $project.Uri; $ptName; "Already Upgraded";
    }
    else
    {
      $projFeatProvDetails = $projFeatProvServ.ValidateProcessTemplates($context, $project.Uri)
      $msgs = @()
      foreach ($deet in $projFeatProvDetails){

        foreach ($issue in $deet.Issues){
            $msgs += $issue.Message
        }
        $msgs | % {Write-Host "IsValid: $($deet.IsValid) - Message: $_"}
      }
      
      $validProcTempDetails = $projFeatProvDetails | ? {$_.IsValid}
      $numValidPTs = $validProcTempDetails.Count

      switch ($numValidPTs)
      {
        0 { 
            Write-Host "$project : No Valid Process Template found." 
            $project.Name; $project.Uri; "No Process Templates"; "Failed";
          }
        1 { 
            $projectFeatureProvisioningDetail = $projFeatProvDetails[0]
            $ptName = $($projectFeatureProvisioningDetail.ProcessTemplateName)
            Write-Host "$project : 1 Valid Process Template found.( $ptName )"
            try{
                $projFeatProvServ.ProvisionFeatures($context, $project.Uri, $projectFeatureProvisioningDetail.ProcessTemplateId)
                Write-Host "$project : Done"
                $project.Name; $project.Uri; $ptName; "Success";
            }catch{
                $project.Name; $project.Uri; $ptName; "Error - $error";
            }
          }
        default {
            Write-Host "$project : Multiple Valid Process Templates found."
            $validProcTempDetails | % { Write-Host "$_.ProcessTemplateName "} | Out-Null
            $project.Name; $project.Uri; "Multiple Process Templates"; "Failed";
        }
      }
   }
}

function Update-TfsTeamProjectFeatureConfiguration()
{
    [CmdLetBinding()]
    param(
        [parameter(Mandatory=$true)]
        [string]$tpcUrl
    )
    begin{}
    process{
        # Do a little Authentication to ensure we can do anything
        $configServer = [Microsoft.TeamFoundation.Client.TfsConfigurationServerFactory]::GetConfigurationServer("http://ditfssb01:8080/tfs")
        [void]$configServer.Authenticate()
        if(!$configServer.HasAuthenticated)
        {
            Write-Host "Not Authenticated"
            exit
        }
        else
        {
            Write-Host "Authenticated"

            [string[]] $urls = @('http://ditfssb01:8080/tfs/projectcollection01', 'http://ditfssb01:8080/tfs/projectcollection02')
            [string] $webConfigLocation = 'C:\Program Files\Microsoft Team Foundation Server 14.0\Application Tier\Web Services\web.config'
    
            # create folder for logging artifacts
            if (!(Test-Path -Path C:\TFS\Results\)){
                $null = New-Item -ItemType directory -Path C:\TFS\Results\
            } else {
                #clean out the folder 
                Remove-Item C:\TFS\Results\PT_Upgrade_Log.csv -Force -ErrorAction SilentlyContinue
            }

#            foreach($url in $urls)
#            {
                $myDsh = Get-TfsDeploymentServiceHost $tpcUrl $webConfigLocation
                $ctx = Get-TfsGetContext $myDsh[0] $myDsh[1]
                # get "Microsoft.TeamFoundation.Server.CommonStructureService" service
                $css = Invoke-GenericMethod -InputObject $ctx -MethodName GetService -GenericType 'Microsoft.TeamFoundation.Integration.Server.CommonStructureService'
                foreach($proj in $css.GetWellFormedProjects($ctx)){
                    $retVal = Invoke-TfsProvisionProjectFeatures $ctx $proj

                    #put retVal into csv
                    $values = $(Get-Date -Format g), $tpcUrl
                    $values += $retVal
                    $csvLine = [string]::Join(",", $values)
                    $csvLine >> "C:\TFS\Results\PT_Upgrade_Log.csv";
                }

                $myDsh[0].Dispose()
#            }
        }
    }
    end{}
}

Import-Module GenericMethods
Import-TFS2015
Set-Tfs2015
Update-TfsTeamProjectFeatureConfiguration "http://ditfssb01:8080/tfs/projectcollection02"
Update-TfsTeamProjectFeatureConfiguration "http://ditfssb01:8080/tfs/projectcollection01"
