#Get-TfsConfigServer "http://divcd83:8080/tfs" "2015" | Get-TfsTeamProjectCollectionAnalysis -Folder "C:\temp\Test_Analysis" -Verbose 4> "C:\Temp\Analysis_log.txt"

Write-Host "Loading DscTfs Module"

# Get TFS Object Model
$vsCommon = "Microsoft.VisualStudio.Services.Common"
$commonName = "Microsoft.TeamFoundation.Common"
$clientName = "Microsoft.TeamFoundation.Client"
$VCClientName = "Microsoft.TeamFoundation.VersionControl.Client"
$WITClientName = "Microsoft.TeamFoundation.WorkItemTracking.Client"
$BuildClientName = "Microsoft.TeamFoundation.Build.Client"
$BuildCommonName = "Microsoft.TeamFoundation.Build.Common"
$BuildvNextName = "Microsoft.TeamFoundation.Build2.WebApi"
#$BuildWorkflowName = "Microsoft.TeamFoundation.Build.Workflow"

#symbolic link (folder) for VS binaries
#available after Visual Studio 2015 install
$symbolicLocation = 'C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\'

#witadmin location after VS 2015 Install
$witadmin = "C:\program files (x86)\Microsoft Visual Studio 14.0\common7\ide\witadmin.exe"

#Module folder
$ModuleRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
#where to put TFS Client OM files
$omBinFolder = $("$ModuleRoot\TFSOM\bin\")

function Get-Nuget(){
<# 
    .SYNOPSIS
    This function gets Nuget.exe from the web
    .DESCRIPTION
    This function gets nuget.exe from the web and stores it somewhere relative to the module folder location
#>
    [CmdLetBinding()]
    param()

    begin{}
    process
    {
        #where to get Nuget.exe from
	    $sourceNugetExe = "http://nuget.org/nuget.exe"
    
        #where to save Nuget.exe too
        $targetNugetFolder = New-Folder $("$ModuleRoot\Nuget")
	    $targetNugetExe = $("$ModuleRoot\Nuget\nuget.exe")

        try
        {
            $nugetExe = $targetNugetFolder.GetFiles() | ? {$_.Name -eq "nuget.exe"}
            if ($nugetExe -eq $null){
                Invoke-WebRequest $sourceNugetExe -OutFile $targetNugetExe
            }
        }
        catch [Exception]
        {
            echo $_.Exception|format-list -force
        }

    	Set-Alias nuget $targetNugetExe -Scope Global -Verbose
    }
    end{}
}

function Get-TfsAssembliesFromNuget(){
<# 
    .SYNOPSIS
    This function gets all of the TFS Object Model assemblies from nuget
    .DESCRIPTION
    This function gets all of the TFS Object Model assemblies from nuget and then creates a bin folder of all of the net45 assemblies
    so that they can be referenced easily and loaded as necessary
#>
    [CmdletBinding()]
    param()

    begin{}
    process{
        #clear out bin folder
        $targetOMbinFolder = New-Folder $omBinFolder
        Remove-Item $targetOMbinFolder -Force -Recurse
        $targetOMbinFolder = New-Folder $omBinFolder
        $targetOMFolder = New-Folder $("$ModuleRoot\TFSOM\")

        #get TFS 2015 Object Model assemblies from nuget
        nuget install "Microsoft.TeamFoundationServer.Client" -OutputDirectory $targetOMFolder -ExcludeVersion -NonInteractive
        nuget install "Microsoft.TeamFoundationServer.ExtendedClient" -OutputDirectory $targetOMFolder -ExcludeVersion -NonInteractive
        nuget install "Microsoft.VisualStudio.Services.Client" -OutputDirectory $targetOMFolder -ExcludeVersion -NonInteractive
        nuget install "Microsoft.VisualStudio.Services.InteractiveClient" -OutputDirectory $targetOMFolder -ExcludeVersion -NonInteractive
    
        #move all of the net45 assemblies to a bin folder so we can reference them and they are colocated so that they can find each other
        #as necessary
        $allDlls = Get-ChildItem -Path $("$ModuleRoot\TFSOM\") -Recurse -File -Filter "*.dll"
        
        # Move all the required .dlls out of the nuget folder structure
        #exclude portable dlls
        $requiredDlls = $allDlls | ? {$_.PSPath.Contains("portable") -ne $true } 
        #exclude resource dlls
        $requiredDlls = $requiredDlls | ? {$_.PSPath.Contains("resources") -ne $true } 
        #include net45, native, and Microsoft.ServiceBus.dll
        $requiredDlls = $requiredDlls | ? { ($_.PSPath.Contains("net45") -eq $true) -or ($_.PSPath.Contains("native") -eq $true) -or ($_.PSPath.Contains("Microsoft.ServiceBus") -eq $true) }
        #copy them all to a bin folder
        $requiredDlls | % { Copy-Item -Path $_.Fullname -Destination $targetOMBinFolder}
    }
    end{}

}

function Import-TFSAssemblies() {
<# 
    .SYNOPSIS
    This function imports TFS Object Model assemblies into the PowerShell session
    .DESCRIPTION
    After the TFS 2015 Object Model has been retrieved from Nuget using Get-TfsAssembliesFromNuget function,
    this function will import the necessary (given current functions) assmeblines into the PowerShell session
#>
    [CmdLetBinding()]
    param()

    begin{}
    process
    {
        $omBinFolder = $("$ModuleRoot\TFSOM\bin\");
        $targetOMbinFolder = New-Folder $omBinFolder;

        try { 
            Add-Type -LiteralPath $($targetOMbinFolder.PSPath + $vsCommon + ".dll")
            Add-Type -LiteralPath $($targetOMbinFolder.PSPath + $commonName + ".dll")
            Add-Type -LiteralPath $($targetOMbinFolder.PSPath + $clientName + ".dll")
            Add-Type -LiteralPath $($targetOMbinFolder.PSPath + $VCClientName + ".dll")
            Add-Type -LiteralPath $($targetOMbinFolder.PSPath + $WITClientName + ".dll")
            Add-Type -LiteralPath $($targetOMbinFolder.PSPath + $BuildClientName + ".dll")
            Add-Type -LiteralPath $($targetOMbinFolder.PSPath + $BuildCommonName + ".dll")
            Add-Type -LiteralPath $($targetOMbinFolder.PSPath + $BuildvNextName + ".dll")
        } 
        catch
        {
            $_.Exception.LoaderExceptions | $ { $_.Message }
        }
    }
    end{}
    #Add-Type -LiteralPath $($targetOMbinFolder.PSPath + $BuildWorkflowName + ".dll")
}

[string]$targetVersion = "2015"
[bool]$importCompleted = $false

function New-Folder() {
<# 
    .SYNOPSIS
    This function creates new folders
    .DESCRIPTION
    This function will create a new folder if required or return a reference to the folder that was requested to be created if it already exists.
    .EXAMPLE
    New-Folder "C:\Temp\MyNewFolder\"
    .PARAMETER folderPath
    String representation of the folder path requested
#>

    [CmdLetBinding()]
    param(
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$folderPath
    )
    process {
        if (!(Test-Path -Path $folderPath)){
            New-Item -ItemType directory -Path $folderPath
        } else {
            Get-Item -Path $folderPath
        }
    }           
} #end Function New-Directory

# better Remove-Item functions -- $RC
function Get-Tree() { 
    [CmdLetBinding()]
    param(
        [parameter(Mandatory=$true)] [string]$Path,
		[parameter()] [string] $Include = '*'
    )
	process
	{
	    @(Get-Item $Path -Include $Include) + 
        (Get-ChildItem $Path -Recurse -Include $Include) | 
        sort pspath -Descending -unique
	}
} 

function Remove-Tree() {
	[CmdLetBinding()]
    param(
        [parameter(Mandatory=$true)] [string]$Path,
		[parameter()] [string] $Include = '*'
    ) 
    process
	{
		Get-Tree $Path $Include | Remove-Item -force -recurse
	}
}

function Get-Hash() {
<# 
    .SYNOPSIS
    Get hash code for a string
    .DESCRIPTION
    Get a cryptographically generated hash code for a string
    .EXAMPLE
    Get-Hash "Test"
    .EXAMPLE
    gh "Test"
    .EXAMPLE
    "Test" | gh
    .PARAMETER inputString
    String variable that the hash should be created from
#>

    [CmdLetBinding()]
    param(
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string[]]$inputString
    )

    begin{}
    process{
        Write-Verbose "Creating new System.Security.Cryptography.MD5] instance."
        $md5 = [System.Security.Cryptography.MD5]::Create();
        $asciiEncoding = New-Object -TypeName System.Text.ASCIIEncoding;

        Write-Verbose "Encoding input string as ASCII bytes"
        $inputBytes = $asciiEncoding.GetBytes($inputString);

        Write-Verbose "Computing Hash"
        $hash = $md5.ComputeHash($inputBytes);
        $sb = New-Object -TypeName System.Text.StringBuilder;

        Write-Verbose "Converting hash bytes to string representation"
        foreach($byte in $hash){
            [void]$sb.Append($byte.ToString("x2"));        
        }
        $outputString = $sb.ToString()
        
        Write-Output $outputString;
    }
    end{}
} #end Function Get-Hash

function Switch-ChildNodes() {
<# 
    .SYNOPSIS
    Sort Xml Nodes into alphabetic order
    .DESCRIPTION
    In order for the lexographic comparison of an Xml document, this function will sort an elements child nodes alphabetically
    .EXAMPLE
    Switch-ChildNodes $parentNode
    .PARAMETER parentNode
    The parent Xml Node whose children will be sorted
#>
    param(
        [parameter(Mandatory = $true, ValueFromPipeline=$true)]
        [System.Xml.XmlNode]$parentNode
    )
    begin{}
    process{
        $children = $parentNode.ChildNodes | Sort
        $children | % {$parentNode.RemoveChild($_)} | Out-Null
        $children | % {$parentNode.AppendChild($_)} | Out-Null
        $children | % {Switch-ChildNodes($_)} | Out-Null
    }
    end{}
} #end Function Switch-ChildNodes

function Remove-Nodes() {
<# 
    .SYNOPSIS
    Remove Xml Nodes from a parent node
    .DESCRIPTION
    Remove Xml Nodes from a root node using an XPath expression to select the children to remove
    .EXAMPLE
    Remove-Nodes $parentNode "<XPath expression here>"
    .PARAMETER parentNode
    The parent Xml Node whose children will be removed
    .PARAMETER xpathExpression
    XPath expression used to select child nodes for removal

#>    
    [CmdLetBinding()]
    param(
        [parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$nodeToProcess, 
        [parameter(Mandatory = $true)]
        [string] $xpathExpression
    )
    begin{}
    process{
        $nodesToRemove = $nodeToProcess.SelectNodes($xpathExpression) 
        foreach($node in $nodesToRemove)
        {
            $parentNode = $node.ParentNode
            [void]$parentNode.RemoveChild($node)
            $parentNode.AppendChild($parentNode.OwnerDocument.CreateTextNode("")) | Out-Null
        }
    }
    end{}
} #end Function Remove-Nodes

function Get-TfsConfigServer() {
<# 
    .SYNOPSIS
    Get a Team Foundation Server (TFS) Configuration Server object
    .DESCRIPTION
    The TFS Configuration Server is used for basic authentication and represents a connection to the server that is running Team Foundation Server. 
    .EXAMPLE
    Get-TfsConfigServer "<Url to TFS>" "<Version of TSF Object Model to use>"
    .EXAMPLE
    Get-TfsConfigServer "http://localhost:8080/tfs" "2013.4"
    .EXAMPLE 
    gtfs "http://localhost:8080/tfs" "2013.4"
    .PARAMETER url
    The Url of the TFS server that you'd like to access
    .PARAMETER tfsVersion
    The version of the TFS server that you'd like to load the object model of
#>

    [CmdletBinding()]
    param( 
        [parameter(Mandatory = $true)]
        [string]$url
        )

    begin {
        Write-Verbose "Loading TFS OM Assemblies for $targetVersion"
        Import-TFSAssemblies
        $importCompleted = $true
    }

    process {
        $retVal = [Microsoft.TeamFoundation.Client.TfsConfigurationServerFactory]::GetConfigurationServer($url)
        [void]$retVal.Authenticate()
        if(!$retVal.HasAuthenticated)
        {
            Write-Host "Not Authenticated"
            Write-Output $null;
        }
        else
        {
            Write-Host "Authenticated"
            Write-Output $retVal;
        }
    }

    end {
        Write-Verbose "ConfigurationServer object created."
    }

} #end Function Get-TfsConfigServer

function Get-TfsTeamProjectCollectionIds() {
<# 
    .SYNOPSIS
    Get a collection of Team Project Collection (TPC) Id
    .DESCRIPTION
    Get a collection of Team Project Collection (TPC) Id from the server provided
    .EXAMPLE
    Get-TfsTeamProjectCollectionIds $configServer
    .EXAMPLE
    Get-TfsConfigServer "http://localhost:8080/tfs" "2013.4" | Get-TfsTeamProjectCollectionIds
    .PARAMETER configServer
    The TfsConfigurationServer object that represents a connection to TFS server that you'd like to access
#>

    [CmdLetBinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer
    )
    begin{}
    process{
        # Get a list of TeamProjectCollections
        [guid[]]$types = [guid][Microsoft.TeamFoundation.Framework.Common.CatalogResourceTypes]::ProjectCollection
        $options = [Microsoft.TeamFoundation.Framework.Common.CatalogQueryOptions]::None
        $configServer.CatalogNode.QueryChildren( $types, $false, $options) | % { $_.Resource.Properties["InstanceId"]}
    }
    end{}
} #end Function Get-TfsTeamProjectCollectionIds

#adapted from http://blogs.msdn.com/b/alming/archive/2013/05/06/finding-subscriptions-in-tfs-2012-using-powershell.aspx
function Get-TFSEventSubscriptions
{
    [CmdLetBinding()]
    param(
        [parameter(Mandatory = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer
    )
 
   begin{}
   process{
    $tpcIds = Get-TfsTeamProjectCollectionIds $configServer
    foreach($tpcId in $tpcIds)
    {
        #Get TPC instance
        $tpc = $configServer.GetTeamProjectCollection($tpcId)
        #TFS Services to be used
        $eventService = $tpc.GetService("Microsoft.TeamFoundation.Framework.Client.IEventService")
        $identityService = $tpc.GetService("Microsoft.TeamFoundation.Framework.Client.IIdentityManagementService")
 
        foreach ($sub in $eventService.GetAllEventSubscriptions())
        {
            #First resolve the subscriber ID
            $tfsId = $identityService.ReadIdentity([Microsoft.TeamFoundation.Framework.Common.IdentitySearchFactor]::Identifier, 
                                                   $sub.Subscriber,
                                                   [Microsoft.TeamFoundation.Framework.Common.MembershipQuery]::None,
                                                   [Microsoft.TeamFoundation.Framework.Common.ReadIdentityOptions]::None )
            if ($tfsId.UniqueName)
            {
                $subscriberId = $tfsId.UniqueName
            }
            else
            {
                $subscriberId = $tfsId.DisplayName
            }
 
            #then create custom PSObject
            $subPSObj = New-Object PSObject -Property @{
                            AppTier        = $tpc.Uri
                            ID             = $sub.ID
                            Device         = $sub.Device
                            Condition      = $sub.ConditionString
                            EventType      = $sub.EventType
                            Address        = $sub.DeliveryPreference.Address
                            Schedule       = $sub.DeliveryPreference.Schedule
                            DeliveryType   = $sub.DeliveryPreference.Type
                            SubscriberName = $subscriberId
                            Tag            = $sub.Tag
                            }
 
            #Send object to the pipeline. You could store it on an Arraylist, but that just
            #consumes more memory
            $subPSObj
 
            ##This is another variation where we just add a property to the existing Subscription object
            ##this might be desirable since it will keep the other members
            #Add-Member -InputObject $sub -NotePropertyName SubscriberName -NotePropertyValue $subscriberId
        }
    }
   }
   end{}
}

function Update-TfsXAMLBuildPlatformConfiguration(){
    [CmdLetBinding(SupportsShouldProcess=$true)]
    param(
        [parameter(Mandatory = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer,
        [parameter(Mandatory = $true)]
        [string]$newPlatform,
        [parameter(Mandatory = $false)]
        [string]$tpcName,
        [parameter(Mandatory = $false)]
        [string]$tpName,
        [parameter(Mandatory = $false)]
        [string]$buildName
    )
   begin{}
   process{
        $tpcIds = Get-TfsTeamProjectCollectionIds $configServer
        foreach($tpcId in $tpcIds)
        {
            #Get TPC instance
            $tpc = $configServer.GetTeamProjectCollection($tpcId)
            if (![string]::IsNullorWhiteSpace($tpcName) -and ($tpcName -ne $tpc.Name) ) { continue; }

            $bs = $tpc.GetService([Microsoft.TeamFoundation.Build.Client.IBuildServer])
 
            #Get WorkItemStore
            $wiService = $tpc.GetService([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])
            #Get a list of TeamProjects
            $tps = $wiService.Projects

            #iterate through the TeamProjects
            foreach ($tp in $tps)
            { 
                if( ![string]::IsNullorWhiteSpace($tpName) -and ($tpName -ne $tp.Name) ) { continue;}

                $buildDefinitionList = $bs.QueryBuildDefinitions($tp.Name) 
                $defaultTemplateDefns = $buildDefinitionList | ? { $_.Process.ServerPath.Contains("DefaultTemplate") }
                
                [array]$bdefs = ,
                [array]$actions = @(, @())
                foreach ($bdef in $buildDefinitionList)
                {
                   [xml]$x = [xml]$bdef.ProcessParameters
                   
                   # guards against build defs with no configurations
                   if($x.SelectSingleNode("//PlatformConfigurationList") -eq $null) {continue;}
                   if($x.Dictionary.BuildSettings.'BuildSettings.PlatformConfigurations'.PlatformConfigurationList.Capacity -eq 0) { continue; }
                   
                   $x.Dictionary.BuildSettings.'BuildSettings.PlatformConfigurations'.PlatformConfigurationList.PlatformConfiguration | % {
                        $messages = @()
                        $messages += "The configuration "
                        $messages += "$($_.Configuration)" 
                        $messages += " in build definition " 
                        $messages += "$($bdef.Name) "
                        $messages += " in "
                        $messages += "$($bdef.TeamProject)"
                        $messages += " would have been converted from $($_.Platform) to $newPlatform."
                        $actions += ,$messages

                        $_.Platform = $newPlatform
                   }
                   $bdef.ProcessParameters = $x.OuterXml
                   $bdefs += $bdef
                }
                if (!$WhatIfPreference) {
                    $bs.SaveBuildDefinitions($bdefs)
                } else {
                    if ($actions.Length -le 0) {
                        Write-Host "No changes were potentially made to any build definitions in $($tp.Name) - $($tpc.Name)"
                    }else { 
                        $actions | % { Write-Host $_[0] -NoNewline
                                       Write-Host $_[1] -ForegroundColor Yellow -NoNewline
                                       Write-Host $_[2] -NoNewline
                                       Write-Host $_[3] -ForegroundColor Magenta -NoNewline
                                       Write-Host $_[4] -NoNewline
                                       Write-Host $_[5] -ForegroundColor Green -NoNewline
                                       Write-Host $_[6] } 
                    }
                }
            }
        }
   }
   end{}
}

function Backup-TfsWorkItems() {
<# 
.SYNOPSIS
  Work Item to CSV formatter
  .DESCRIPTION
  Simple function to dump the current state of a work item type to a csv format
  .EXAMPLE
  Backup-TfsWorkItems $configServer "ProjectCollection01" "<Team Project Name>" "Bug"
  .EXAMPLE
  Backup-TfsWorkItems $configServer "ProjectCollection01" "<Team Project Name>" "Task" "C:\TFS\Results\"
  .PARAMETER configServer
   The TfsConfigurationServer object that represents a connection to TFS server that you'd like to access
  .PARAMETER tpcName
  The name of the Team Project Collection that contains the Team Project
  .PARAMETER tpName
  The version of the TFS server that you'd like to load the object model of
  .PARAMETER witType
  The type of the work items that you would like to convert to a CSV format
  .PARAMETER rootFolder
  The folder to which all .csv files and folder structure will be written to

#>
    [CmdLetBinding()]
    param(
        [parameter(Mandatory = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer,
        [parameter(Mandatory = $true)]
        [string]$tpcName,
        [parameter(Mandatory = $true)]
        [string]$tpName,
        [parameter(Mandatory = $true)]
        [string]$witType,
        [parameter(Mandatory = $true)]
        [string]$rootFolder)

    begin{
    }

    process {
        $wiql = "SELECT [System.ID], [System.Title] FROM WorkItems WHERE [System.WorkItemType] = '$witType'  AND [System.TeamProject] = '$tpName' ORDER BY [Changed Date] DESC"
        $tpcIds = Get-TfsTeamProjectCollectionIds $configServer

        foreach($tpcId in $tpcIds){
            #Get TPC instance
            [Microsoft.TeamFoundation.Client.TfsTeamProjectCollection]$tpc = $configServer.GetTeamProjectCollection($tpcId)

            if (!$tpc.Name.ToLower().Contains(([string]$($tpcName)).ToLower())) { continue }
                
            #Get WorkItemStore`
            $wiService = New-Object "Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore" -ArgumentList $tpc
            $wiQuery = New-Object -TypeName "Microsoft.TeamFoundation.WorkItemTracking.Client.Query" -ArgumentList $wiService, $wiql
            $results = $wiQuery.RunQuery()

            if ($results.Count -gt 0) {
                
                $folder = New-Folder $rootFolder
                $folder.CreateSubdirectory("WorkItemArchive\$tpcName\$tpName")

                [PSObject[]]$csv = @()

                foreach($result in $results){
                    $fields = $result.Fields
                    $representation = @{}
                    foreach($field in $fields.Name){
                        try{
                            $representation.Add($field, $result.Item($field))
                        } catch {
                            Write-Host "Failed to add field: $field"
                        }
                    }

                    $output = New-Object PSObject -Property $representation
                    $csv = $csv + $output
                }
                $csvOutput = $csv | ConvertTo-Csv -NoTypeInformation
                $fileName = $($folder.FullName) + "\$witType.csv"
                $csvOutput >> $fileName

            } else {
                Write-Host "There are no work items of type $witType in Team Project $tpName"
            }

        }
    }

    end{}
} #end Function Backup-TfsWorkItems

function Remove-TfsWorkItems() {
<# 
.SYNOPSIS
  Describe the function here
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Delete-TfsWorkItems $configServer "ProjectCollection01" "<Team Project Name>" "Bug" "build failure in build"
  .EXAMPLE
  Delete-TfsWorkItems $configServer "ProjectCollection01" "<Team Project Name>" "Task"
  .PARAMETER configServer
   The TfsConfigurationServer object that represents a connection to TFS server that you'd like to access
  .PARAMETER tpcName
  The Url of the TFS server that you'd like to access
  .PARAMETER tpName
  The version of the TFS server that you'd like to load the object model of
  .PARAMETER witType
  The Url of the TFS server that you'd like to access
  .PARAMETER titleTextContains
  The version of the TFS server that you'd like to load the object model of
#>
    [CmdLetBinding(SupportsShouldProcess=$true, ConfirmImpact="High")]
    param(
        [parameter(Mandatory = $true, ValueFromPipeLine = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer,
        [parameter(Mandatory = $true)]
        [string] $tpcName,
        [parameter(Mandatory = $true)]
        [string] $tpName,
        [parameter(Mandatory = $true)]
        [string] $witType,
        [parameter(Mandatory = $false)]
        [string] $titleTextContains)

    begin{}

    process {
   
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
                    & $witadmin destroywi /collection:$($tpc.Uri.AbsoluteUri) /id:$ids /noprompt | Out-Null
                    $totalWIDestroyed += $idList.Count
                }
                Write-Verbose "$($tp.Name) - $witType WI Destroyed: $($idList.Count)"
            }
            Write-Verbose "Total WI Destroyed: $totalWIDestroyed"
        }
    }

    end{}
} #end Function Delete-TfsWorkItems

function Remove-TfsWorkItemTemplate() {
<# 
.SYNOPSIS
  Describe the function here
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Destroy-TfsWorkItemTemplate "http://<TFS Server>:8080/tfs/" "2013.4" "ProjectCollection01" "<Team Project Name>" "Bug"
  .PARAMETER configServer
   The TfsConfigurationServer object that represents a connection to TFS server that you'd like to access
  .PARAMETER tpcName
  The name of the Team Project Collection on the TFS server that you'd like to access
  .PARAMETER tpName
  The name of the Team Project in the Team Project Collection that you'd like to access
  .PARAMETER witType
  The type of the WIT that you'd like to remove from the Team Project
#>
    [CmdLetBinding(SupportsShouldProcess=$true, ConfirmImpact="High")]
    param(
        [parameter(Mandatory = $true, ValueFromPipeLine = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer,
        [parameter(Mandatory = $true)]
        [string] $tpcName,
        [parameter(Mandatory = $true)]
        [string] $tpName,
        [parameter(Mandatory = $true)]
        [string] $witType)

    begin{
        #ensure we can call witadmin and that the witType exists
        try{
            $tpcUrl =  "$($configServer.Uri.AbsoluteUri)/$tpcName"
            $witList = & $witadmin listwitd /collection:$tpcUrl  /p:$tpName | Sort-Object;
            if (!($witList -contains $witType)){
                Write-Error "$($tpcName)/$($tpName) does not contain a work item type called $witType" -ErrorAction Stop
            }
        }
        catch [Microsoft.PowerShell.Commands.WriteErrorException]{
            Write-Host "The following work items exist in the TeamProject:"
            Write-Host $witList
            throw
        }
        catch { #witadmin failure
            Write-Error "Failed to call witadmin. Please check that your environmental path contains an entry for the location of witadmin." -ErrorAction Stop
        }
    }

    process {
        $tpcIds = Get-TfsTeamProjectCollectionIds $configServer

        foreach($tpcId in $tpcIds){
            #Get TPC instance
            $tpc = $configServer.GetTeamProjectCollection($tpcId)

            if (!$tpc.Name.ToLower().Contains(([string]$($tpcName)).ToLower())) { continue }

            Write-Verbose "Preparing to destroy WIT in TeamProjectCollection $($tpc.Name) - $($witType)" 
                
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
                $wiql = "SELECT [System.Id], [Changed Date] FROM WorkItems WHERE [System.WorkItemType] = '$witType' AND [System.TeamProject] = '$($tp.Name)' ORDER BY [Changed Date] DESC"
                $wiQuery = New-Object -TypeName "Microsoft.TeamFoundation.WorkItemTracking.Client.Query" -ArgumentList $wiService, $wiql
                $results = $wiQuery.RunQuery()
                    
                if ($results.Count -eq 0){
                    & $witadmin destroywitd /collection:$($tpc.Uri.AbsoluteUri)  /p:$($tp.Name) /n:"$witType" /noprompt | Out-Null
                    Write-Verbose "$($tp.Name) - $witType WIT Destroyed"
                    break
                } else {
                    Write-Verbose "$($tp.Name) - $witType has active work items. It has not been destroyed."
                }

            }
            break
        }
    }

    end{}
} #end Function Destroy-TfsWorkItemTemplate

function Import-TfsWorkItemTemplate() {
<# 
  .SYNOPSIS
  Import TFS Work Item Templates
  .DESCRIPTION
  Import a folder of TFS Work Item Template Xml Documents
  .EXAMPLE
  Import-TfsWorkItemTemplate $configServer "ProjectCollection01" "<Team Project Name>" "C:\TFS\WITD_Files"
  .PARAMETER configServer
   The TfsConfigurationServer object that represents a connection to TFS server that you'd like to access
  .PARAMETER tpcName
  The name of the Team Project Collection that contains the Team Project
  .PARAMETER tpName
  The type of the work items that you would like to convert to a CSV format
  .PARAMETER sourcePath
  The folder in which to find WIT Template .xml files
#>

    [CmdLetBinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeLine = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer,
        [parameter(Mandatory = $true)]
        [string] $tpcName,
        [parameter(Mandatory = $true)]
        [string] $tpName,
        [parameter(Mandatory = $true)]
        [string] $sourcePath)

    begin{
        $filesToImport = Get-ChildItem -Path $sourcePath
    }
    process {
       foreach ($file in $filesToImport){
            
            if ($file.Extension -ne ".xml" -or $file.Name.Contains("categories")) { continue }
            
            $fileName = $($file.FullName);
            $tpcUrl =  "$($configServer.Uri.AbsoluteUri)/$tpcName"
            & $witadmin importwitd /collection:$tpcUrl /p:"$tpName" /f:"$fileName"
        }
    }
    end{}
}

function Find-TfsFieldDescription{

    [CmdLetBinding()]
    param(
        [parameter(Mandatory = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer,
        [parameter(Mandatory = $true)]
        [string] $tpcName,
        [parameter(Mandatory = $true)]
        [string] $refNameContains
        )

    begin {}
    process {
        $tpcUrl =  "$($configServer.Uri.AbsoluteUri)/$tpcName"
        $arrayEntries = & $witadmin listfields /collection:$tpcUrl
        $fieldRefList = $arrayEntries | ? {$_.Contains($refnameContains) -and $_.Contains("Field:")}
        $allFields = @();
        foreach($field in $fieldRefList){
            $indexOfField = [array]::indexOf($arrayEntries, $field)
            $listOfValues = @{}
            $fieldEntries = $arrayEntries | Select-Object -Skip $indexOfField | Select-Object -First 7 | % { $_.Trim()}
            foreach($line in $fieldEntries)
            {
                if ([string]::IsNullOrEmpty($line)) {continue;}
                $listOfValues.Add($line.split(":")[0].Trim(), $line.split(":")[1].Trim())
            }
            $allFields += $listofValues
        }
        $allFields;
    }
    end {}
}


function Update-TfsFieldNames {

    [CmdLetBinding()]
    param(
        [parameter(Mandatory = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer,
        [parameter(Mandatory = $true)]
        [string] $tpcName
        )

    begin {}
    process {
        $fieldsToRename = @()
        $fieldsToRename += @{"Name" = "Hyperlink Count"; "Refname" = "System.HyperLinkCount"};
        $fieldsToRename += @{"Name" = "External Link Count"; "Refname" = "System.ExternalLinkCount"};
        $fieldsToRename += @{"Name" = "Related Link Count"; "Refname" = "System.RelatedLinkCount"};
        $fieldsToRename += @{"Name" = "Attached File Count"; "Refname" = "System.AttachedFileCount"};
        $fieldsToRename += @{"Name" = "Area ID"; "Refname" = "System.AreaId"};
        $fieldsToRename += @{"Name" = "Iteration ID"; "Refname" = "System.IterationId"};

        $tpcUrl =  "$($configServer.Uri.AbsoluteUri)/$tpcName"
        foreach($field in $fieldsToRename)
        { 
            & $witadmin changefield /collection:$tpcUrl /n:"$($field.Refname)" /name:"$($field.Name)" /noprompt 
        }
    }
    end {}
}

function Update-TfsWorkItemTemplate() {
<# 
  .SYNOPSIS
  Import TFS Work Item Templates
  .DESCRIPTION
  When Importing Work Item Template files into some Team Projects, slight modifications need to be made to those files so that they will
  import correctly between versions of TFS

  Looking at the source of this function, you can Add Rules as required
  .EXAMPLE
  Update-TfsWorkItemTemplate "C:\TFS\Original_WITD_Files\" "C:\TFS\Modified_WITD_Files\"
  .PARAMETER sourcePath
  The folder in which to find WIT Template .xml files
  .PARAMETER targetPath
  The folder in which to place the modified WIT Template .xml files
#>

    [CmdLetBinding()]
    param(
        [parameter(Mandatory=$true, ValueFromPipeline=$false)]
        [string]$sourcePath,
        [parameter(Mandatory=$true, ValueFromPipeline=$false)]
        [string]$targetPath
    )

    begin{ }
    process {
        $files = Get-ChildItem -path $sourcePath
        foreach ($file in $files){
            $sourceName = $file.FullName

            $tDir = Get-Item -Path $targetPath
            $targetName = Join-Path -Path $tDir -ChildPath $file.Name
            $template = cat $sourceName| % {$_ -replace "Iteration ID", "IterationID"}

            # INSERT ADDITIONAL REPLACEMENTS HERE
            $template = $template | % { $_ -replace "External Link Count", "ExternalLinkCount" }
            $template = $template | % { $_ -replace "Hyperlink Count", "HyperLinkCount" }
            $template = $template | % { $_ -replace "Attached File Count", "AttachedFileCount" }
            $template = $template | % { $_ -replace "Related Link Count", "RelatedLinkCount" }
            # END ADDITIONAL REPLACEMENTS 

            $template | % {$_ -replace "Area ID", "AreaID" } > "$targetName" 
        }
    }
    end{}
}

function Get-TfsTeamProjectCollectionAnalysis() {

    [CmdLetBinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipeLine = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer,
        [parameter(Mandatory=$true, ValueFromPipeline=$false)]
        [alias("Folder")]
        [string]$analysisRoot
    )

    begin{ }

    process {
        Write-Verbose "$(Get-Date -Format g) - Beginning Analysis"
        $tpcIds = Get-TfsTeamProjectCollectionIds $configServer

        Write-Verbose "Saving a dictionary of all fields found on PTCs"
        [hashtable]$witFieldDictionary = @{}
        [int]$totalFieldsFound = 6
    
        Write-Verbose "Iterating through each TPC"
        foreach($tpcId in $tpcIds){
            Write-Verbose "Get TPC instance for $tpcId"
            $tpc = $configServer.GetTeamProjectCollection($tpcId)

            Write-Verbose "----------------- $($tpc.Name) ---------------------------"

            Write-Verbose "Get list of version control repos in $($tpc.Name)"
            $vcs = $tpc.GetService([Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer])
            $vspec = [Microsoft.TeamFoundation.VersionControl.Client.VersionSpec]::Latest
            $recursionTypeOne = [Microsoft.TeamFoundation.VersionControl.Client.RecursionType]::OneLevel
            $recursionTypeFull = [Microsoft.TeamFoundation.VersionControl.Client.RecursionType]::Full
            $deletedState = [Microsoft.TeamFoundation.VersionControl.Client.DeletedState]::NonDeleted
            $itemType = [Microsoft.TeamFoundation.VersionControl.Client.ItemType]::Any
            $allReposInTPC = $vcs.GetItems("`$/", $vspec, $recursionTypeOne, $deletedState, $itemType, $false).Items

            Write-Verbose "Get WorKItemStore for $($tpc.Name)"
            $wiService = New-Object "Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore" -ArgumentList $tpc
        
            Write-Verbose "Get list of TeamProjects from WorkItemStore for $($tpc.Name)"
            $tps = $wiService.Projects

            Write-Verbose "Iterating through the Team Projects"
            Write-Host "Iteration through the Team Projects"
            [int]$count = 0
            foreach ($tp in $tps)
            { 
                if ($count %= 2) { Write-Host "." -NoNewline }else{ Write-Host "." -NoNewline -ForegroundColor Green }
                $count++

                Write-Verbose "---------------- $($tp.Name) ----------------"
                Write-Verbose "Get most recent changeset check-in for this TP"
                $currentTP = $allReposInTPC | ? {$_.ServerItem.Substring(2) -eq $($tp.Name)}
                $newitemSpec = New-Object Microsoft.TeamFoundation.VersionControl.Client.ItemSpec -ArgumentList $currentTP.ServerItem, $recursionTypeFull
                $latestChange = $vcs.GetItems($newitemSpec, $vspec, $deletedState, $itemType, $getItemsOptions).Items | Sort-Object CheckinDate -Desc
                [string]$mostRecentCheckin = $latestChange[0].CheckinDate.ToShortDateString()
            
                Write-Verbose "Most recent checkin date is $mostRecentCheckin"

                Write-Verbose "Get most recent work item change"
                $wiql = "SELECT [System.Id], [Changed Date] FROM WorkItems WHERE [System.TeamProject] = '$($tp.name)' ORDER BY [Changed Date] DESC"
                $wiQuery = New-Object -TypeName "Microsoft.TeamFoundation.WorkItemTracking.Client.Query" -ArgumentList $wiService, $wiql
                $results = $wiQuery.RunQuery()
                [string]$mostRecentChangedWorkItem = ""
                if ($results.Count -gt 0) {
	                $mostRecentChangedWorkItem = $results[0].ChangedDate.ToShortDateString()
                }
                Write-Verbose "Most recent work item changed date is $mostRecentChangedWorkItem"

                Write-Verbose "Get a list of WIT in this TP"
                $witList = & $witadmin listwitd /collection:$($tpc.Uri.AbsoluteUri) /p:$($tp.Name) | Sort-Object;
                $stringifiedWITList = [string]::Join("", $witList) -replace " ", ""
                $fieldsFingerprint = Get-Hash $stringifiedWITList
                $witList | % { if ($witFieldDictionary.ContainsKey($_) -eq $false) { $witFieldDictionary.Add($_, $totalFieldsFound); $totalFieldsFound += 3; } } | Out-Null
            
                Write-Verbose "Hash value of all field names is $fieldsFingerprint . (Process Template Id)"
            
                $csvLine = New-Object object[] 100;
                $csvLine[0] = $($tpc.Name); $csvLine[1] = $($tp.Name); $csvLine[2] = $fieldsFingerprint;
                $csvLine[3] = ""; $csvLine[4] = $mostRecentCheckin; $csvLine[5] = $mostRecentChangedWorkItem;
            
                Write-Verbose "Creating folder for Team Project artifacts"
                Write-Verbose $analysisRoot
                $exportFolder = New-Folder $analysisRoot
                $exportFolder = $exportFolder.CreateSubdirectory("WIT_EXPORTS\$($tp.Name)\")
                Write-Verbose $exportFolder

                $allWITHash = ""
                Write-Verbose "Iterating through the WIT"
                for ($i=0; $i -lt $witList.Length; $i++)
                {
                    $wit = $witList[$i]; 
                    $witStartCol = $witFieldDictionary[$wit]
                    $csvLine[$witStartCol] = $wit
                
                    Write-Verbose "Getting total number of work items for this WIT"
                    $wiql = "Select [System.Id] from WorkItems WHERE [System.WorkItemType] = '$wit' AND [System.TeamProject] = '$($tp.Name)'"
                    $wiQuery = New-Object -TypeName "Microsoft.TeamFoundation.WorkItemTracking.Client.Query" -ArgumentList $wiService, $wiql
                    $csvLine[$witStartCol+1] = $wiQuery.RunQuery().Count
                
                    Write-Verbose "Getting WITD xml"
                    [xml]$wit_definition_xml = & $witadmin exportwitd /collection:$($tpc.Uri.AbsoluteUri) /p:$($tp.Name) /n:$wit;

                    Write-Verbose "Cleaning up WITD xml"
                    Write-Verbose "Removing some nodes that cause fingerprinting problems but not a meaningful difference"
                    Remove-Nodes $wit_definition_xml "//SUGGESTEDVALUES"
                    Remove-Nodes $wit_definition_xml "//VALIDUSER"
                    Remove-Nodes $wit_definition_xml "//comment()"

                    Write-Verbose "Sorting all of the fields by name"
                    $fields = $wit_definition_xml.WITD.WORKITEMTYPE.FIELDS
                    $sortedFields = $fields.FIELD | Sort Name

                    Write-Verbose "Sorting all of the fields children alphabetically"
                    foreach($field in $sortedFields){
                        Switch-ChildNodes $field
                    }

                    Write-Verbose "Converting all field elements to non-self closing elemenets"
                    $sortedFields | % {$_.AppendChild($_.OwnerDocument.CreateTextNode("")) }| Out-Null

                    Write-Verbose "Replacing original field nodelist with sorted field node list" 
                    [void]$fields.RemoveAll()
                    $sortedFields | foreach { $fields.AppendChild($_) } | Out-Null

                    Write-Verbose "Save a copy of the un-compressed wit for manual comparisons"
                    $fileLocation = $exportFolder.FullName + $wit + "-" + "$($tp.Name).xml"
                    Write-Verbose "File location: $fileLocation"
                    [void]$wit_definition_xml.Save($fileLocation)
                
                    # http://blogs.technet.com/b/heyscriptingguy/archive/2011/03/21/use-powershell-to-replace-text-in-strings.aspx
                    $compressed = $wit_definition_xml.InnerXML -replace " ", ""
                    $output = Get-Hash $compressed
                    $allWITHash += $output
                    $csvLine[$witStartCol+2]= $output
                }
           
                $csvLine[3] =  Get-Hash $allWITHash
                $csvLine = [string]::Join(",", $csvLine)
                Write-Debug $csvLine
                Write-Verbose "Writing csv entry to $($analysisRoot + "\witd.csv")"
                $csvLine >> $($analysisRoot + "\witd.csv")
            }
            Write-Host "Completed Iterating Through Team Projects"
        }

        Write-Verbose "$(Get-Date -Format g) - Analysis Complete"
    }
}

function Save-TfsCleanedWITD(){
    [CmdLetBinding()]
    param(
        [parameter(Mandatory = $true)]
        [xml]$witd,
        [parameter(Mandatory=$true)]
        [alias("Folder")]
        [string]$exportRoot, 
        [parameter(Mandatory=$false)]
        [Alias("Prefix")]
        [string]$pre
    )

    begin{ }
    process {
        $exportFolder = New-Folder $exportRoot

        Write-Verbose "Cleaning up WITD xml"
        Write-Verbose "Removing some nodes that cause fingerprinting problems but not a meaningful difference"
        Remove-Nodes $witd "//SUGGESTEDVALUES"
        Remove-Nodes $witd "//VALIDUSER"
        Remove-Nodes $witd "//comment()"

        Write-Verbose "Sorting all of the fields by name"
        $fields = $witd.WITD.WORKITEMTYPE.FIELDS
        $sortedFields = $fields.FIELD | Sort Name

        Write-Verbose "Sorting all of the fields children alphabetically"
        foreach($field in $sortedFields){
        Switch-ChildNodes $field
        }

        Write-Verbose "Converting all field elements to non-self closing elemenets"
        $sortedFields | % {$_.AppendChild($_.OwnerDocument.CreateTextNode("")) }| Out-Null

        Write-Verbose "Replacing original field nodelist with sorted field node list" 
        [void]$fields.RemoveAll()
        $sortedFields | foreach { $fields.AppendChild($_) } | Out-Null

        Write-Verbose "Save a copy of the un-compressed wit for manual comparisons"
        $fileLocation = $exportRoot + "$($pre + $witd.WITD.WORKITEMTYPE.name).xml"
        Write-Verbose "File location: $fileLocation"
        [void]$witd.Save($fileLocation)
    }
}
function Get-TfsXAMLBuildsCreatingWorkItems(){
    [CmdLetBinding(SupportsShouldProcess=$true)]
    param(
        [parameter(Mandatory = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer,
        [parameter(Mandatory = $false)]
        [string]$tpcName,
        [parameter(Mandatory = $false)]
        [string]$tpName,
        [parameter(Mandatory = $false)]
        [string]$buildName
    )
   begin{}
   process{
        $tpcIds = Get-TfsTeamProjectCollectionIds $configServer

        [array]$bdefs = ,
        [array]$actions = @(, @())

        foreach($tpcId in $tpcIds)
        {
            #Get TPC instance
            $tpc = $configServer.GetTeamProjectCollection($tpcId)
            if (![string]::IsNullorWhiteSpace($tpcName) -and ($tpcName -ne $tpc.Name) ) { continue; }

            $bs = $tpc.GetService([Microsoft.TeamFoundation.Build.Client.IBuildServer])
 
            #Get WorkItemStore
            $wiService = $tpc.GetService([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])
            #Get a list of TeamProjects
            $tps = $wiService.Projects

            #iterate through the TeamProjects
            foreach ($tp in $tps)
            { 
                if( ![string]::IsNullorWhiteSpace($tpName) -and ($tpName -ne $tp.Name) ) { continue;}

                $buildDefinitionList = $bs.QueryBuildDefinitions($tp.Name) 
                               
                foreach ($bdef in $buildDefinitionList)
                {
                   if( ![string]::IsNullorWhiteSpace($buildName) -and ($buildName -ne $bdef.Name) ) { continue;}

                    $buildPT = $bdef.Process
                    [xml]$buildPTXml = $buildPT.Parameters
                    if ($buildPTXml.Activity.'Process.CreateWorkItem' -ne $null){
                        $bdefs += $bdef
                    }
                }
            }
        }
        Write-Output $bdefs
   }
   end{}
}

function Update-TfsXAMLBuildDefintionDropFolder(){
    [CmdLetBinding(SupportsShouldProcess=$true)]
    param(
        [parameter(Mandatory = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer,
        [parameter(Mandatory = $true)]
        [string]$newFolder,
        [parameter(Mandatory = $false)]
        [string]$tpcName,
        [parameter(Mandatory = $false)]
        [string]$tpName,
        [parameter(Mandatory = $false)]
        [string]$buildName,
        [parameter(Mandatory = $false)]
        [boolean]$useFullName
    )
   begin{
        if (!(Test-Path -Path $folderPath)){
            Write-Error "The folder path requested does not exist"
        } else {
            $path = Get-Item -Path $folderPath 
            $newFolder = $path.PSPath
        }
   }
   process{
        $tpcIds = Get-TfsTeamProjectCollectionIds $configServer
        foreach($tpcId in $tpcIds)
        {
            #Get TPC instance
            $tpc = $configServer.GetTeamProjectCollection($tpcId)
            if (![string]::IsNullorWhiteSpace($tpcName) -and ($tpcName -ne $tpc.Name) ) { continue; }

            $bs = $tpc.GetService([Microsoft.TeamFoundation.Build.Client.IBuildServer])
 
            #Get WorkItemStore
            $wiService = $tpc.GetService([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])
            #Get a list of TeamProjects
            $tps = $wiService.Projects

            #iterate through the TeamProjects
            foreach ($tp in $tps)
            { 
                if( ![string]::IsNullorWhiteSpace($tpName) -and ($tpName -ne $tp.Name) ) { continue;}

                $buildDefinitionList = $bs.QueryBuildDefinitions($tp.Name) 
                               
                [array]$bdefs = ,
                [array]$actions = @(, @())
                foreach ($bdef in $buildDefinitionList)
                {
                   if( ![string]::IsNullorWhiteSpace($buildName) -and ($buildName -ne $bdef.Name) ) { continue;}

                    $messages = @()
                    $messages += "The default drop location for build definition "
                    $messages += "$($bdef.Name) "
                    $messages += " in "
                    $messages += "$($bdef.TeamProject)"
                    $messages += " would have been changed from "
                    $messages += "$($bdef.DefaultDropLocation)"
                    $messages += " to " 
                        
                    # update drop location
                    if ($useFullName){
                        $bdef.DefaultDropLocation = $newFolder + $tp.Name 
                    }else {
                        $bdef.DefaultDropLocation = $newFolder + $tp.Id
                    }

                    $messages += "$($bdef.DefaultDropLocation)."
                    $actions += ,$messages
                    
                    $bdefs += $bdef
                }
                if (!$WhatIfPreference) {
                    $bs.SaveBuildDefinitions($bdefs)
                } else {
                    if ($actions.Length -le 0) {
                        Write-Host "No changes were potentially made to any build definitions in $($tp.Name) - $($tpc.Name)"
                    }else { 
                        $actions | % { Write-Host $_[0] -NoNewline
                                       Write-Host $_[1] -ForegroundColor Yellow -NoNewline
                                       Write-Host $_[2] -NoNewline
                                       Write-Host $_[3] -ForegroundColor Magenta -NoNewline
                                       Write-Host $_[4] -NoNewline
                                       Write-Host $_[5] -ForegroundColor Red -NoNewline
                                       Write-Host $_[6] -NoNewline
                                       Write-Host $_[7] -ForegroundColor Green
                                     } 
                    }
                }
            }
        }
   }
   end{}
}

function Update-TfsXAMLBuildDefintionCurrentController(){
    [CmdLetBinding(SupportsShouldProcess=$true)]
    param(
        [parameter(Mandatory = $true)]
        [Microsoft.TeamFoundation.Client.TfsConfigurationServer]$configServer,
        [parameter(Mandatory = $true)]
        [string]$newController,
        [parameter(Mandatory = $false)]
        [string]$tpcName,
        [parameter(Mandatory = $false)]
        [string]$tpName,
        [parameter(Mandatory = $false)]
        [string]$buildName
    )
   begin{}
   process{
        $tpcIds = Get-TfsTeamProjectCollectionIds $configServer
        foreach($tpcId in $tpcIds)
        {
            #Get TPC instance
            $tpc = $configServer.GetTeamProjectCollection($tpcId)
            if (![string]::IsNullorWhiteSpace($tpcName) -and ($tpcName -ne $tpc.Name) ) { continue; }

            $bs = $tpc.GetService([Microsoft.TeamFoundation.Build.Client.IBuildServer])
 
            #Get WorkItemStore
            $wiService = $tpc.GetService([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])
            #Get a list of TeamProjects
            $tps = $wiService.Projects

            #iterate through the TeamProjects
            foreach ($tp in $tps)
            { 
                if( ![string]::IsNullorWhiteSpace($tpName) -and ($tpName -ne $tp.Name) ) { continue;}

                $buildDefinitionList = $bs.QueryBuildDefinitions($tp.Name) 
                               
                [array]$bdefs = ,
                [array]$actions = @(, @())
                $controllerName = $bs.QueryBuildControllers() | ? {$_.Name.ToLower().Contains($newController.ToLower())} | % {$_.Name}
                try{
                    $controller = $bs.GetBuildController($controllerName) 
                }
                catch [Exception]{
                    Write-Host "There is no build controller named " -ForegroundColor Red -NoNewline
                    Write-Host $newController -ForegroundColor Yellow -NoNewline
                    Write-Host " associated with " -NoNewLine
                    Write-Host $($tpc.Name) -ForegroundColor Magenta
                    continue
                }

                $buildDefinitionList | % {
                    Write-Verbose "Checking $($_.Name)"
                    Write-Verbose "$($_.Name) is using $($_.BuildController.Name)"
                    if ($_.BuildController.Uri -ne $controller.Uri) {
                            $messages = @()
                            $messages += "Would have set "
                            $messages += "$($_.Name)" 
                            $messages += " to use " 
                            $messages += "$($controllerName)."
                            $actions += ,$messages
                            $_.BuildController = $controller
                    }
                    else {
                        Write-Verbose "Build controller is already set. Taking no action."
                    }
                
                    if (!$WhatIfPreference) {
                        # update controller
                        Write-Host "Stop here"
                        #$_.Save()
                    } else {
                        if ($actions.Length -le 0) { 
                            Write-Host "No changes were potentially made to any build definitions in " -NoNewline
                            Write-Host  $($tp.Name) -ForegroundColor Yellow -NoNewline
                            Write-host " - $($tpc.Name)" -ForegroundColor Magenta
                        }else { 
                            $actions | % { Write-Host $_[0] -NoNewline
                                           Write-Host $_[1] -ForegroundColor Yellow -NoNewline
                                           Write-Host $_[2] -NoNewline
                                           Write-Host $_[3] -ForegroundColor Magenta -NoNewline 
                                         } 

                        }
                    }
                }
            }
        }
   }
   end{}
}

Set-Alias gh Get-Hash
Set-Alias gtfs Get-TfsConfigServer

Export-ModuleMember -Alias *
Export-ModuleMember -Function "New-Folder", "Get-Hash", "Switch-ChildNodes", "Remove-Nodes"
Export-ModuleMember -Function "Get-Nuget", "Get-TfsAssembliesFromNuget"
Export-ModuleMember -Function "Get-TfsConfigServer", "Get-TfsTeamProjectCollectionIds"
Export-ModuleMember -Function "Get-TfsEventSubscriptions"
Export-ModuleMember -Function "Update-TfsXAMLBuildPlatformConfiguration", "Update-TfsXAMLBuildDefintionCurrentController", "Update-TfsXAMLBuildDefintionDropFolder", "Get-TfsXAMLBuildsCreatingWorkItems"
Export-ModuleMember -Function "Backup-TfsWorkItems", "Remove-TfsWorkItems", "Remove-TfsWorkItemTemplate", "Import-TfsWorkItemTemplate", "Update-TfsWorkItemTemplate", "Save-TfsCleanedWITD"
Export-ModuleMember -Function "Get-TfsTeamProjectCollectionAnalysis"
Export-ModuleMember -Function "Update-TfsFieldNames", "Find-TfsFieldDescription"

