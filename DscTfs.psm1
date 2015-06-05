#Get-TfsConfigServer "http://divcd83:8080/tfs" "2013.4" | Get-TfsTeamProjectCollectionAnalysis -Folder "C:\temp\Test_Analysis" -Verbose 4> "C:\Temp\Analysis_log.txt"

Write-Host "Loading DscTfs Module"

function Import-TFSAssemblies_2010 {
    Add-Type -AssemblyName "Microsoft.TeamFoundation.Client, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
    Add-Type -AssemblyName "Microsoft.TeamFoundation.Common, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
    Add-Type -AssemblyName "Microsoft.TeamFoundation.VersionControl.Client, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
    Add-Type -AssemblyName "Microsoft.TeamFoundation.WorkItemTracking.Client, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
}

function Import-TFSAssemblies_2013 {
    Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.Client.dll";
    Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.Common.dll";
    Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.VersionControl.Client.dll";
    Add-Type -LiteralPath "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0\Microsoft.TeamFoundation.WorkItemTracking.Client.dll";
}

[string]$targetVersion = "2013.4"
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
        [string]$url,
        [parameter(Mandatory = $true)]
        [string]$tfsVersion)

    begin {
        Write-Verbose "Loading TFS OM Assemblies for $tfsVersion"
        $targetVersion = $tfsVersion
        if ($tfsVersion.Contains("2010")){
            Import-TFSAssemblies_2010
        } elseif ($tfsVersion.Contains("2013")) {
            Import-TFSAssemblies_2013
        } else{
            Import-TFSAssemblies_2013
        }
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
                
            #Get WorkItemStore
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
                    witadmin destroywi /collection:$($tpc.Uri.AbsoluteUri) /id:$ids /noprompt | Out-Null
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
            $witList = witadmin listwitd /collection:$tpcUrl  /p:$tpName | Sort-Object;
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
                    witadmin destroywitd /collection:$($tpc.Uri.AbsoluteUri)  /p:$($tp.Name) /n:"$witType" /noprompt | Out-Null
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
            witadmin importwitd /collection:$tpcUrl /p:"$tpName" /f:"$fileName"
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
        $arrayEntries = witadmin listfields /collection:$tpcUrl
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
            witadmin changefield /collection:$tpcUrl /n:"$($field.Refname)" /name:"$($field.Name)" /noprompt 
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
                $witList = witadmin listwitd /collection:$($tpc.Uri.AbsoluteUri) /p:$($tp.Name) | Sort-Object;
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
                    [xml]$wit_definition_xml = witadmin exportwitd /collection:$($tpc.Uri.AbsoluteUri) /p:$($tp.Name) /n:$wit;

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

Set-Alias gh Get-Hashtfpt
Set-Alias gtfs Get-TfsConfigServer

Export-ModuleMember -Alias *
Export-ModuleMember -Function "New-Folder", "Get-Hash", "Switch-ChildNodes", "Remove-Nodes"
Export-ModuleMember -Function "Get-TfsConfigServer", "Get-TfsTeamProjectCollectionIds"
Export-ModuleMember -Function "Backup-TfsWorkItems", "Remove-TfsWorkItems", "Remove-TfsWorkItemTemplate", "Import-TfsWorkItemTemplate", "Update-TfsWorkItemTemplate"
Export-ModuleMember -Function "Get-TfsTeamProjectCollectionAnalysis", "Update-TfsFieldNames", "Find-TfsFieldDescription", "Save-TfsCleanedWITD"

