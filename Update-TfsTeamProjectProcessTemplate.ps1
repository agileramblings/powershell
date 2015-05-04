#Master TeamProject Upgrade Script
Clear-Host

#setup global variaables
Write-Host "Building Global Variables list" -ForegroundColor Magenta
$myVar = @{ 
    "TFSUrl" = "http://divcd83:8080/tfs/";
    "TPC01" = "ProjectCollection01";  
    "TPC02" = "ProjectCollection02";  
    "TPScrum" = "Scrum 2013.4";
    "ResultsDir" = "C:\TFS\Results\";
    "DefaultDir" = "Default_2013_Scrum\";
    "ImportableDir" = "Importable_2013_Scrum\";
    "TFSVersion" = "2013.4";
}
$myVar.Values | % { Write-Host $_ -ForegroundColor Magenta }

Write-Host "List of TP that require manual upgrade " -ForegroundColor Cyan
$TPsToUpgrade = @(
     "(APM_CAR) APM - Calgary Awards Rewrite"
    ,"(APM_GCR) APM - General Current Reconcile"
    ,"(APM_ITSC) APM - IT Software Contracts"
    ,"(ARB) Assessment Review Board"
    ,"(BIS) Business Information System"
    ,"(CEM) Corporate Event Management"
    ,"(CIAO) Calgary Integrated Assessment Office"
    ,"(CTPI) Calgary Transit Property Information"
    ,"(DRMCC) Document Records Management Competency Centre"
    ,"(EAF) Engagement Assessment Form on the Web"
    ,"(ELAP) Environmental Liabilities Assessment Program"
    ,"(ESD) Enterprise Service Desk"
    ,"(GISGeorge) GISGeorge"
    ,"(IdM) Identity Management"
    ,"(LLAMA) Lease License Asset Mgmt App"
    ,"(STS) Security Token Service"
    ,"Calgarys Environmental Footprint Action Inventory"
    ,"CIAO Reporting"
    ,"Reuse Library"
    ,"Roads Asset Management"
)
$TPsToUpgrade | % { Write-Host $_ -ForegroundColor Cyan }

#export work items from Scrum
Write-Host "Getting WITD List for Scrum 2013.4 Process Template" -ForegroundColor Yellow
$canonicalScrumWITD = witadmin listwitd /collection:$($myVar.TFSUrl + $myVar.TPC02) /p:$($myVar.TPScrum)
$canonicalScrumWITD | % { witadmin exportwitd /collection:$($myVar.TFSUrl + $myVar.TPC02) /p:$($myVar.TPScrum) /n:"$($_)" /f:"$($myVar.ResultsDir + $myVar.DefaultDir)$_.xml" }

#transform Scrum WITD so they will import
Write-Host "Transforming WITD so they will Import" -ForegroundColor Yellow
Update-TfsWorkItemTemplate $($myVar.ResultsDir + $myVar.DefaultDir) $($myVar.ResultsDir + $myVar.ImportableDir)

#export categories for scrum
Write-Host "Getting Categories from Scrum 2013.4 Process Template" -ForegroundColor Yellow
witadmin exportcategories /collection:$($myVar.TFSUrl + $myVar.TPC02) /p:$($myVar.TPScrum) /f:"$($myVar.ResultsDir + $myVar.ImportableDir)categories.xml"

#create empty categories file
Write-Host "Creating Emtpy Categories files..." -ForegroundColor Yellow
[xml]$exportedCats = Get-Content -Path "$($myVar.ResultsDir + $myVar.ImportableDir)categories.xml"
$exportedCats.CATEGORIES.RemoveAll()
$exportedCats.Save("$($myVar.ResultsDir + $myVar.ImportableDir)Empty_categories.xml")

foreach ($tp in $TPsToUpgrade){
    Write-Host "Deleting all Work Items from $tp" -ForegroundColor Red
    #Delete All Work Items from Target TP
    Delete-TfsWorkItems $myVar.TFSUrl $myVar.TFSVersion $myVar.TPC01 $tp "Bug"

    #Change Categories to be empty
    Write-Host "Changing all Categories to be empty in $tp" -ForegroundColor Red
    witadmin importcategories /collection:$($myVar.TFSUrl + $myVar.TPC01) /p:$tp /f:"$($myVar.ResultsDir + $myVar.ImportableDir)Empty_categories.xml"

    #Delete WITD
    Write-Host "Destroying all WITD in $tp" -ForegroundColor Red
    Destroy-TfsWorkItemTemplate $myVar.TFSUrl $myVar.TFSVersion $myVar.TPC01 $tp "Bug"

    #Import new Scrum WITD
    Write-Host "Importing new Scrum 2013.4 WITD into $tp" -ForegroundColor Yellow
    Import-TfsWorkItemTemplate $($myVar.TFSUrl + $myVar.TPC01) $tp $($myVar.ResultsDir + $myVar.ImportableDir)
    
    #import new categories for Scrum
    Write-Host "Importing Scrum 2013.4 Categories into $tp" -ForegroundColor Yellow
    witadmin importcategories /collection:$($myVar.TFSUrl + $myVar.TPC01) /p:$($myVar.TP01) /f:"$($myVar.ResultsDir + $myVar.ImportableDir) + categories.xml"

    Write-Host "$tp should now be ready for an upgrade" -ForegroundColor Green
}

#upgrade TP


