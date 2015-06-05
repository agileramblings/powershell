#Master TeamProject Upgrade Script
Clear-Host

Set-Tfs2013

###### Clean up all the work items that are noisey before we do anything
.\Clean-TfsWorkItems.ps1

###### Setup global variaables
Write-Host "Building Global Variables list" -ForegroundColor Magenta
$myVar = @{ 
    "TFSUrl" = "http://ditfssb01:8080/tfs/";
    "TPC01" = "ProjectCollection01";  
    "TPC02" = "ProjectCollection02";  
    "TPScrum" = "(SAMPLE) Scrum 2013.4";
    "TPAgile" = "(SAMPLE) Agile 2013.4";
    "TPCMMI"  = "(SAMPLE) CMMI 2013.4";
    "ResultsDir" = "C:\TFS\FullTrial-Upgrade-Results";
    "DefaultDir" = "WITD_Templates";
    "ScrumDir" = "Scrum";
    "AgileDir" = "Agile";
    "CMMIDir" = "CMMI";
    "TFSVersion" = "2013.4";
}
$myVar.GetEnumerator() | % { Write-Host $("{0}" -f $_.Key) -ForegroundColor Green -NoNewline; Write-Host $(" = {0}" -f $_.Value) -ForegroundColor Magenta }

Write-Host "List of TP that require manual upgrade " -ForegroundColor Cyan
$OldAgileTPsToUpgrade = @(
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
    ,"eServicesFeeEstimator"
)

$OldScrumTPsToUpgrade = @(
 "(CDE) Census Data Entry"
,"(CIAO-BA) Business Assessment into CIAO"
,"(DSCTEST) DSC Test Repository"
,"(LRT) LRT Relay"
,"(OCS) Online Customer Service"
,"(PTWeb) Property Tax on Web"
,"CAAF"
,"Test sfts v2"
)

$CMMI_TPs_To_Upgrade = @(
 "(ABSS) Ambulance Billing Source System"
,"(ACE) Access Calgary Extra"
,"(AIM_WTS) AIM Work Request Tracking System"
,"(ANP) Assessment Notice Print"
,"(APEB) AP Enmax Batch"
,"(APEFT) AP Electronic Funds Transfer"
,"(APESV) AP Enmax Site Verification"
,"(APM_APP_TEMPLATE) APM - Application Template"
,"(APM_MF) APM - Master Form Catalog"
,"(APM_SMS) APM - Secretariat Meeting Statistics"
,"(APM_TLP) APM - Transit Lost Property"
,"(ArcGIS) Server Services"
,"(ArcGIS) Utilities"
,"(ATPA) ARINC Train Performance Analysis"
,"(BCC) Boards Commissions Committees"
,"(BDS) Bylaw Debenture Integration"
,"(BISINFO) Business Inventory System Information"
,"(BOS) Bill of Sale"
,"(Calgro IS) Calgro Information System"
,"(CAM) Corporate Addressing and Mapping"
,"(CARS) Contracts and Agreements Recording System"
,"(CBET) Contextual Building Envelope Tool"
,"(CEFAI) Calgarys Environmental Footprint Action Inventory"
,"(CHALK) Chalkboard"
,"(CLINKS) Housekeeping Scheduling System"
,"(CLNDR) Calendar Web Service"
,"(CM) CoreMapping"
,"(CMA) Control Monuments Application"
,"(CMS) Cemetery Management System"
,"(CoCIS) City of Calgary Internet Site"
,"(COL) City Online Rewrite"
,"(COOL) Calgary Ownership OnLine"
,"(COS) Calgary Online Store"
,"(CPAL) Corp Prop Address Lookup"
,"(CPDMS) Capital Projects Data Maintenance System"
,"(CryptoConn) CryptoConn"
,"(CWAM) Corporate Web Authentication Model"
,"(DBASDM) DevBuildingApproval Spatial Data Management"
,"(DHC) Discover Historic Calgary"
,"(DMS) Debenture Management System"
,"(DSC) Development Support Centre"
,"(DSCommon) DevSupport Common"
,"(DSDataLayer) DevSupport Data Layer"
,"(DTR) Developers Test Repository"
,"(EAI) Enterprise Application Integration"
,"(EDIW) Enterprise Directory Internal Webservice"
,"(ETIME) Electronic Time Project"
,"(FAM) Fixing Addresses Manually"
,"(FCL) Finance Codes Lookup"
,"(FireRMS) Fire Record Management System"
,"(FLTR) Fleet Training"
,"(FMIS_MIIS) Employee Data Feed from MIIS to FMIS"
,"(FRED) Footprint Reporting Environmental Data"
,"(FSII) FCSS Social Inclusion Indicators"
,"(ICollS) Integrated Collection System"
,"(IdMEAS) IdM External Account Service"
,"(ITWP) webwave"
,"(LAT) Legacy Application Transfer"
,"(LI) Local Improvement Replacement Project"
,"(LINDA) Land Inventory Data Application)"
,"(LIPS) Low Income Pass System"
,"(LRVTB) LRV Tracking Board"
,"(LTIR) Land Title Information Reports"
,"(MAF) Mailing Address Formatter"
,"(MSPS) MS Project Server"
,"(MTS) Mail Tracking System"
,"(NESAA) IdM Non-Employee System of Record"
,"(PARIS) PARIS_Launcher"
,"(PARIS) PARIS_Reports"
,"(PMR) PM Routes"
,"(RACE) Reporting and Analysis"
,"(RB) Retiree Benefits"
,"(RD) Road Detours"
,"(SNAG) Subdivision Notice of Appeal Generator"
,"(STAR) Staff Directory Tabular Access and Standard Reports"
,"(SWB) Solid Waste Billing"
,"(TAP) Temporary Administration Privileges"
,"(TAR) IT Temporary Asset Register"
,"(TCM) Transparent Cost Model"
,"(TEMPS) Temp Employment Management"
,"(TG) Trip Generation"
,"(TP) Triangle Project"
,"(TRFCNT) Traffic Counts"
,"(UDRS) Urban Development Reporting System"
,"(ULA) Utility Line Assignments"
,"(VAMIS) Vendor Account Management Information System"
,"(WBOR) Water Billing Operational Repository"
,"(WRSM) Waste and Recycling Service Management"
,"(WSA) Windows Server Administration"
,"Eclipse Test Project01"
,"M5 Reporting"
,"Management Dashboard"
,"Management Dashboard Sustainment"
,"myCity"
,"Remedy 7"
)

$CMMI_TPs_To_Upgrade_TPC02 = @(
 "(FCSWS) Family Community Survey Web Service"
,"(GP) Glacier Project"
,"(POWS) Property Ownership Web Service"
,"(LABC) Labour Action Business Continuity"
)

###### Ge ConfigServer (authenticate) and setup TPC urls
$configServer = gtfs $myVar.TfsUrl $myVar.TFSVersion
$tpc01Url = $($myVar.TFSUrl + $myVar.TPC01)
$tpc02Url = $($myVar.TFSUrl + $myVar.TPC02)

###### CMMI Team Project Upgrade in TPC01
Update-TfsFieldNames $configServer $myVar.TPC01

## create the three new team projects
tfpt createteamproject /collection:$tpc01Url /teamproject:"(SAMPLE) CMMI 2013.4" /processtemplate:"MSF for CMMI Process Improvement 2013.4" /sourcecontrol:None /noreports /noportal
tfpt createteamproject /collection:$tpc01Url /teamproject:"(SAMPLE) Scrum 2013.4" /processtemplate:"Microsoft Visual Studio Scrum 2013.4" /sourcecontrol:None /noreports /noportal
tfpt createteamproject /collection:$tpc01Url /teamproject:"(SAMPLE) Agile 2013.4" /processtemplate:"MSF for Agile Software Development 2013.4" /sourcecontrol:None /noreports /noportal

###### Build Folder Structure
$folder = New-Folder $myVar.ResultsDir
$defaultDir = Join-Path -Path $folder -ChildPath $myVar.DefaultDir | New-Folder 

###### delete WITD and categories files between runs
Remove-Item "$defaultDir\*" -Recurse

###### Recreate child folders
$scrumDir = Join-Path -Path $defaultDir -ChildPath $myVar.ScrumDir | New-Folder 
$agileDir = Join-Path -Path $defaultDir -ChildPath $myVar.AgileDir | New-Folder 
$CMMIDir = Join-Path -Path $defaultDir -ChildPath $myVar.CMMIDir | New-Folder 

Write-Host "Getting WITD List for Scrum 2013.4 Process Template" -ForegroundColor Yellow
$defaultScrumWITD = witadmin listwitd /collection:$tpc01Url /p:$($myVar.TPScrum)
$defaultAgileWITD = witadmin listwitd /collection:$tpc01Url /p:$($myVar.TPAgile)
$defaultCMMIWITD = witadmin listwitd /collection:$tpc01Url /p:$($myVar.TPCMMI)
                                                 
###### Get WITD files from Templates
foreach($witdName in $defaultScrumWITD)
{ 
    witadmin exportwitd /collection:$tpc01Url /p:$($myVar.TPScrum) /n:"$($witdName)" /f:"$scrumDir\$witdName.xml" 
}
foreach($witdName in $defaultAgileWITD)
{ 
    witadmin exportwitd /collection:$tpc01Url /p:$($myVar.TPAgile) /n:"$($witdName)" /f:"$agileDir\$witdName.xml" 
}
foreach($witdName in $defaultCMMIWITD)
{ 
    witadmin exportwitd /collection:$tpc01Url /p:$($myVar.TPCMMI) /n:"$($witdName)" /f:"$CMMIDir\$witdName.xml" 
}

###### export categories for process templates
Write-Host "Getting Categories from 2013.4 Process Templates" -ForegroundColor Yellow
$scrumCat = Join-Path -Path $scrumDir -ChildPath categories.xml
witadmin exportcategories /collection:$tpc01Url /p:$($myVar.TPScrum) /f:"$scrumCat"
$agileCat = Join-Path -Path $agileDir -ChildPath categories.xml
witadmin exportcategories /collection:$tpc01Url /p:$($myVar.TPAgile) /f:"$agileCat"
$cmmiCat = Join-Path -Path $CMMIDir -ChildPath categories.xml
witadmin exportcategories /collection:$tpc01Url /p:$($myVar.TPCMMI) /f:"$cmmiCat"

foreach ($tp in $CMMI_TPs_To_Upgrade)
{
   #import in tweaked requirement WITD
   Write-Host "Importing new CMMI 2013.4 WITD into $tp" -ForegroundColor Yellow
   Import-TfsWorkItemTemplate $configServer $myVar.TPC01 $tp $CMMIDir
   witadmin importcategories /collection:$tpc01Url /p:$($tp) /f:"$cmmiCat"
}

###### CMMI Team Project Upgrade in TPC02
foreach ($tp in $CMMI_TPs_To_Upgrade_TPC02)
{
   #import in tweaked requirement WITD
   Write-Host "Importing new CMMI 2013.4 WITD into $tp" -ForegroundColor Yellow
   Import-TfsWorkItemTemplate $configServer $myVar.TPC02 $tp $CMMIDir
   witadmin importcategories /collection:$tpc02Url /p:$($tp) /f:"$cmmiCat"
}

###### Old Scrum and Agile Upgrade/Conversion
$scorchAndReplaceTPList = $OldAgileTPsToUpgrade + $OldScrumTPsToUpgrade
$scorchAndReplaceTPList | % { Write-Host $_ -ForegroundColor Cyan }
foreach ($tp in $scorchAndReplaceTPList){

    #get work item types from TP
    Write-Host "Get all Work Item Templates from $tp" -ForegroundColor Yellow
    $currentTPWITList = witadmin listwitd /collection:$($myVar.TFSUrl + $myVar.TPC01) /p:$($tp)
    
    Write-host "Backing up work items from $tp" -ForegroundColor Green
    Write-Host "Deleting all Work Items from $tp" -ForegroundColor Red
    #Delete All Work Items from Target TP
    foreach($wit in $currentTPWITList){
        Backup-TfsWorkItems $configServer $tpc01Url $tp $wit $folder
        Remove-TfsWorkItems $configServer $myVar.TPC01 $tp $wit
    }
    
    #Delete WITD
    Write-Host "Destroying all WITD in $tp" -ForegroundColor Red
    foreach($wit in $currentTPWITList){
        Remove-TfsWorkItemTemplate $configServer $myVar.TPC01 $tp $wit
    }

    #Import new Scrum WITD
    Write-Host "Importing new Scrum 2013.4 WITD into $tp" -ForegroundColor Yellow
    Import-TfsWorkItemTemplate $configServer $myVar.TPC01 $tp $scrumDir
    
    #import new categories for Scrum
    Write-Host "Importing Scrum 2013.4 Categories into $tp" -ForegroundColor Yellow
    witadmin importcategories /collection:$tpc01Url /p:$($tp) /f:"$scrumCat"

    Write-Host "$tp should now be ready for an feature configuration on 2013.4" -ForegroundColor Green
}
