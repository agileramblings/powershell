#Master TeamProject Upgrade Script
Clear-Host

ipmo DscTfs
#Get-Nuget
#Get-TfsAssembliesFromNuget

$witadmin = "C:\program files (x86)\Microsoft Visual Studio 14.0\common7\ide\witadmin.exe"
#$symbolicLocation = 'C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\'

###### Clean up all the work items that are noisey before we do anything
$ScriptPath = Split-Path $MyInvocation.InvocationName
#& "$ScriptPath\Clean-TfsWorkItems.ps1"

###### Setup global variaables
Write-Host "Building Global Variables list" -ForegroundColor Magenta
$myVar = @{ 
    "TFSUrl" = "http://ditfssb01:8080/tfs/";
    "TPC01" = "ProjectCollection01";  
    "TPC02" = "ProjectCollection02";  
    "ResultsDir" = "C:\TFS\Upgrade-Results";
    "DefaultDir" = "WITD_Templates";
    "ProcessTemplateRoot" = "\\cocdata1\dwhite2$\TFS\ProcessTemplates 2015";
    "ScrumDir" = "Scrum\WorkItem Tracking\TypeDefinitions";
    "AgileDir" = "Agile\WorkItem Tracking\TypeDefinitions";
    "CMMIDir" = "CMMI\WorkItem Tracking\TypeDefinitions";
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

$AGILE_TPs_To_Upgrade_TPC02 = @(
"(ADWAT) AD Web Application Template"
,"(AS5) Application Support - Team 5"
,"(DAMS) Development Approval Management System"
,"(EAM) EA Modeling"
,"(HPWS) Hansen Permits Web Service"
,"(IdM) Identity Management"
,"(ITDN) IT Dynamics"
,"(PITS) Plan Input Tracking System"
,"(QIReviewer) Quality Improvement Data Reporting"
,"(ROPP) Roads Online Permits Application"
,"(RTCM) Refined Transparent Cost Model"
,"Census Online"
,"CPA Park Online"
,"ERequests"
,"GIROACCES Conversion System"
,"WebApi"
)

$SCRUM_TPs_To_Upgrade_TPC02 = @(
"(ALPO) Animal Licence Payment Online"
,"(APGS) Advanced Print GIS SOE"
,"(APIS) Advanced Passenger Information Systems"
,"(APM_CC) APM Commuter Challenge"
,"(ArcGIS) Server Services"
,"(ArcGIS) Utilities"
,"(ATIS) Advanced Travellers Information System"
,"(AVL ) Approved Vendor List"
,"(BPM) BPM Common"
,"(BR) Building Repository"
,"(BSS) Business Support Services"
,"(CA GIS) CALGARY.CA GIS"
,"(CAAF) City Application Architecture Framework"
,"(CC) Commuter Challenge"
,"(CCMS) Catastrophe Claims Management System"
,"(CDAPOST) Canada Post Addressing"
,"(CDGIS) CityData GIS"
,"(CEM) Corporate Event Management"
,"(CEmobile) Census and Enumeration"
,"(CEMS) Crisis Emergency Management System"
,"(CENSUS GIS) CENSUS ANNNUAL MAPPING GIS"
,"(CENTS) Corporate Encroachment Tracking System"
,"(CFOS) Common Fleet Operation System"
,"(CIAOTBAM) CIAO Toolbar for ArcMap"
,"(CLIIP) Corporate Level Infrastructure Investment Planning"
,"(CLocker) CLocker"
,"(CM) Claims Mapping"
,"(CMS) Cemetery Management System"
,"(CoCISM) City of Calgary Internet Site – Mapping"
,"(COL) City Online Rewrite"
,"(COOL) Calgary Ownership OnLine"
,"(COP) Common Operating Picture"
,"(COPAP) COP Agreement Page"
,"(CPAMS) Civic Partner Management System"
,"(CPIP) Capital Planning Implementation Program"
,"(CR) Create Analysis"
,"(CSI) Customer Service Inquiries"
,"(CTEA) Calgary Transit Email Alerts"
,"(CTEFC) Calgary Transit Electronic Fare Collection"
,"(CTRPR) Calgary Transit Reserved Park and Ride"
,"(CTW) Calgary Transit Web Services"
,"(CWS) Corporate Web Services"
,"(DNRA) Donation Receipts Application"
,"(DSC) DevSupport Centre"
,"(DTR) Developers Test Repository"
,"(EAR) EA Repository"
,"(ELAP GIS) Environmental Liabilities Assessment Program GIS"
,"(FIMA) Fire Incidents Mobile Application"
,"(FIMS) FCSS Information Management System"
,"(FIS_IMP) Fire Information System Importer"
,"(GEM) Geospatial Emergency Management"
,"(GISDB) GIS Centre Database"
,"(HPI) Historical Photo Index"
,"(ICollS) Integrated Collection System"
,"(IMS) Issue Management System"
,"(INFRANET) Customized Microstation Drawing Tool"
,"(IPRM) Interactive Parks and Roads Map"
,"(IRAMS GIS) SIMS IRAMS Exporter GIS"
,"(ITCC) IT Cryptoconn"
,"(LTIAR) Land Titles Application Rewrite"
,"(LUAM) Land Use Amendment Maps"
,"(LUCC) Land Use Common Code"
,"(MACE) Management and Administration of Census Elections"
,"(MOST) Mobile Survey Tool"
,"(MXD GIS) MXD Converter GIS"
,"(OALG) Owner Address List Generation"
,"(OCS) Online Customer Service"
,"(OERS) Online Event Registration"
,"(PARIS) PARIS ParcMap"
,"(PARM) Police Accident Reconstruction Map"
,"(PDRO Election Reporting) PDRO Election Reporting"
,"(PDS) Planning and Development Support"
,"(PIP) Project Intake Process"
,"(PMRGIS) PM Routes GIS"
,"(PTWeb) Property Tax on Web"
,"(PUMA) Parcel Update Mapping Application Upgrade"
,"(RAMP) Roads Asset Management"
,"(RBLVD) Roads Boulevard"
,"(RC2I) Resource Centre 2 Integration"
,"(ReCaPT) Recreation Capital Planning Tool"
,"(RMBI) RiskMaster Batch Interfaces"
,"(RN2Ramp) Get RoadNet Data Batch"
,"(RWS) RemedyWS"
,"(SAMS) Subsidy Assistance Management System"
,"(SDAB) Subdivision and Development Appeal Board"
,"(SIMS) Site Information Management System"
,"(SLGIS) Street Light GIS"
,"(SPVWS) Survey Plan Validation Web Service"
,"(SSM) Software Solutions Methodology Site"
,"(TCFBIS) Transit Cash Fare Box Inventory System"
,"(TrafCntGIS) Traffic Count GIS"
,"(TRGIS) Transit Routes GIS"
,"(UDM) Utility Display Map for CFD"
,"(WARDGIS) Ward Redistricting GIS"
,"(WFAS) WaterFront for ArcGIS Server 10"
,"(WRSM) Waste and Recycling Service Management"
,"Access Calgary Reports"
,"AHS ArcReader Mobile Offline Map"
,"AS2 Work Load"
,"CAAF 3 Sample Application"
,"Census BlockFace"
,"CFOS Blackline Proof of Concept"
,"Chameleon"
,"CitizenPortal"
,"DocumentRouter"
,"DSC Test"
,"DSC-Test"
,"EAREP (EA Repository)"
,"ECustomerAgents"
,"ePayment"
,"eServices"
,"eServicesAdmin"
,"eServicesInspections"
,"eTime"
,"IRIS (Investment Recovery Information System)"
,"MoodleTest"
,"PD Bridge"
,"PD Web"
,"PDADEV"
,"ProxyHandler"
,"Roads Hansen Upgrade"
,"SWEEPSGIS"
,"Test"
,"Test-Project-From-VS13"
,"Water Center Information Touchscreen"
)

###### Get ConfigServer (authenticate) and setup TPC urls
$configServer = gtfs $myVar.TfsUrl
$tpc01Url = $($myVar.TFSUrl + $myVar.TPC01)
$tpc02Url = $($myVar.TFSUrl + $myVar.TPC02)

###### CMMI Team Project Upgrade in TPC01
#Update-TfsFieldNames $configServer $myVar.TPC01

###### create Folder Structure
$folder = New-Folder $myVar.ResultsDir
$defaultDir = Join-Path -Path $folder -ChildPath $myVar.DefaultDir | New-Folder 

###### Recreate child folders
$scrumDir = Join-Path -Path $myVar.ProcessTemplateRoot -ChildPath $myVar.ScrumDir | New-Folder 
$agileDir = Join-Path -Path $myVar.ProcessTemplateRoot -ChildPath $myVar.AgileDir | New-Folder 
$CMMIDir  = Join-Path -Path $myVar.ProcessTemplateRoot -ChildPath $myVar.CMMIDir  | New-Folder 

#foreach ($tp in $CMMI_TPs_To_Upgrade)
#{
#   #import in tweaked requirement WITD
#   Write-Host "Importing new CMMI (2015) WITD into $tp" -ForegroundColor Yellow
#   Import-TfsWorkItemTemplate $configServer $myVar.TPC01 $tp $CMMIDir
#   & $witadmin importcategories /collection:$tpc01Url /p:$($tp) /f:"$($CMMIDir.Parent.FullName.ToString() + "\categories.xml")"
#}

###### Team Project Upgrades in TPC02
#foreach ($tp in $CMMI_TPs_To_Upgrade_TPC02)
#{
#   #import in tweaked requirement WITD
#   Write-Host "Importing new CMMI (2015) WITD into $tp" -ForegroundColor Yellow
#   Import-TfsWorkItemTemplate $configServer $myVar.TPC02 $tp $CMMIDir
#   & $witadmin importcategories /collection:$tpc02Url /p:$($tp) /f:"$($CMMIDir.Parent.FullName.ToString() + "\categories.xml")"
#}
#foreach ($tp in $AGILE_TPs_To_Upgrade_TPC02)
#{
#   #import in tweaked requirement WITD
#   Write-Host "Importing new Agile (2015) WITD into $tp" -ForegroundColor Yellow
#   Import-TfsWorkItemTemplate $configServer $myVar.TPC02 $tp $agileDir
#   & $witadmin importcategories /collection:$tpc02Url /p:$($tp) /f:"$($agileDir.Parent.FullName.ToString() + "\categories.xml")"
#}
foreach ($tp in $SCRUM_TPs_To_Upgrade_TPC02)
{
   #import in tweaked requirement WITD
   Write-Host "Importing new Scrum (2015) WITD into $tp" -ForegroundColor Yellow
   Import-TfsWorkItemTemplate $configServer $myVar.TPC02 $tp $scrumDir
   & $witadmin importcategories /collection:$tpc02Url /p:$($tp) /f:"$($scrumDir.Parent.FullName.ToString() + "\categories.xml")"
}



###### Old Scrum and Agile Upgrade/Conversion
$scorchAndReplaceTPList = $OldAgileTPsToUpgrade + $OldScrumTPsToUpgrade
$scorchAndReplaceTPList | % { Write-Host $_ -ForegroundColor Cyan }
foreach ($tp in $scorchAndReplaceTPList){

    #get work item types from TP
    Write-Host "Get all Work Item Templates from $tp" -ForegroundColor Yellow
    $currentTPWITList = & $witadmin listwitd /collection:$($myVar.TFSUrl + $myVar.TPC01) /p:$($tp)
    
    Write-host "Backing up work items from $tp" -ForegroundColor Green
    foreach($wit in $currentTPWITList){
        Backup-TfsWorkItems $configServer $tpc01Url $tp $wit $folder
    }

    Write-Host "Deleting all Work Items from $tp" -ForegroundColor Red
    #Delete All Work Items from Target TP
    foreach($wit in $currentTPWITList){
        Remove-TfsWorkItems $configServer $myVar.TPC01 $tp $wit
    }
    
    #Delete WITD from Target TP
    Write-Host "Destroying all WITD in $tp" -ForegroundColor Red
    & $witadmin importcategories /collection:$tpc01Url /p:$($tp) /f:"$($myVar.ProcessTemplateRoot + "\TotallyEmptyCat.txt")"

    foreach($wit in $currentTPWITList){
        Remove-TfsWorkItemTemplate $configServer $myVar.TPC01 $tp $wit
    }

    #Import new Scrum WITD
    Write-Host "Importing new Scrum 2013.4 WITD into $tp" -ForegroundColor Yellow
    Import-TfsWorkItemTemplate $configServer $myVar.TPC01 $tp $scrumDir
    
    #import new categories for Scrum
    Write-Host "Importing Scrum 2013.4 Categories into $tp" -ForegroundColor Yellow
    & $witadmin importcategories /collection:$tpc01Url /p:$($tp) /f:"$($scrumDir.Parent.FullName.ToString() + "\categories.xml")"

    Write-Host "$tp should now be ready for an feature configuration on 2013.4" -ForegroundColor Green
}