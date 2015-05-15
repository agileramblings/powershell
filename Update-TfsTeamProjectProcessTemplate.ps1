#Master TeamProject Upgrade Script
Clear-Host

Set-Tfs2013

#setup global variaables
Write-Host "Building Global Variables list" -ForegroundColor Magenta
$myVar = @{ 
    "TFSUrl" = "http://ditfssb01:8080/tfs/";
    "TPC01" = "ProjectCollection01";  
    "TPC02" = "ProjectCollection02";  
    "TPScrum" = "DocumentRouter";
    "ResultsDir" = "C:\TFS\Trial-Upgrade-Results\";
    "DefaultDir" = "Default_2013_Scrum\";
    "ImportableDir" = "Importable_2013_Scrum\";
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

# Won't be able to do this - leave in for instruction sake
# Invoke-Command -ComputerName divcd163 -ScriptBlock { TFSBackup -i "PISQLSK801\TFSPD" -l "\\exchange\exchange\TFS_Migration" }
# Invoke-Command -ComputerName divcd163 -ScriptBlock { TfsRestore -i "DIVCD163" -l "\\exchange\exchange\TFS_Migration" }

$scorchAndReplaceTPList = $OldAgileTPsToUpgrade + $OldScrumTPsToUpgrade
$scorchAndReplaceTPList | % { Write-Host $_ -ForegroundColor Cyan }

$folder = New-Folder $myVar.ResultsDir
$defaultDir = New-Folder $($myVar.ResultsDir + $myVar.DefaultDir)
$importableDir = New-Folder $($myVar.ResultsDir + $myVar.ImportableDir)

#delete WITD and categories files between runs
Remove-Item "$($myVar.ResultsDir + $myVar.DefaultDir)\*"
Remove-Item "$($myVar.ResultsDir + $myVar.ImportableDir)\*"

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

$configServer = gtfs $myVar.TfsUrl $myVar.TFSVersion
foreach ($tp in $scorchAndReplaceTPList){

    #get work item types from TP
    Write-Host "Get all Work Item Templates from $tp" -ForegroundColor Yellow
    $currentTPWITList = witadmin listwitd /collection:$($myVar.TFSUrl + $myVar.TPC01) /p:$($tp)
    
    Write-host "Backing up work items from $tp" -ForegroundColor Green
    Write-Host "Deleting all Work Items from $tp" -ForegroundColor Red
    #Delete All Work Items from Target TP
    foreach($wit in $currentTPWITList){
        Backup-TfsWorkItems $configServer $($myVar.TFSUrl + $myVar.TPC01) $tp $wit $($myVar.ResultsDir)
        Remove-TfsWorkItems $configServer $myVar.TPC01 $tp $wit
    }
    
    #Change Categories to be empty
    Write-Host "Changing all Categories to be empty in $tp" -ForegroundColor Red
    witadmin importcategories /collection:$($myVar.TFSUrl + $myVar.TPC01) /p:$tp /f:"$($myVar.ResultsDir + $myVar.ImportableDir)Empty_categories.xml"

    #Delete WITD
    Write-Host "Destroying all WITD in $tp" -ForegroundColor Red
    foreach($wit in $currentTPWITList){
        Remove-TfsWorkItemTemplate $configServer $myVar.TPC01 $tp $wit
    }

    #Import new Scrum WITD
    Write-Host "Importing new Scrum 2013.4 WITD into $tp" -ForegroundColor Yellow
    Import-TfsWorkItemTemplate $configServer $myVar.TPC01 $tp $($myVar.ResultsDir + $myVar.ImportableDir)
    
    #import new categories for Scrum
    Write-Host "Importing Scrum 2013.4 Categories into $tp" -ForegroundColor Yellow
    witadmin importcategories /collection:$($myVar.TFSUrl + $myVar.TPC01) /p:$($tp) /f:"$($myVar.ResultsDir + $myVar.ImportableDir)categories.xml"

    Write-Host "$tp should now be ready for an upgrade" -ForegroundColor Green
}

#get Requirement WITD
[xml]$orgReg = witadmin exportwitd /collection:"http://ditfssb01:8080/tfs/ProjectCollection01" /p:"(ABSS) Ambulance Billing Source System" /n:"Requirement" /f:"C:\Temp\Test\Requirement.xml"

#modify it slightly 
# <FIELD name="Stack Rank" refname="Microsoft.VSTS.Common.StackRank" type="Double" />
$newStackRank = $orgReq.CreateElement("FIELD")
$newStackRank.SetAttribute('name','Stack Rank')
$newStackRank.SetAttribute('refname','Microsoft.VSTS.Common.StackRank')
$newStackRank.SetAttribute('type','Double')
$orgReq.WITD.WORKITEMTYPE.FIELDS.AppendChild($newStackRank)

# <FIELD name="Size" refname="Microsoft.VSTS.Scheduling.Size" type="Integer" />
$newSize = $orgReq.CreateElement("FIELD")
$newSize.SetAttribute('name','Size')
$newSize.SetAttribute('refname','Microsoft.VSTS.Scheduling.Size')
$newSize.SetAttribute('type','Integer')
$orgReq.WITD.WORKITEMTYPE.FIELDS.AppendChild($newSize)
$orgReq.Save($($myVar.ResultsDir + "NewReq.xml"))

foreach ($tp in $CMMI_TPs_To_Upgrade)
{
   #import in tweaked requirement WITD
   witadmin importwitd /collection:$($myVar.TFSUrl + $myVar.TPC01) /p:$tp  /f:"$($myVar.ResultsDir + "NewReq.xml")"
}



#upgrade TP


