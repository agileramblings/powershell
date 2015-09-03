cls
ipmo DscTfs

function Find-Button($ieDoc, $btnText){
    $btns = $ieDoc.getElementsByTagName("button")
    foreach($innerBtn in $btns) 
    {
        if($innerBtn.parentElement.className -ne "ui-dialog-buttonset") {continue}
        $innerSpans = $innerBtn.getElementsByTagName("span")
        foreach($span in $innerSpans)
        {
            if (($span.InnerText) -and ($span.InnerText.Contains($btnText))) {
                #find the button that has a span that has the text btnText
                $span.parentElement
                break;
            }
        }
    }
}

$ie = new-object -ComObject "InternetExplorer.Application"

$cs = Get-TfsConfigServer http://ditfssb01:8080/tfs
$tpcIds = Get-TfsTeamProjectCollectionIds $cs

foreach ($tpcId in $tpcIds){
    
    $tpc = Get-TfsTeamProjectCollection $cs -teamProjectCollectionId $tpcId
    [string]$tpcUri =  $tpc.Uri.AbsoluteUri

    $projects = Get-TfsTeamProjects -configServer $cs -teamProjectCollectionId $tpcId
    foreach ($proj in $projects){
        [string]$projectName = $proj.Name
        #if (!$projectName.Contains("TTest")) { continue; } #filter if necessary
        $requestUri = [string]::Format("{0}/{1}/_admin#_a=enableFeatures", $tpcUri, $projectName.Replace(" ", "%20"))
        $verifyButtonText = "Verify"
        $configureButtonText = "Configure"
        $closeButtonText = "Close"

        $ie.visible = $true
        $ie.silent = $true
        $ie.navigate($requestUri)
        while($ie.Busy) { Start-Sleep -Milliseconds 100 }

        $doc = $ie.Document
        
        #discover Verification button 
        $btn = Find-Button $doc $verifyButtonText

        if ($btn -eq $null) { continue }
    
        #start Verification
        $btn.click()

        Start-Sleep -Milliseconds 1000

        $buttonNotFound = $true
        #wait for verification to complete
        while ($buttonNotFound) {
            $closeBtn = $null;$configBtn = $null;
            $closeBtn = Find-Button $doc $closeButtonText
            $configBtn = Find-Button $doc $configureButtonText
            if (($closeBtn -ne $null) -or ($configBtn -ne $null)){
                $buttonNotFound = $false;
            }else {
                Start-Sleep -Milliseconds 1000
            }
        }

        if ($closeBtn -ne $null) {
            Write-Host "Cannot configure features for TeamProject " -NoNewline
            Write-host "($($proj.Name)). " -NoNewLine -ForegroundColor Yellow
            Write-Host "It needs to be upgraded first."
            $warningText = $doc.getElementById("issues-textarea-id").InnerText
            Write-Host $warningText -ForegroundColor Red | fl -Force
            $closeBtn.click()
        }
        elseif ($configBtn -ne $null) {
            #start Configuration
            $configBtn.click()

            #wait for configuration to complete
            Start-Sleep -Milliseconds 500

            #close Configuration
            $buttonNotFound = $true
            while ($buttonNotFound) {
                $closeBtn = $null;
                $closeBtn = Find-Button $doc $closeButtonText
                if ($closeBtn -ne $null){
                    $buttonNotFound = $false;
                }else {
                    Start-Sleep -Milliseconds 500
                }
            }

            $closeBtn.click()
            
        }
        else{
            Write-Host "Failed to find a button"
        }
    }
}


