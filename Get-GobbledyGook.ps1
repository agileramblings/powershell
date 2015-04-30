$ie = new-object -com "InternetExplorer.Application"
$ie.navigate("http://www.plainenglish.co.uk/gobbledygook-generator.html")
$btn = $doc.getElementsByTagName("input")
$newBtn = $btn | ? {$_.value -eq "Generate some gobbledygook"}
$newBtn.click()
$response = $doc.getElementsByName("insight")
$newResponse = $response | ? {$_.name -eq "insight"}
Write-Host $newResponse.value