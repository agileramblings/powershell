function Get-GobbledyGook() {
<# 
    .SYNOPSIS
    Get a GobbledyGook saying
    .DESCRIPTION
    Get a GobbledyGook saying from http://www.plainenglish.co.uk. The resulting response is copied to the clipboard
    .EXAMPLE
    Get-GobbledyGook
    .LINK
    http://www.plainenglish.co.uk
#>

    [CmdLetBinding()]
    param()

    process {
        $ie = new-object -com "InternetExplorer.Application"
        $ie.navigate("http://www.plainenglish.co.uk/gobbledygook-generator.html")
        While($ie.Busy) { Start-Sleep -Milliseconds 100 }
        $doc = $ie.Document
        $btn = $doc.getElementsByTagName("input")
        $newBtn = $btn | ? {$_.value -eq "Generate some gobbledygook"}
        $newBtn.click()
        $response = $doc.getElementsByName("insight")
        $newResponse = $response | ? {$_.name -eq "insight"}
        $value = $newResponse.value
        Write-Verbose $value
        Write-Output $value | clip
    }           
} #end Function New-Directory

