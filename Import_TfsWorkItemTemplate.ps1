function Import-TfsWorkItemTemplate() {
    [CmdLetBinding()]
    param($collectionUrl, $tpName, $sourcePath)

    begin{
        $filesToImport = Get-ChildItem -Path $sourcePath
    }
    process {
       foreach ($file in $filesToImport){
            
            if ($file.Extension -ne "xml") { continue }
            
            $fileName = $($file.FullName);
            witadmin importwitd /collection:"$collectionUrl" /p:"$tpName"  /f:"$fileName"
        }
    }
    end{}
}
