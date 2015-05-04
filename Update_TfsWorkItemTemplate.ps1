Function Update-TfsWorkItemTemplate() {
    [CmdLetBinding]
    begin{
        $sourcePath = "C:\TFS\Results\Default_2013_Scrum\"
        $targetPath = "C:\TFS\Results\Importable_2013_Scrum\"
    }
    process {
        $files = Get-ChildItem -path $sourcePath
        foreach ($file in $files){
            $sourceName = $file.FullName
            $targetName = $targetPath + "\" + $file.Name
            cat "$sourceName" | % {$_ -replace "Iteration ID", "IterationID"} > "$targetName"
            cat "$targetName" | % {$_ -replace "Area ID", "AreaID" } > "$targetName" 
        }
    }
    end{}
}
