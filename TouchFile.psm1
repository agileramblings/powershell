function Invoke-TouchFile([string[]]$paths) {

<#
.SYNOPSIS
 A simple implementation of the unix "touch" 
.DESCRIPTION
Touch-File will create a file if none exists or will update the file system timestamp if the file does exist
.EXAMPLE
Touch-File x.txt 
.EXAMPLE
Touch-File x
.PARAMETER path
The path, ending with the file name, that you would like to create or update the file system timestamp on
#>

	begin {
          function updateFileSystemInfo([System.IO.FileSystemInfo]$fsInfo) {
               $datetime = get-date
               $fsInfo.CreationTime = $datetime
               $fsInfo.LastWriteTime = $datetime
               $fsInfo.LastAccessTime = $datetime
          }

          function touchExistingFile($arg) {
               if ($arg -is [System.IO.FileSystemInfo]) {
                    updateFileSystemInfo($arg)
               } else {
                    $resolvedPaths = resolve-path $arg
                    foreach ($rpath in $resolvedPaths) {
                         if (test-path -type Container $rpath) {
                              $fsInfo = new-object System.IO.DirectoryInfo($rpath)
                         } else {
                              $fsInfo = new-object System.IO.FileInfo($rpath)
                         }
                         updateFileSystemInfo($fsInfo)
                    }
               }
          }
          
          function touchNewFile([string]$path) {
             $null > $path
          }
    }

    process {
 	    if ($_) {
 		    if (test-path $_) {
 			    touchExistingFile($_)
 		    } else {
 			    touchNewFile($_)
 		    }
 	    }
     }

    end {
	    if ($paths) {
 		    foreach ($path in $paths) {
 			    if (test-path $path) {
 				    touchExistingFile($path)
 			    } else {
 				    touchNewFile($path)
 			    }
 		    }
 	    }
 	}
 }

Set-Alias touch Invoke-TouchFile

Export-ModuleMember -Alias *
Export-ModuleMember -Function *

