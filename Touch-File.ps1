function Touch-File([string[]]$paths) {
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
             $null > $pathTest
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
