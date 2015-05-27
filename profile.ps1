[string]$Global:currentEnv = ""

function Set-Tfs2013 (){
    if (![string]::IsNullOrEmpty($Global:currentEnv)){
        Write-Host "The Visual Studio ($($Global:currentEnv)) environmental variables for this session have already been set. Start a new PowerShell session." -ForegroundColor Red
        return
    }
    # http://stackoverflow.com/questions/2124753/how-i-can-use-powershell-with-the-visual-studio-command-prompt
    # Set environment variables for Visual Studio Command Prompt
    if (Test-Path -Path 'c:\Program Files (x86)\Microsoft Visual Studio 12.0\VC') { 
        $Global:currentEnv = "VS 2013"
        pushd 'c:\Program Files (x86)\Microsoft Visual Studio 12.0\VC'
        cmd /c "vcvarsall.bat&set" |
        foreach {
          if ($_ -match "=") {
            $v = $_.split("="); set-item -force -path "ENV:\$($v[0])"  -value "$($v[1])"
          }
        }
        popd
        write-host "`nVisual Studio 2013 Command Prompt variables set." -ForegroundColor Yellow

        #add to env variables
        if (Test-Path -Path 'C:\Program Files (x86)\Microsoft Team Foundation Server 2013 Power Tools') {
            $env:Path += ";C:\Program Files (x86)\Microsoft Team Foundation Server 2013 Power Tools"
        } else {
            Write-Verbose "Team Foundation Server 2013 Power Tools are unavailable."
        }
    } else {
        Write-Verbose "Visual Studio 2013 Admin Console Tools are not available."
    }
}

function Set-Tfs2010 (){
    if (![string]::IsNullOrEmpty($Global:currentEnv)){
        Write-Host "The Visual Studio ($($Global:currentEnv)) environmental variables for this session have already been set. Start a new PowerShell session." -ForegroundColor Red
        return
    }

    # Set environment variables for Visual Studio Command Prompt
    if (Test-Path -Path 'c:\Program Files (x86)\Microsoft Visual Studio 10.0\VC') { 
       $Global:currentEnv = "VS 2010"
        pushd 'c:\Program Files (x86)\Microsoft Visual Studio 10.0\VC'
        cmd /c "vcvarsall.bat&set" |
        foreach {
          if ($_ -match "=") {
            $v = $_.split("="); set-item -force -path "ENV:\$($v[0])"  -value "$($v[1])"
          }
        }
        popd
        write-host "`nVisual Studio 2010 Command Prompt variables set." -ForegroundColor Yellow

        #add to env variables
        if (Test-Path -Path 'C:\Program Files (x86)\Microsoft Team Foundation Server 2010 Power Tools') {
            $env:Path += ";C:\Program Files (x86)\Microsoft Team Foundation Server 2010 Power Tools"
        }else {
            Write-Verbose "Team Foundation Server 2010 Power Tools are unavailable."
        }
    }
    else {
        Write-Verbose "Visual Studio 2010 Admin Console Tools are not available."
    }
}

if ($host.Name -eq 'ConsoleHost')
{
    $mod = Get-Module | ? {$_.Name -eq "PsReadLine"}
    if ($mod -ne $null){

        Import-Module PSReadline

        #region Smart Insert/Delete

        # The next four key handlers are designed to make entering matched quotes
        # parens, and braces a nicer experience.  I'd like to include functions
        # in the module that do this, but this implementation still isn't as smart
        # as ReSharper, so I'm just providing it as a sample.

        Set-PSReadlineKeyHandler -Chord 'Oem7','Shift+Oem7' `
                                 -BriefDescription SmartInsertQuote `
                                 -LongDescription "Insert paired quotes if not already on a quote" `
                                 -ScriptBlock {
            param($key, $arg)

            $line = $null
            $cursor = $null
            [PSConsoleUtilities.PSConsoleReadline]::GetBufferState([ref]$line, [ref]$cursor)

            if ($line[$cursor] -eq $key.KeyChar) {
                # Just move the cursor
                [PSConsoleUtilities.PSConsoleReadline]::SetCursorPosition($cursor + 1)
            }
            else {
                # Insert matching quotes, move cursor to be in between the quotes
                [PSConsoleUtilities.PSConsoleReadline]::Insert("$($key.KeyChar)" * 2)
                [PSConsoleUtilities.PSConsoleReadline]::GetBufferState([ref]$line, [ref]$cursor)
                [PSConsoleUtilities.PSConsoleReadline]::SetCursorPosition($cursor - 1)
            }
        }

        Set-PSReadlineKeyHandler -Key '(','{','[' `
                                 -BriefDescription InsertPairedBraces `
                                 -LongDescription "Insert matching braces" `
                                 -ScriptBlock {
            param($key, $arg)

            $closeChar = switch ($key.KeyChar)
            {
                <#case#> '(' { [char]')'; break }
                <#case#> '{' { [char]'}'; break }
                <#case#> '[' { [char]']'; break }
            }

            [PSConsoleUtilities.PSConsoleReadLine]::Insert("$($key.KeyChar)$closeChar")
            $line = $null
            $cursor = $null
            [PSConsoleUtilities.PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor)
            [PSConsoleUtilities.PSConsoleReadLine]::SetCursorPosition($cursor - 1)        
        }

        Set-PSReadlineKeyHandler -Key ')',']','}' `
                                 -BriefDescription SmartCloseBraces `
                                 -LongDescription "Insert closing brace or skip" `
                                 -ScriptBlock {
            param($key, $arg)

            $line = $null
            $cursor = $null
            [PSConsoleUtilities.PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor)

            if ($line[$cursor] -eq $key.KeyChar)
            {
                [PSConsoleUtilities.PSConsoleReadLine]::SetCursorPosition($cursor + 1)
            }
            else
            {
                [PSConsoleUtilities.PSConsoleReadLine]::Insert("$($key.KeyChar)")
            }
        }

        Set-PSReadlineKeyHandler -Key Backspace `
                                 -BriefDescription SmartBackspace `
                                 -LongDescription "Delete previous character or matching quotes/parens/braces" `
                                 -ScriptBlock {
            param($key, $arg)

            $line = $null
            $cursor = $null
            [PSConsoleUtilities.PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor)

            if ($cursor -gt 0)
            {
                $toMatch = $null
                switch ($line[$cursor])
                {
                    <#case#> '"' { $toMatch = '"'; break }
                    <#case#> "'" { $toMatch = "'"; break }
                    <#case#> ')' { $toMatch = '('; break }
                    <#case#> ']' { $toMatch = '['; break }
                    <#case#> '}' { $toMatch = '{'; break }
                }

                if ($toMatch -ne $null -and $line[$cursor-1] -eq $toMatch)
                {
                    [PSConsoleUtilities.PSConsoleReadLine]::Delete($cursor - 1, 2)
                }
                else
                {
                    [PSConsoleUtilities.PSConsoleReadLine]::BackwardDeleteChar($key, $arg)
                }
            }
        }

        #endregion Smart Insert/Delete
    }
}

function Make-PSDrive(){
    [CmdLetBinding()]
    param(
        [parameter(Mandatory=$true)]
        [string]$name,

        [parameter(Mandatory=$true)]
        [string]$root
    )

    begin{}
    process {
        $d = Get-PSDrive | ? { $_.name -eq $name}
        if ($d -eq $null)
        {
            New-PSDrive $name -PSProvider FileSystem -Root $root
        }
    }
    end{}
}

#add PSDrive to ps scripts folder
Make-PSDrive -Name "scripts" -Root "\\cocdata1\$($env:username)`$\TFS"
Make-PSDrive -Name "modules" -Root "\\cocdata1\$($env:username)`$\data\WindowsPowerShell\Modules"
Make-PSDrive -Name "powershell" -Root "\\cocdata1\$($env:username)`$\data\WindowsPowerShell"

