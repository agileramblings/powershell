# http://community.spiceworks.com/scripts/show/2784-powershell-script-to-truncate-sql-log-files-for-all-user-databases
# Script to truncate SQL log file for all user DBs on current SQL server
# Sam Boutros - 6/5/2014 - V1.0
# 7/28/14 - V1.1 - Cosmetic re-write..
# Truncate-Log.ps1
#
function Log {
    [CmdletBinding()]
    param(
        [Parameter (Mandatory=$true,Position=1,HelpMessage="String to be saved to log file and displayed to screen: ")][String]$String,
        [Parameter (Mandatory=$false,Position=2)][String]$Color = "White",
        [Parameter (Mandatory=$false,Position=3)][String]$Logfile = $myinvocation.mycommand.Name.Split(".")[0] + "_" + (Get-Date -format yyyyMMdd_hhmmsstt) + ".txt"
    )
    write-host $String -foregroundcolor $Color  
    ((Get-Date -format "yyyy.MM.dd hh:mm:ss tt") + ": " + $String) | out-file -Filepath $Logfile -append
}
#
# Import-Module SQLPS # See notes..
$Logfile = (Get-Location).path + "\Truncate_" + $env:COMPUTERNAME + (Get-Date -format yyyyMMdd_hhmmsstt) + ".txt"
# skipping first 4 databases: master, tempdb, model, msdb
(Invoke-SQLCMD -Query "SELECT * FROM sysdatabases WHERE dbid > 4") | ForEach-Object {
    $SQLLogString = "N'" + (Invoke-SQLCMD -Query ("SELECT name FROM sys.master_files WHERE database_id = " + $_.dbid + " AND type = 1;")).name + "'"
    Set-Location -Path ($Logfile.Split(":")[0] + ":")
    log ("Truncating log file $SQLLogString for database " + $_.name + " (database_id = " + $_.dbid + ")") Cyan $Logfile
    Invoke-SQLCMD -Query ("USE [" + $_.name + "]; ALTER DATABASE [" + $_.name + "] SET RECOVERY SIMPLE WITH NO_WAIT;")
    Invoke-SQLCMD -Query ("USE [" + $_.name + "]; DBCC SHRINKFILE($SQLLogString, 1); ALTER DATABASE [" + $_.name + "] SET RECOVERY FULL WITH NO_WAIT")
}