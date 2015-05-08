Where I put Powershell Stuff
============================

These are PowerShell nuggets that I use all the time
-----------------------------------------------------

###Load PSGet
(new-object Net.WebClient).DownloadString("http://psget.net/GetPsGet.ps1") | iex

### PSReadLine
*After getting PSGet*

install-module PsReadLine

### Chocolatey
(new-object Net.WebClient).DownloadString("https://chocolatey.org/install.ps1") | iex
 
### ConsoleZ download page
https://github.com/cbucher/console/wiki/Downloads

http://stackoverflow.com/questions/7803076/can-console2-be-used-with-powershell
 
### Sample profile.ps1
https://github.com/agileramblings/powershell/blob/master/profile.ps1

### Generic Methods
https://gallery.technet.microsoft.com/scriptcenter/Invoke-Generic-Methods-bf7675af
