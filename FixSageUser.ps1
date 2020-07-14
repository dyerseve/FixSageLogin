<#
FixSageUser.ps1

For Sage 50 Premium Accounting 2018
Problem: Sage claims a user is already logged in or that they are out of licenses.
Causes: Usually this happens if the user doesn't log out properly, sage on the workstations crashes or goes unresponsive. Not entirely the end users fault in a lot of cases.
Solution: Run this to terminate one user or restart the entire service. I tested one user and it appears to work.

Put this in a shortcut on the server desktop with this command: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe c:\scripts\FixSageUser.ps1

Manual steps:
To restart database:
Type sc stop psqlwge and press Enter
Type sc start psqlwge and press Enter

To see who is "stuck" logged in (this shows other users as well not just stuck users):
net files | find "peachw"

You can disconnect open files in compmgmt.msc, the script closes all files for that user,
in this specific case it shouldn't impact anything but sage.

#>
#If you want to focus on Sage users only:
#$opfiles = (Get-SmbOpenFile | Where-Object -Property Path -Like "*peach*")
$opfiles = Get-SmbOpenFile

$opfileusernames = $opfiles.ClientUserName | select -uniq
#echo $opfileusernames

$menu = @{}
for ($i=1;$i -le $opfileusernames.count; $i++) 
{ Write-Host "$i. $($opfileusernames[$i-1])" 
$menu.Add($i,($opfileusernames[$i-1]))}
$lastoptioncount = ($opfileusernames.count)+1
Write-Host "$lastoptioncount. Restart Pervasive (all users must have sage closed)"
$menu.Add($lastoptioncount,"restart service")

[int]$ans = Read-Host 'Enter selection'
If ($ans -ne $lastoptioncount){
$selection = $menu.Item($ans)
Get-SmbOpenFile -ClientUserName $selection | Close-SmbOpenFile -Force
}
elseif ($ans -eq $lastoptioncount){
restart-service -name psqlWGE -force #-whatif
}