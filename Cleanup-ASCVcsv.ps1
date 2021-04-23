# C:\Users\312127\Desktop\New_folder
$Logfile = "\\ctsfile\public\larkin1\ASCVScript.log"
Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}
$items = get-childitem -Path \\ctsfile\public\stlouis\ASCV_Pull_list\*.csv -Recurse | Where-Object {$_.LastWriteTime -lt (get-date).AddDays(-7)}
$items | Remove-Item
$rmitems = $items | select -Property Name
#[string]::Join("; ",$rmitems.name)
logwrite -logstring "----------------------------------------"
logwrite -logstring "Cleanup script removed: $([string]::Join("; ",$rmitems.name))"

