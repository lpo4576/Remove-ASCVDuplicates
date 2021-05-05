# 
# Remove-duplicates - This parses the ASCV pending list, outputting day 5 samples and any new samples in day 6
# Author            - Larkin O'quinn
# Date              - 040221
#
####--Changelog--####
# v1.0  - Initial version - 040221
# v1.1  - Added 5-6 days back filter to today's samples. Changed today's foreach loop to output a pscustomobject with sample ID and collection 
#        date, comparison loop to use and output pscustomobject, and final results to be output as csv - 040621
# v1.2  - Changed to include all day 5 samples, and uniques from day 6 - 040921
# v1.21 - Speed enhancements. $previoustxt now only contains day 6 samples. Added additional status message when export to file begins 
#         (this was taking the longest). Added invoke-item to automatically open the output in Excel. Moved changelog to top, because, 
#         sure. 040921
# v1.22 - Added debug switch for outputting file to local drive when developing. Added days back subtractor and prompt when in debug mode. 
#         Added more detail to output message, specifying day 5 and day 6 counts. Added logfile generation, with stopwatch. 041221
# v1.23 - Added env:Computername to logfile. 04/26/21
# v1.3  - Added seperate log filepath for testing. Added dynamic property naming for COLL_DATE, so users will no longer have to sort the
#         table ascending prior to exporting from the utility. Replaced all instances of '+=' in the script with Generic.List add 
#         methods. LO 05/04/21
#       - Changed output to place day 5 and day 6 samples in separate csvs. Adjusted output variables and summary text as such. Adjusted logwrite
#         calls to properly interpret length of new Generic.List arrays. LO 05/05/21

$debug = $false
$stopwatch =  [system.diagnostics.stopwatch]::StartNew()

Write-host "Please select yesterday's ASCV file" -BackgroundColor Black
Add-Type -AssemblyName System.Windows.Forms
$oldfile = New-Object System.Windows.Forms.OpenFileDialog
$oldfile.filter = "csv (*.csv)| *.csv"
$oldfile.Title = "Please select yesterday's ASCV file"
[void]$oldfile.ShowDialog()
$oldfile.FileName

Write-host "Please select today's ASCV file" -BackgroundColor Black
Add-Type -AssemblyName System.Windows.Forms
$newfile = New-Object System.Windows.Forms.OpenFileDialog
$newfile.filter = "csv (*.csv)| *.csv"
$newfile.Title = "Please select today's ASCV file"
[void]$newfile.ShowDialog()
$newfile.FileName
if ($debug -eq $true) {
    write-host ""
    $daysback = Read-host "Days Back?"
    }

Write-host ""
Write-host ""
write-host "Working..." -ForegroundColor Yellow


$previous = import-csv -Path $oldfile.filename
$new = import-csv -Path $newfile.filename
[System.Collections.Generic.List[string]]$previoustxt = @()
[string[]]$newtxt = $null
[string[]]$uniques = $null
$duplicates = 0
$unique5count = 0
$unique6count = 0
[System.Collections.Generic.List[pscustomobject]]$tempresults = @()
[System.Collections.Generic.List[pscustomobject]]$finalresults5 = @()
[System.Collections.Generic.List[pscustomobject]]$finalresults6 = @()
if ($debug -eq $false) {
    $Logfile = "\\ctsfile\public\larkin1\ASCVScript.log"
    }
    else {$logfile = "C:\Users\312127\Desktop\Testing\ASCVScriptTEST.log"}

Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

#Gets date, finds 5 and 6 days back, and if in debug, subtracts days ago
$date = Get-Date
if ($debug -eq $true) {
    $date = $date.adddays(-$daysback)
    }
$date5 = $date.AddDays(-5)
$date6 = $date.AddDays(-6)
logwrite -logstring "----------------------------------------"
Logwrite -logstring "Ran on $date, by $env:username. Day 5 was $($date5.toshortdatestring()). Day 6 was $($date6.toshortdatestring())."
Logwrite -logstring "Ran on $env:COMPUTERNAME"
logwrite -logstring "Old file: $($oldfile.FileName)"
logwrite -logstring "New file: $($newfile.FileName)"

#Place previous day's day 6 samples into string
$collname = $(($previous | gm) -match "COLL").Name
foreach ($item in $previous) {
    if ($item.$collname -eq $date6.ToShortDateString()) {
        $null = $previoustxt.Add($item.Unit_ID)
        }  
    }
logwrite -logstring "$(($previoustxt | Measure-Object).count) day 6 samples found in previous file"

#Place today's samples into customobject
$ncollname = $(($new | gm) -match "COLL").Name
foreach ($item in $new) {
    if ($item.$ncollname -eq $date5.ToShortDateString() -or $item.$ncollname -eq $date6.ToShortDateString()) {
        if ($item.Unit_ID -like "W*") {
            $results = '' | select 'Sample ID', CollectionDate
            $results.'Sample ID' = $item.UNIT_ID
            $results.CollectionDate = $item.$ncollname
            $null = $tempresults.Add($results)
            }
        }
    }
logwrite -logstring "$(($tempresults | Measure-Object).count) day 5 and 6 samples found in new file"

#Check previous day's string for presence of today's samples, and keep if unique
foreach ($sample in $tempresults) {
    if ($sample.CollectionDate -eq $date6.ToShortDateString()) {
        if ($previoustxt -notcontains $sample.'Sample ID') {
            $null = $finalresults6.Add($sample)          
            $unique6count ++
            }
        else {$duplicates ++}
        }
    if ($sample.CollectionDate -eq $date5.ToShortDateString()) {
        $null = $finalresults5.Add($sample)
        $unique5count ++
        }
    }
logwrite -logstring "$unique5count day 5 samples output. $unique6count day 6 samples output. $duplicates day 6 duplicates removed"

Write-host ""
Write-host "Exporting to files..." -ForegroundColor Yellow

$numdate = (get-date).ToString('MMddyyyy')
if ($debug -eq $false) {
    $outputpath5 = "\\ctsfile\public\stlouis\ascv_pull_list\Finished\$numdate day 5 ASCV.csv"
    $outputpath6 = "\\ctsfile\public\stlouis\ascv_pull_list\Finished\$numdate new day 6 ASCV.csv"
    }
Else {
    $outputpath5 = "C:\Users\312127\Desktop\Testing\$numdate day 5 ASCV.csv"
    $outputpath6 = "C:\Users\312127\Desktop\Testing\$numdate new day 6 ASCV.csv"
    }
$finalresults5 | Export-csv -Path $outputpath5 -NoTypeInformation
$finalresults6 | Export-csv -Path $outputpath6 -NoTypeInformation

Write-Host ""
Write-Host "$duplicates duplicates from $($date6.ToShortDateString()) have been removed" -ForegroundColor Green
Write-Host "$unique6count new samples from collection date $($date6.ToShortDateString()) have been added to '$numdate new day 6 ASCV.csv'" -ForegroundColor Green
Write-host "$unique5count new samples from collection date $($date5.ToShortDateString()) have been added to '$numdate day 5 ASCV.csv'" -ForegroundColor Green
#Write-host "These samples have been added to '$numdate day 5 ASCV.csv' and '$numdate new day 6 ASCV.csv'," -ForegroundColor Green
Write-host "These files have been placed in S:\ASCV_Pull_list\Finished" -ForegroundColor Green
Write-Host ""

invoke-item -Path $outputpath5
invoke-item -Path $outputpath6

logwrite -logstring "Time taken: $($stopwatch.Elapsed.TotalSeconds)"
logwrite -logstring "Errors: $error"
Pause