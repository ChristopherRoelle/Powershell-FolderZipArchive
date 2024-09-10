clear

$jobHistoryRoot = "" #Root Directory where job processing starts. 
$jobProcessingFolder = "" #Folder where log file will be placed. This script is automated, and uses its own job processing folder to manage individual run logs.
$statusFile = "log.txt" #name of the log file.

#Anything older than this is archived off (i.e. older than 6 months)
$dateThresholdMonths = 6

$maxBatchSize = 250 #How many directories to archive per batch?
$maxBatches = 20 # set to -1 for unlimited | This is how many batches will be ran. (i.e. maxBatchSize = 250 and maxBatches = 20 = 5000 directories per run.

$firstZippedFolder = $null
$lastZippedFolder = $null

#Compression Settings
$sevenZipExePath = "" #Point this to 7Zip's CLI exe file.
$compressionLevel = 5 #0-9 - https://7-zip.opensource.jp/chm/cmdline/switches/method.htm#ZipX
$Pass = "Password123" #Password

#System - DO NOT MODIFY
$caughtUp = $false
$batch = @()
$skippedPaths = @()
$batchCount = 1;

################################
# Functions
################################
function DisplayEndReport(){
    $elapsedTime = $(Get-Date) - $global:startTime
    $totalTime = "{0:HH:mm:ss.fff}" -f ([datetime]$elapsedTime.Ticks)

    Write-Output("")
    Write-Output("++====================================")
    Write-Output("||Threshold: {0:MM/dd/yyyy}" -f $($global:dateThreshold))
    Write-Output("||Runtime: $($totalTime)")
    Write-Output("||")
    Write-Output("||First Zipped: $($global:firstZippedFolder)")
    Write-Output("||Last Zipped: $($global:lastZippedFolder)")
    Write-Output("||")
    Write-Output("||Total Zipped: $($global:totalZipCount)")
    Write-Output("||Total Deleted: $($global:totalDeleteCount)")
    Write-Output("||Total Skipped: $($global:skippedPaths.Count)")
    Write-Output("++====================================")

    #Write to log
    Add-Content -Path $global:statusFile -Value $("<br><b>Runtime</b>: $($totalTime)")


    #Notify if we caught up
    if($global:caughtUp){
        Add-Content -Path $global:statusFile -Value $('<br><b style="color:green;">Zipping is caught up for threshold date!</b>')
    }

    #First and last zips
    Add-Content -Path $global:statusFile -Value $("<b>First Zipped</b>: $($global:firstZippedFolder)")
    Add-Content -Path $global:statusFile -Value $("<b>Last Zipped</b>: $($global:lastZippedFolder)")

    #Totals
    Add-Content -Path $global:statusFile -Value $("<br><b>Total Zipped</b>: $($global:totalZipCount)")
    Add-Content -Path $global:statusFile -Value $("<b>Total Deleted</b>: $($global:totalDeleteCount)")
    Add-Content -Path $global:statusFile -Value $("<b>Total Skipped</b>: $($global:skippedPaths.Count)")

    if($skippedPaths.Count -gt 0)
    {
        Write-Output("||Skipped Paths Due to Size")
        Add-Content -Path $global:statusFile -Value "<br><b>Skipped Paths Due to Size [Greater than $($global:maxFolderSizeMB) MB]</b>"
        foreach($path in $skippedPaths)
        {
            Write-Output("||$($path)")
            Add-Content -Path $global:statusFile -Value $path
        }
    }

}

function CompressFile()
{
    param (
        [Parameter(Mandatory=$true)]
        [string]$fileToEncrypt
    )

    #Get the destination based on parent path (folderName.zip)
    $destination = (Get-Item -Path $fileToEncrypt).Parent.FullName
    $destination = $destination + "\\" + (Get-Item -Path $fileToEncrypt).BaseName + ".zip"

    $password = $global:Pass

    #Write-Output("`nFile: $($fileToEncrypt)")
    #Write-Output("Pass: $($password)")

    $args = @("a -tzip"
          "$destination" 
          "$fileToEncrypt" 
          "-p$password" 
          "-mx=$compressionLevel"
          )

    Write-Output($fileToEncrypt)

    try{
        #Attempt to compress the file
        Start-Process -FilePath $sevenZipExePath -ArgumentList $args -NoNewWindow -Wait
        $global:totalZipCount++

        #Deleting original
        Remove-Item -Path $entry -Recurse
        $global:totalDeleteCount++
    } 
    catch {
        throw "Failed to compress! $($_)"
    }

}

function ZipBatch() {

    $curPath = "";

    #Zip the directories
    Write-Output("`tZipping...")

    try {
        #Attempt a zip of the entries in the batch. Record the current in case of error.
        foreach($entry in $batch)
        {
            $curPath = $entry

            CompressFile -fileToEncrypt $entry
            
        }

        #Update the first/last zipped
        if($global:firstZippedFolder -eq $null)
        {
            $global:firstZippedFolder = "<a href=`"$($entry).zip`">$($entry.Replace($jobHistoryRoot, '..\'))</a>"
        }

        $global:lastZippedFolder = "<a href=`"$($entry).zip`">$($entry.Replace($jobHistoryRoot, '..\'))</a>"

        Write-Output("`t`tDone!")
    } catch {
        #Output error file and exit
        Write-Output("`t`tError Zipping: $($curPath)")
        Write-Output("`t`t$($_)")
        DisplayEndReport
        exit 1
    }

    Write-Output("`t`tDone!")

    Write-Output("`tClearing batch...")
    $global:batch = @()
    $global:batchCount++
    Write-Output("`t`tDone!`n")


    #Check if we have ran the max number of batches, if so, quit, else continue scanning.
    if($batchCount -gt $maxBatches -and $maxBatches -ne -1)
    {
        Write-Output("`tMax batches exceeded, quitting...")
        DisplayEndReport
        exit
    }
}

################################
# Start Processing
################################

$startTime = Get-Date
$dateThreshold = Get-Date
$dateThreshold = $dateThreshold.AddMonths(-$dateThresholdMonths)
$thresholdDateHash = "{0:D4}{1:D2}" -f $dateThreshold.Year, $dateThreshold.Month
$totalZipCount = 0
$totalDeleteCount = 0

#Verify 7-Zip exists
if(!(Test-Path -Path $sevenZipExePath))
{
    Write-Output("Could not find 7-Zip.exe!")
    Write-Output("Please verify the path is valid.")
    exit 1
}

Write-Output("Threshold: {0:MM/dd/yyyy}`n" -f $($dateThreshold))
$statusFile = $jobProcessingFolder + '//' + $statusFile

#Create the log file if not exists
if(!(Test-Path -Path $statusFile))
{
    New-Item -Path $statusFile -ItemType 'file'
}

#Write the settings to the file
Add-Content -Path $statusFile -Value $("<b>Threshold</b>: {0:MM/dd/yyyy}" -f $($dateThreshold))
Add-Content -Path $statusFile -Value $("<b>Max Batch Size</b>: $($maxBatchSize)")
Add-Content -Path $statusFile -Value $("<b>Max Batches</b>: $($maxBatches)")

#loop the YYYY-MM folder structure
Get-ChildItem -Path $jobHistoryRoot | Where-Object{ $_.PSIsContainer -and $_.BaseName -ne 'Archive' } |
ForEach-Object {
    
    $maxDay = -1
    $yearMonthFolder = $_.FullName
    $folderName = Split-Path -Path $_.FullName -leaf
    Write-Output("`nRoot $($folderName)")
    
    $folderPart = $folderName.Split("-")

    #Check if month and year is equivalent to the threshold, if so, we only want to look at less than today's date.
    #else, we can look at all folders
    if([int]$folderPart[0] -eq $dateThreshold.Year -and [int]$folderPart[1] -eq $dateThreshold.Month)
    {
        #This is the max month/year combo for the threshold, so set the maxDay to the threshold day
        $maxDay = $dateThreshold.Day
    }
    else {
        #Set to the last day of the month
        $maxDay = [DateTime]::DaysInMonth($folderPart[0], $folderPart[1])
    }

    $folderDateHash = [int]("{0:D4}{1:D2}" -f $folderPart[0], $folderPart[1])

    #Determine if the folder is less than or equal to the age threshold
    #if($folderPart[0] -le $dateThreshold.Year -and $folderPart[1] -le $dateThreshold.Month)
        #2024                #2024                     #06               #01
    if($folderDateHash -le $thresholdDateHash)
    {
        #loop the day folders
        Get-ChildItem -Path $yearMonthFolder | Where-Object{ $_.PSIsContainer } |
        ForEach-Object {

            $dayFolder = $_.FullName
            $subFolderName = Split-Path -Path $_.FullName -leaf

            #Check the day is less than the $max Day
            if([int]$subFolderName -le [int]$maxDay)
            {
                Write-Output("|- $($subFolderName)")

                #Get the actual JHID folders, these can be archived
                Get-ChildItem -Path $dayFolder | Where-Object{ $_.PSIsContainer } | 
                ForEach-Object {
                    $curJHID = $_.FullName

                    #Add the folder's path to the batch
                    $batch += $curJHID

                    #Did we exceed the maxBatchSize? If so, time to zip the folders in the batch and clear
                    if($batch.Length -eq $maxBatchSize -and $maxBatchSize -ne -1)
                    {
                        Write-Output("`tBatch Size Met! Batch #$($batchCount)")

                        ZipBatch
                    }
                }
            }
            else
            {
                Write-Output("|- $($subFolderName) - [Outside of Threshold]")
            }
        }
    }
    else
    {
        Write-Output("`tNot old enough. Skipping.`n")
    }
}

Write-Output("")

#Zip any remaining
if($batch.Length -gt 0)
{
    Write-Output("Zipping remaining $($batch.Length) directories...")
    ZipBatch
    Write-Output("`tDone!")
}

#If we hit this, then we should be caught up. Set flag.
$global:caughtUp = $true

Write-Output("`nProcessing Complete!")

DisplayEndReport
