# Powershell-JobProcessingZipArchive
Archives automated job processing stored in ..\YYYY-MM\DD\.. directory structure.

## What is it?
This script is designed to Zip, Password-Protect, then delete the original files.
The implementation expects jobs to be stored in the following structure:

ROOT\YYYY-MM\DD\JOB DIRECTORY

## How it works
The job creates a hash of today's date minus the months you want to go back.
This is compared with the YYYY-MM folders to determine how far back we need to look.
Then we recursively look at the DD folders within the eligible YYYY-MM folders.
We recursively check the DD folders up to (and including) today's date minus the thresholdMonths.
For each of these folders stored in the DD folder (now refered to as JobHistoryID folders, per original intention),
we zip the folder with a password (using 7-Zip encryption), then delete the original file. The Zipped file is placed back
into the original folder.

i.e. ROOT\2024-01\01\JHID12345
We would be left with:
ROOT\2024-01\01\JHID12345.zip

## Requirements
* Directory structure: ROOT\YYYY-MM\DD\*
* 7-Zip - $sevenZipExePath must be pointed here

## Optional
The script will output a log.txt file which contains HTML tags.
This can be injected into an HTML eMail's body to produce an HTML log email with run details.
