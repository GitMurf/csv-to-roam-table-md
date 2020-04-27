#v0.2.2
#Version Comments: Created a summary page for the CSV Import/Conversion
#Repository: https://github.com/GitMurf/csv-to-roam-table-md
#Code written by:       Murf
#Design/Concept by:     Rob Haisfield @RobertHaisfield on Twitter

#Set the indent type. In Roam a single space at beginning of a line works just like a TAB.
#Can use either way to bring into Roam, just your preference. Default we will keep simple and just use Spaces " ".
#If you want to use tab, use $indentType = "`t"
$indentType = " "

#Bullet type. Leave blank if don't need to show a character for bullets which Roam does NOT need to import into table format
#Can use for example "*" or "-"
$bulletType = "-"

#Set the delimiter variable (default is "," comma)
$strDelim = ","

#This function writes to a specified file with specified text
Function Write-Roam-File
{
    Param(
        [string]$filePath,
        [string]$strToWrite = ""
    )

    Add-content -LiteralPath $filePath -value $strToWrite
    $logInfo = "Added '$strToWrite' to the File: '$filePath'"
    Write-Roam-Log $logInfo
}

#This function writes to log file and echos on screen if "show" parameter is present
Function Write-Roam-Log
{
    Param(
        [string]$logstring,
        [string]$show = "Hide"
    )

    #If passed the show parameter then write to the powershell window
    if($show.ToLower() -eq "show"){Write-Host($logstring)}
    if($logstring){$logstring = ("[" + (Get-Date) + "] $logstring")}
    Add-content -LiteralPath $tempLogFile -value $logstring
}

#Add a blank line for easier reading of prompts in powershell window
Write-Host

#Ask for user input to create pages for each row, otherwise will just default to creating a single markdown file with the table markdown for Roam
$respPages = Read-Host "Do you want to create a Page for each Row in the CSV file? (Enter y or n)"

#Check if user decided to create new pages (e.g., a CRM import)
if($respPages -eq "y" -or $respPages -eq "Y" -or $respPages -eq "yes" -or $respPages -eq "Yes")
{
    $bPages = $true
    Write-Host
    #Ask user for the type of csv import (e.g., People, Company, CRM etc.)
    $csvType = Read-Host "Enter CSV Type (e.g., Contacts, Tools, Company). Will use in csv-type:: Attribute"
}else{$bPages = $false}

Write-Host

#Ask for user input. If left blank and user presses ENTER, then continue with default (comma). Otherwise they can enter their own option.
$respDelim = Read-Host "Default CSV Delimiter is '$strDelim' (comma). Press ENTER to Continue or input 'n' to change it."

Write-Host

#Check if user decided to change to a different delimiter
if($respDelim -eq "n" -or $respDelim -eq "N" -or $respDelim -eq "'n'" -or $respDelim -eq "'N'")
{
    $strDelim = Read-Host "What would you like your CSV Delimiter to be? (If Tab delimited, enter 'TAB')"
    if($strDelim -eq "TAB" -or $strDelim -eq "tab" -or $strDelim -eq "'TAB'" -or $strDelim -eq "'tab'"){$strDelim = "`t"}
}

Write-Host

#Get path where script is running from so you can target CSV
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

#Ask for user input. If left blank and user presses ENTER, then use the script path. Otherwise they can enter their own custom path.
$respPath = Read-Host "Is your target CSV file located here: '$scriptPath'? Press ENTER to Continue or input 'n' to change it."

Write-Host

#Check if user decided to change to a different path
if($respPath -eq "n" -or $respPath -eq "N" -or $respPath -eq "'n'" -or $respPath -eq "'N'")
{
    $scriptPath = Read-Host "Enter the folder path of your CSV file (do NOT include the file name)"
}

Write-Host

#Get the file name from user and create the full path
$fileNameStr = Read-Host "Name of CSV file with extension? Do NOT include path. Example: FileName.csv"
$fileNameStrPath = $scriptPath + "\" + $fileNameStr

Write-Host

#Set the Results folder to store all the outputs from the script
$resultsFolder = "$scriptPath\Results"

#Get a date string down to the second we can add to new markdown file we will be creating so no duplicates if we run multiple times
$fullDateStr = get-date
$dateStrName = $fullDateStr.ToString("yyyy_MM_dd-HH_mm_ss")
$csvFileName = "$fileNameStr" + "_$dateStrName"
$newMarkdownFile = "$resultsFolder\" + "$csvFileName" + ".md"
$tempLogFile = "$resultsFolder\" + "roamCsvLog" + "_$dateStrName" + ".log"

#Import .CSV file into a Variable to loop through and parse
$csvObject = Import-Csv -Delimiter $strDelim -Path "$fileNameStrPath"

#If $bPages -eq $true, then create the csv-import page name to store all the info about this import and the pages it creates (summary and log)
if($bPages)
{
    $csvImportName = "[[csv-import]] - " + $csvFileName
    $csvImportNamePath = "$resultsFolder\" + "$csvImportName" + ".md"
    #Create Results folder if it doesn't already exist
    if(!(Test-Path $resultsFolder)){New-Item -ItemType Directory -Force -Path $resultsFolder | Out-Null}

    #Get current date and put into Roam format
    $roamMonth = $fullDateStr.ToString("MMMM") #April, August, January
    $roamDay1 = $fullDateStr.ToString("%d") #1, 13, 28
    $roamYear = $fullDateStr.ToString("yyyy") #2020

    #Find the "day" suffix for Roam format (i.e., Xst, Xnd, Xrd, Xth)
    $roamDay2 = switch($roamDay1)
    {
        {$_ -in 1,21,31} {$roamDay1 + 'st'; break}
        {$_ -in 2,22} {$roamDay1 + 'nd'; break}
        {$_ -in 3,23} {$roamDay1 + 'rd'; break}
        Default {$roamDay1 + 'th'; break}
    }
    $roamDate = "[[" + "$roamMonth $roamDay2, $roamYear" + "]]" #[[April 26th, 2020]]
    #Time string
    $strTime = $fullDateStr.ToString("HH:mm") #17:43, 05:21
    
    Write-Roam-Log "Created the .MD markdown file '$csvImportName' that will Summarize the CSV Conversion activities and store a Log of all actions."
    Write-Roam-Log "Converted today's date to Roam format: $roamDate"

    #Write attribute for csv-import to first line of this new .md file (need to use LiteralPath parameter because of [[]] characters in path)
    Write-Roam-File $csvImportNamePath ("csv-date:: " + $roamDate)
    #Import time attribute
    Write-Roam-File $csvImportNamePath ("csv-time:: " + $strTime)
    #Filename attribute
    Write-Roam-File $csvImportNamePath ("csv-filename:: " + $fileNameStr)
    #Type of CSV file attribute (example could be: People, CRM, Company)
    Write-Roam-File $csvImportNamePath ("csv-type:: " + $csvType)
    
    Write-Roam-Log "Finished adding primary Attributes to the CSV Conversion Summary page: '$csvImportName'"
}

#TESTING... Exit the script
Read-Host -Prompt "Script complete. Press any key to exit."
Exit

#Collapse the entire table under a parent bullet with name of the CSV file
$tableCell = "TABLE IMPORT FROM CSV: " + $fileNameStr
$tableCell = $bulletType + $tableCell
Add-content -LiteralPath $newMarkdownFile -value $tableCell

#Add {{table}}
$tableCell = "{{table}}"
$tableCell = $bulletType + $tableCell
$tableCell = $indentType + $tableCell
Add-content -LiteralPath $newMarkdownFile -value $tableCell

#Start by adding the table header to the markdown results file
$ctr = 2
foreach($col in $csvObject[0].psobject.properties.name)
{
    $tableCell = $col
    $tableCell = $bulletType + $tableCell
    #Add the proper indentation based on looping through x number of times based on $ctr
    $tmpCtr = $ctr
    while($tmpCtr -gt 0)
    {
        $tableCell = $indentType + $tableCell
        $tmpCtr = $tmpCtr - 1
    }

    Add-content -LiteralPath $newMarkdownFile -value $tableCell
    $ctr = $ctr + 1
}

#Loop through each row of the csv file
foreach($row in $csvObject)
{
    #Set a counter which will decide how spacing is done for indents in the Roam table structure
    #Start at 2 instead of 0 to account for CSV file name parent bullet and then {{table}} being second indent level, and everything needing to start indented under it
    $ctr = 2
    #For each row of csv file, loop through each column
    foreach($col in $row.psobject.properties.name)
    {
        $tableCell = $row.$col
        $tableCell = $bulletType + $tableCell
        #Add the proper indentation based on looping through x number of times based on $ctr
        $tmpCtr = $ctr
        while($tmpCtr -gt 0)
        {
            $tableCell = $indentType + $tableCell
            $tmpCtr = $tmpCtr - 1
        }

        Add-content -LiteralPath $newMarkdownFile -value $tableCell
        $ctr = $ctr + 1
    }
}

#Exit the script
Read-Host -Prompt "Script complete. Press any key to exit."
Exit