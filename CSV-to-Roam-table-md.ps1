#v0.3.2
#Version Comments: Attributes cannot have bullets
#Repository: https://github.com/GitMurf/csv-to-roam-table-md
#Code written by:       Murf
#Design/Concept by:     Rob Haisfield @RobertHaisfield on Twitter

#If $bTesting = $true then add "TESTING_ to the front of any page created
$bTesting = $true

#Array variables
$arrSummary = @()
$arrTable = @()
$arrLog = @()

#Page counter
$pgCtr = 0
$attrCtr = 0

#Set the indent type. In Roam a single space at beginning of a line works just like a TAB.
#Can use either way to bring into Roam, just your preference. Default we will keep simple and just use Spaces " ".
#If you want to use tab, use $indentType = "`t"
$indentType = " "

#Bullet type. Leave blank if don't need to show a character for bullets which Roam does NOT need to import into table format
#Can use for example "* " or "- " or ""
#NOTE: Whatever bullet type you use (other than if you use non and leave it empty) you NEED to add a space after it
#NOTE: Attribute fields canNOT have bullets added in front so just leave without a bullet
$bulletType = "- "

#Set the delimiter variable (default is "," comma)
$strDelim = ","

#Root bullets to nest results below under
$arrSummary += , ($bulletType + "SUMMARY") #Collapse page and attribute names created under this main bullet
$arrTable += , ($bulletType + "TABLE") #Collapse the entire table under a parent bullet for Table
$arrLog += , ($bulletType + "LOGS") #Creating "LOGS" parent bullet to have all logs nested under it

#This function writes to a specified file with specified text
Function Write-Roam-File
{
    Param(
        [string]$filePath,
        [string]$strToWrite = ""
    )

    Add-content -LiteralPath $filePath -value $strToWrite
    $logInfo = $indentType + $indentType + $bulletType + "Added '$strToWrite' to the File '$filePath'"
    $logInfo = $logInfo -Replace "\:","_" -Replace "\{","_" -Replace "\}","_"
    $script:arrLog += , $logInfo
}

#Add a blank line for easier reading of prompts in powershell window
Write-Host

#Ask for user input to create pages for each row, otherwise will just default to creating a single markdown file with the table markdown for Roam
$respPages = Read-Host "Do you want to create a Page for each Row in the CSV file? (Enter y or n)"

#Check if user decided to create new pages (e.g., a CRM import)
if($respPages -eq "y" -or $respPages -eq "Y" -or $respPages -eq "yes" -or $respPages -eq "Yes"){$bPages = $true}else{$bPages = $false}

Write-Host

#Ask user for the type of csv import (e.g., People, Company, CRM etc.)
$csvType = Read-Host "Enter the Type/Category of your CSV data (e.g., Contacts, Books, Videos, etc.) to allow for searching of similar data types in Roam"

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

#Try to find the CSV to scan automatically by looking in current folder and also selecting one CSV file sorting by most recent edited
$foundCsv = Get-ChildItem -Path $scriptPath -Filter *.csv | Sort-Object -Property LastWriteTime -Descending | Select-Object -first 1
$foundName = $foundCsv.Name

if($foundPath -ne "")
{
    $respFound = Read-Host "Is this your target CSV: '$foundName'? Press ENTER to Continue or input 'n' to select a different CSV"
    Write-Host
}

if($respFound -ne "")
{
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
}
else
{
    #The CSV that was auto picked is the correct one (CSV in same folder as script)
    #Get the file name and create the full path
    $fileNameStr = $foundName
    $fileNameStrPath = $foundCsv.FullName
    Write-Host
}

#Set the Results folder to store all the outputs from the script
$resultsFolder = "$scriptPath\Results"

#Get a date string down to the second we can add to new markdown file we will be creating so no duplicates if we run multiple times
$fullDateStr = get-date
$dateStrName = $fullDateStr.ToString("yyyyMMdd_HHmmss")
$csvFileName = "$fileNameStr" + "_$dateStrName"

#Import .CSV file into a Variable to loop through and parse
$csvObject = Import-Csv -Delimiter $strDelim -Path "$fileNameStrPath"

#Create the csv-import page name to store all the info about this import and the pages it creates (summary and log)
$csvImportName = "CSV_import_" + $csvFileName
if($bTesting){$csvImportName = "TESTING_" + $csvImportName}
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

$arrLog += , ($indentType + $bulletType + "Created the .MD markdown file '$csvImportName' that will Summarize the CSV Conversion activities and store a Log of all actions.")
$arrLog += , ($indentType + $bulletType + "Converted today's date to Roam format: $roamDate")

$pgCtr = $pgCtr + 1
#Write attribute for csv-import to first line of this new .md file (need to use LiteralPath parameter because of [[]] characters in path)
Write-Roam-File $csvImportNamePath ("csv-date:: " + $roamDate)
#Import time attribute
Write-Roam-File $csvImportNamePath ("csv-time:: " + $strTime)
#Filename attribute
Write-Roam-File $csvImportNamePath ("csv-filename:: " + $fileNameStr)
#Type of CSV file attribute (example could be: People, CRM, Company)
Write-Roam-File $csvImportNamePath ("csv-type:: " + $csvType)

$arrLog += , ($indentType + $bulletType + "Finished adding primary Attributes to the CSV Conversion Summary page: '$csvImportName'")

#Creation of the CSV into Roam table format
$arrLog += , ($indentType + $bulletType + "Converting the CSV into the Roam table markdown format.")

#Add {{table}}
$tableCell = "{{table}}"
$tableCell = $bulletType + $tableCell
$tableCell = $indentType + $tableCell
$arrTable += , $tableCell

#Start by adding the table header
$arrLog += , ($indentType + $bulletType + "Adding table headers to $csvFileName")
if($bPages)
{
    $arrSummary += , ($indentType + $bulletType + "ATTRIBUTES") #Will create links to each attribute created under this bullet
}
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

    $arrTable += , $tableCell
    $ctr = $ctr + 1

    #If creating new pages for each CSV row, then need to add attributes to the summary page
    if($bPages -and $ctr -gt 3) #Need to skip the first column because that is what you are creating pages from
    {
        if($bTesting){$col = "TESTING_" + $col}
        $arrSummary += , ($indentType + $indentType + $bulletType + "#[[" + $col + "]]")
        $attrCtr = $attrCtr + 1
    }
}

if($bPages)
{
    #Create new page/file for each CSV row
    $arrLog += , ($indentType + $bulletType + "Creating new Pages for each CSV row")
    $arrSummary += , ($indentType + $bulletType + "PAGES CREATED")
}

#Loop through each row of the csv file
foreach($row in $csvObject)
{
    #Create new page/file for each CSV row
    if($bPages)
    {
        $colHeaderNames = $row.psobject.properties.name
        $rowPageName = $row.($colHeaderNames[0])
        if($bTesting){$rowPageName = "TESTING_" + $rowPageName}
        $arrSummary += , ($indentType + $indentType + $bulletType + "[[" + $rowPageName + "]]")
        $rowPageNamePath = "$resultsFolder\" + "$rowPageName" + ".md"
        $pgCtr = $pgCtr + 1

        #Commenting out the CSV import attribute data becuase isn't needed on each page... instead link to the csv summary page which has all that info
        #Check if any of the Windows filename illegal characters are present and if so, do NOT write to the file and instead just store the attributes on the summary page
            #Then the user can go into Summary page in Roam and copy the attributes, click the page name and then add there so that can keep the special character in name
            #The characters not allowed are: \ / : * ? " < > |
        $bInvalidChar = $false
        if($rowPageName.Contains("\") -or $rowPageName.Contains("/") -or $rowPageName.Contains(":") -or $rowPageName.Contains("*") -or $rowPageName.Contains("?") -or $rowPageName.Contains('"') -or $rowPageName.Contains("<") -or $rowPageName.Contains(">") -or $rowPageName.Contains("|"))
        {
            $bInvalidChar = $true
            $arrLog += , ($indentType + $bulletType + "**Invalid character** for Windows found in Filename for PAGE: [[" + $rowPageName + "]]")
        }
        else{$arrLog += , ($indentType + $bulletType + "Created the Page: [[" + $rowPageName + "]]")}

        if($bInvalidChar)
        {
            #Add under each page name in summary as this is what we will do if a bad character for Windows in file name
            $arrSummary += , ($indentType + $indentType + $indentType + "csv-import:: [[" + $csvImportName + "]]")
            #General attributes for the CSV import. These are in the Summary page for the import so do we need them also on every page?
            #$arrSummary += , ($indentType + $indentType + $indentType + $bulletType + "csv-date:: " + $roamDate)
            #$arrSummary += , ($indentType + $indentType + $indentType + $bulletType + "csv-time:: " + $strTime)
            #$arrSummary += , ($indentType + $indentType + $indentType + $bulletType + "csv-filename:: " + $fileNameStr)
            #$arrSummary += , ($indentType + $indentType + $indentType + $bulletType + "csv-type:: " + $csvType)
        }
        else
        {
            Write-Roam-File $rowPageNamePath ("csv-import:: [[" + $csvImportName + "]]")
            #General attributes for the CSV import. These are in the Summary page for the import so do we need them also on every page?
            #Write-Roam-File $rowPageNamePath ("csv-date:: " + $roamDate)
            #Write-Roam-File $rowPageNamePath ("csv-time:: " + $strTime)
            #Write-Roam-File $rowPageNamePath ("csv-filename:: " + $fileNameStr)
            #Write-Roam-File $rowPageNamePath ("csv-type:: " + $csvType)
        }
    }

    #Set a counter which will decide how spacing is done for indents in the Roam table structure
    #Start at 2 instead of 0 to account for CSV file name parent bullet and then {{table}} being second indent level, and everything needing to start indented under it
    $ctr = 2
    #For each row of csv file, loop through each column
    foreach($col in $row.psobject.properties.name)
    {
        $tableCellOrig = $row.$col
        $tableCell = $tableCellOrig
        $tableCell = $bulletType + $tableCell
        #Add the proper indentation based on looping through x number of times based on $ctr
        $tmpCtr = $ctr
        while($tmpCtr -gt 0)
        {
            $tableCell = $indentType + $tableCell
            $tmpCtr = $tmpCtr - 1
        }

        $arrTable += , $tableCell
        $ctr = $ctr + 1

        if($bPages -and $ctr -gt 3) #Need to skip the first column because that is what you are creating pages from
        {
            #Add attribute for the new page (row)
            if($bTesting){$col = "TESTING_" + $col}
            if($bInvalidChar)
            {
                #Add under each page name in summary as this is what we will do if a bad character for Windows in file name
                $arrSummary += , ($indentType + $indentType + $indentType + $col + ":: " + $tableCellOrig)
            }
            else
            {
                Write-Roam-File $rowPageNamePath ($col + ":: " + $tableCellOrig)
            }
        }
    }
}

Write-Roam-File $csvImportNamePath ($bulletType + "CSV Conversion Script created **$pgCtr Pages** and **$attrCtr Attributes**")
$arrLog += , ($indentType + $bulletType + "Merge the Summary into $csvImportName")

#Add Summary array to CSV-Import summary markdown file
Foreach($summRow in $arrSummary)
{
    Write-Roam-File $csvImportNamePath $summRow
}

$arrLog += , ($indentType + $bulletType + "Merge the Roam table markdown format into $csvImportName")

#Add the Roam table markdown format code from array to CSV-Import summary markdown file
Foreach($tableRow in $arrTable)
{
    Write-Roam-File $csvImportNamePath $tableRow
}

$arrLog += , ($indentType + $bulletType + "Merge the Logs from this CSV conversion/import into $csvImportName")

#Add the log array values to CSV-Import summary markdown file
Foreach($logRow in $arrLog)
{
    Write-Roam-File $csvImportNamePath $logRow
}

#Exit the script
Read-Host -Prompt "Script complete. Press any key to exit."
Exit