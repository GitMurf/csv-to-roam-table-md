#v0.5
#Version Comments: Adding parent child bullet nesting
#Repository: https://github.com/GitMurf/csv-to-roam-table-md
#Code written by:       Murf @shawnpmurphy8 on Twitter
#Design/Concept by:     Rob Haisfield @RobertHaisfield on Twitter
#Design/Concept by:     EA @ec_anderson on Twitter

#If $bTesting = $true then add "TESTING_ to the front of any page created
$bTesting = $false

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
#NOTE: Whatever bullet type you use (other than if you use non and leave it empty) you should to add a space after it to match Roam export format
#NOTE: As of april 28, 2020 there is a Roam bug that creates issues with adding Attributes:: if not at Root level of page
    #To workaround this bug, will leave a space between the double colon "attr: : item" so it doesnt get imported as attribute and you just have to remove the space
#NOTE: Attribute fields canNOT have bullets added in front so just leave without a bullet... not sure if this has to do with Bug above or not
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

#This function checks if trailing comma and then removes it
Function Remove-Trailing-Comma
{
    Param(
        [string]$tmpJason
    )

    if($tmpJason.substring($tmpJason.length-1) -eq ",")
    {
        $tmpJason = $tmpJason.substring(0,$tmpJason.length - 1) #Remove trailing comma
    }
    return $tmpJason
}

#Add a blank line for easier reading of prompts in powershell window
Write-Host

#Exporting to JSON or Markdown (JSON is default as Markdown has limitations such as max 10 import at once and issues with Windows file name characters, attribute names etc.)
$bJSON = $true
$jsonString = "["

#Ask for user input on whether JSON or Markdown export
$respExport = Read-Host "Export to JSON (j) or Markdown (m)? Recommend JSON as Markdown has a 10 page import limit. Enter 'j' or 'm'"

#Check for JSON or Markdown
if($respExport -eq "m" -or $respExport -eq "'m'" -or $respExport -eq "M" -or $respExport -eq "'M'"){$bJSON = $false}else{$bJSON = $true}

Write-Host

#Importing attributes vs blocks
$bAttributes = $false

#Ask for user input on whether have one CSV row per page and then many columns that will be attributes OR many rows because adding blocks of data to each page.
$respPages = Read-Host "Are you adding attributes (one row per page with many columns) OR blocks of text (many rows per page, one 'Block' column)? (Enter 'a' for attributes, 'b' for blocks)"

#Check if user is importing attributes or blocks
if($respPages -eq "a" -or $respPages -eq "'a'" -or $respPages -eq "A" -or $respPages -eq "'A'"){$bAttributes = $true}else{$bAttributes = $false}

Write-Host

if($bAttributes)
{
    #Ask user for the type of csv import (e.g., People, Company, CRM etc.)
    $csvType = Read-Host "Enter the Type/Category of your CSV data (e.g., Contacts, Books, Videos, etc.) to allow for searching of similar data types in Roam"

    Write-Host
}
else
{
    if($bJSON -eq $false)
    {
        #Exit the script
        Read-Host -Prompt "Press any key to exit. Currently it is not Recommended to Import a CSV with Blocks (vs Attributes) with Markdown output. Restart the script and select the JSON option."
        Exit
    }
}

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

if($foundName -ne "")
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

if($bJSON)
{
    #Create JSON file
    $jsonFilePath = "$resultsFolder\" + "$csvImportName" + ".json"
}

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

if($bAttributes)
{
    #Write attribute for csv-import to first line of this new .md file (need to use LiteralPath parameter because of [[]] characters in path)
    Write-Roam-File $csvImportNamePath ("csv-date:: " + $roamDate)
    #Import time attribute
    Write-Roam-File $csvImportNamePath ("csv-time:: " + $strTime)
    #Filename attribute
    Write-Roam-File $csvImportNamePath ("csv-filename:: " + $fileNameStr)
    #Type of CSV file attribute (example could be: People, CRM, Company)
    Write-Roam-File $csvImportNamePath ("csv-type:: " + $csvType)
}

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
$arrSummary += , ($indentType + $bulletType + "ATTRIBUTES") #Will create links to each attribute created under this bullet

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

    #Need to add attributes to the summary page
    if($ctr -gt 3) #Need to skip the first column because that is what you are creating pages from
    {
        if($bTesting){$col = "TESTING_" + $col}
        $arrSummary += , ($indentType + $indentType + $bulletType + "#[[" + $col + "]]")
        $attrCtr = $attrCtr + 1
    }
}

if($ctr -gt 4)
{
    $checkLastCol = $col
    if($bTesting){$checkLastCol = $col.substring(8)}
    $bNestedBlocks = $false
    if($bAttributes -eq $false -and $checkLastCol -eq "Parent3"){$bNestedBlocks = $true}

    if($bAttributes -eq $false -and $bNestedBlocks -ne $true)
    {
        #Exit the script
        Read-Host -Prompt "Press any key to exit. You selected the 'Block' import as opposed to 'Attributes'. You can only have 2 columns with Block import (Page Name and Block). Please fix and restart the script."
        Exit
    }
}

#Create new page/file for each CSV row
$arrLog += , ($indentType + $bulletType + "Creating new Pages for each CSV row")
$arrSummary += , ($indentType + $bulletType + "PAGES CREATED")

$rowCtr = 1
$lastRowName = ""

$lastPar1 = ""
$lastPar2 = ""
$lastPar3 = ""

#Loop through each row of the csv file
foreach($row in $csvObject)
{
    write-host "Row: $rowCtr"
    #Create new page/file for each CSV row
    $colHeaderNames = $row.psobject.properties.name
    $rowPageName = $row.($colHeaderNames[0])
    if($bTesting){$rowPageName = "TESTING_" + $rowPageName}

    if($bJSON)
    {
        if($lastRowName -ne $rowPageName)
        {
            if($bAttributes -eq $false -and $rowCtr -gt 1)
            {
                $jsonString = Remove-Trailing-Comma $jsonString #Remove trailing comma
                #Close par3 if there previously was one
                if($lastPar3 -ne ""){$jsonString = $jsonString + ']}'}
                #Close par2 if there previously was one
                if($lastPar2 -ne ""){$jsonString = $jsonString + ']}'}
                #Close par1 if there previously was one
                if($lastPar1 -ne ""){$jsonString = $jsonString + ']}'}
                #Close title for previous page
                $jsonString = $jsonString + ']},'
            }
            $jsonString = $jsonString + '{"title":"' + $rowPageName + '","children":['
        }
    }

    #Dont add to the array if importing "blocks" as for big files like the Bible (31k blocks) it slows the script way down as the array grows
    if($bAttributes){$arrSummary += , ($indentType + $indentType + $bulletType + "[[" + $rowPageName + "]]")}
    $rowPageNamePath = "$resultsFolder\" + "$rowPageName" + ".md"
    $pgCtr = $pgCtr + 1

    #Commenting out the CSV import attribute data becuase isn't needed on each page... instead link to the csv summary page which has all that info
    #Check if any of the Windows filename illegal characters are present and if so, do NOT write to the file and instead just store the attributes on the summary page
        #Then the user can go into Summary page in Roam and copy the attributes, click the page name and then add there so that can keep the special character in name
        #The characters not allowed are: \ / : * ? " < > |
    $bInvalidChar = $false
    if($bJSON -eq $false)
    {
        if($bAttributes)
        {
            if($rowPageName.Contains("\") -or $rowPageName.Contains("/") -or $rowPageName.Contains(":") -or $rowPageName.Contains("*") -or $rowPageName.Contains("?") -or $rowPageName.Contains('"') -or $rowPageName.Contains("<") -or $rowPageName.Contains(">") -or $rowPageName.Contains("|"))
            {
                $bInvalidChar = $true
                $arrLog += , ($indentType + $bulletType + "**Invalid character** for Windows found in Filename for PAGE: [[" + $rowPageName + "]]")
            }
            else{$arrLog += , ($indentType + $bulletType + "Created the Page: [[" + $rowPageName + "]]")}
        }
    }

    if($bInvalidChar)
    {
        #Add under each page name in summary as this is what we will do if a bad character for Windows in file name
        #NOTE: As of april 28, 2020 there is a Roam bug that creates issues with adding Attributes:: if not at Root level of page
            #To workaround this bug, will leave a space between the double colon so it doesnt get imported as attribute and you just have to remove the space
        #Dont add to the array if importing "blocks" as for big files like the Bible (31k blocks) it slows the script way down as the array grows
        if($bAttributes){$arrSummary += , ($indentType + $indentType + $indentType + "csv-import: : [[" + $csvImportName + "]]")}
        #General attributes for the CSV import. These are in the Summary page for the import so do we need them also on every page?
        #$arrSummary += , ($indentType + $indentType + $indentType + $bulletType + "csv-date:: " + $roamDate)
        #$arrSummary += , ($indentType + $indentType + $indentType + $bulletType + "csv-time:: " + $strTime)
        #$arrSummary += , ($indentType + $indentType + $indentType + $bulletType + "csv-filename:: " + $fileNameStr)
        #$arrSummary += , ($indentType + $indentType + $indentType + $bulletType + "csv-type:: " + $csvType)
    }
    else
    {
        if($bJSON -ne $true){Write-Roam-File $rowPageNamePath ("csv-import:: [[" + $csvImportName + "]]")}
        #General attributes for the CSV import. These are in the Summary page for the import so do we need them also on every page?
        #Write-Roam-File $rowPageNamePath ("csv-date:: " + $roamDate)
        #Write-Roam-File $rowPageNamePath ("csv-time:: " + $strTime)
        #Write-Roam-File $rowPageNamePath ("csv-filename:: " + $fileNameStr)
        #Write-Roam-File $rowPageNamePath ("csv-type:: " + $csvType)
    }

    #Set a counter which will decide how spacing is done for indents in the Roam table structure
    #Start at 2 instead of 0 to account for CSV file name parent bullet and then {{table}} being second indent level, and everything needing to start indented under it
    $ctr = 2
    #For each row of csv file, loop through each column
    $listOfColumns = $row.psobject.properties.name
    foreach($col in $listOfColumns)
    {
        $tableCellOrig = $row.$col
        if($bAttributes)
        {
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
        }

        $ctr = $ctr + 1
        if($ctr -gt 3) #Need to skip the first column because that is what you are creating pages from
        {
            #Add attribute for the new page (row)
            if($bTesting){$col = "TESTING_" + $col}
            if($bInvalidChar)
            {
                #Add under each page name in summary as this is what we will do if a bad character for Windows in file name
                #NOTE: As of april 28, 2020 there is a Roam bug that creates issues with adding Attributes:: if not at Root level of page
                    #To workaround this bug, will leave a space between the double colon so it doesnt get imported as attribute and you just have to remove the space
                if($bAttributes){$arrSummary += , ($indentType + $indentType + $indentType + $col + ": : " + $tableCellOrig)}
            }
            else
            {
                if($bJSON -ne $true)
                {
                    Write-Roam-File $rowPageNamePath ($col + ":: " + $tableCellOrig)
                }
                else
                {
                    if($bAttributes -eq $false)
                    {
                        #For block import only need to look at the one column next to the page name column
                        if($ctr -eq 4)
                        {
                            if($bNestedBlocks)
                            {
                                $par1 = $row.($listOfColumns[$ctr-2])
                                $par2 = $row.($listOfColumns[$ctr-2+1])
                                $par3 = $row.($listOfColumns[$ctr-2+2])

                                if($lastRowName -ne $rowPageName) #New set of rows / page name section in CSV
                                {
                                    if($par1 -ne "")
                                    {
                                        $jsonString = $jsonString + '{"string":"' + $par1 + '","children":['
                                        if($par2 -ne "")
                                        {
                                            $jsonString = $jsonString + '{"string":"' + $par2 + '","children":['
                                            if($par3 -ne ""){$jsonString = $jsonString + '{"string":"' + $par3 + '","children":['}
                                        }
                                    }

                                    $lastPar1 = $par1
                                    $lastPar2 = $par2
                                    $lastPar3 = $par3
                                }

                                #par1 changed which means changing all 3 parents
                                if($par1 -ne $lastPar1)
                                {
                                    $jsonString = Remove-Trailing-Comma $jsonString #Remove trailing comma
                                    #Close par3 if there previously was one
                                    if($lastPar3 -ne ""){$jsonString = $jsonString + ']}'}
                                    #Close par2 if there previously was one
                                    if($lastPar2 -ne ""){$jsonString = $jsonString + ']}'}
                                    #Close par1 if there previously was one
                                    if($lastPar1 -ne ""){$jsonString = $jsonString + ']}'}

                                    #If not blank then add new par1
                                    if($par1 -ne "")
                                    {
                                        $jsonString = $jsonString + ',{"string":"' + $par1 + '","children":['
                                        #If not blank then add new par2
                                        if($par2 -ne "")
                                        {
                                            $jsonString = $jsonString + '{"string":"' + $par2 + '","children":['
                                            #If not blank then add new par3
                                            if($par3 -ne "")
                                            {
                                                $jsonString = $jsonString + '{"string":"' + $par3 + '","children":['
                                            }
                                            else{$jsonString = $jsonString + ','} #Have to add the trailing comma again
                                        }
                                        else{$jsonString = $jsonString + ','} #Have to add the trailing comma again
                                    }
                                    else{$jsonString = $jsonString + ','} #Have to add the trailing comma again
                                }
                                elseif($par2 -ne $lastPar2) #If par2 changed then change par2 and par3
                                {
                                    $jsonString = Remove-Trailing-Comma $jsonString #Remove trailing comma
                                    #Close par3 if there previously was one
                                    if($lastPar3 -ne ""){$jsonString = $jsonString + ']}'}
                                    #Close par2 if there previously was one
                                    if($lastPar2 -ne ""){$jsonString = $jsonString + ']}'}

                                    #If not blank then add new par2
                                    if($par2 -ne "")
                                    {
                                        $jsonString = $jsonString + ',{"string":"' + $par2 + '","children":['
                                        #If not blank then add new par3
                                        if($par3 -ne "")
                                        {
                                            $jsonString = $jsonString + '{"string":"' + $par3 + '","children":['
                                        }
                                        else{$jsonString = $jsonString + ','} #Have to add the trailing comma again
                                    }
                                    else{$jsonString = $jsonString + ','} #Have to add the trailing comma again
                                }
                                elseif($par3 -ne $lastPar3) #If only par3 changed
                                {
                                    $jsonString = Remove-Trailing-Comma $jsonString #Remove trailing comma
                                    #Close par3 if there previously was one
                                    if($lastPar3 -ne ""){$jsonString = $jsonString + ']}'}

                                    #If not blank then add new par3
                                    if($par3 -ne "")
                                    {
                                        $jsonString = $jsonString + ',{"string":"' + $par3 + '","children":['
                                    }
                                    else{$jsonString = $jsonString + ','} #Have to add the trailing comma again
                                }

                                $lastPar1 = $par1
                                $lastPar2 = $par2
                                $lastPar3 = $par3
                            }

                            $jsonString = $jsonString + '{"string":"' + $tableCellOrig + '"},'
                        }
                    }
                    else
                    {
                        $jsonString = $jsonString + '{"string":"' + $col + ':: ' + $tableCellOrig + '"},'
                    }
                }
            }
        }
    }
    #Account for the extra "," added in last children item
    if($bJSON)
    {
        if($bAttributes)
        {
            $jsonString = Remove-Trailing-Comma $jsonString #Remove trailing comma
            $jsonString = $jsonString + ']},'
        }
    }
    $lastRowName = $rowPageName
    $rowCtr = $rowCtr + 1
}

#Close out the final children and title ]}
if($bJSON)
{
    #If importing block style CSV file then there is an extra set of [] that needs to be closed before the final ] closure
    if($bAttributes -eq $false)
    {
        $jsonString = Remove-Trailing-Comma $jsonString #Remove trailing comma
        #Close par3 if there previously was one
        if($lastPar3 -ne ""){$jsonString = $jsonString + ']}'}
        #Close par2 if there previously was one
        if($lastPar2 -ne ""){$jsonString = $jsonString + ']}'}
        #Close par1 if there previously was one
        if($lastPar1 -ne ""){$jsonString = $jsonString + ']}'}
        #Close title for previous page
        $jsonString = $jsonString + ']},'
    }
    $jsonString = Remove-Trailing-Comma $jsonString #Remove trailing comma
    $jsonString = $jsonString + ']'
    Write-Roam-File $jsonFilePath $jsonString
}
else
{
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
}

#Exit the script
Read-Host -Prompt "Script complete. Press any key to exit."
Exit
