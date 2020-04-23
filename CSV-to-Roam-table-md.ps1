#v0.2.1
#Version Comments: First working version to be tested
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

#Ask for user input. If left blank and user presses ENTER, then continue with default (comma). Otherwise they can enter their own option.
$respDelim = Read-Host "Default CSV Delimiter is '$strDelim' (comma). Press ENTER to Continue or input 'n' to change it."

#Add a blank line for easier reading of prompts in powershell window
Write-Host

#Check if user decided to change to a different delimiter
if($respDelim -eq "n" -or $respDelim -eq "N" -or $respDelim -eq "'n'" -or $respDelim -eq "'N'")
{
    $strDelim = Read-Host "What would you like your CSV Delimiter to be? (If Tab delimited, enter 'TAB')"
    if($strDelim -eq "TAB" -or $strDelim -eq "tab" -or $strDelim -eq "'TAB'" -or $strDelim -eq "'tab'"){$strDelim = "`t"}
}

#Add a blank line for easier reading of prompts in powershell window
Write-Host

#Get path where script is running from so you can target CSV
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

#Ask for user input. If left blank and user presses ENTER, then use the script path. Otherwise they can enter their own custom path.
$respPath = Read-Host "Is your target CSV file located here: '$scriptPath'? Press ENTER to Continue or input 'n' to change it."

#Add a blank line for easier reading of prompts in powershell window
Write-Host

#Check if user decided to change to a different path
if($respPath -eq "n" -or $respPath -eq "N" -or $respPath -eq "'n'" -or $respPath -eq "'N'")
{
    $scriptPath = Read-Host "Enter the folder path of your CSV file (do NOT include the file name)"
}

#Add a blank line for easier reading of prompts in powershell window
Write-Host

#Get the file name from user and create the full path
$fileNameStr = Read-Host "Name of CSV file with extension? Do NOT include path. Example: FileName.csv"
$fileNameStrPath = $scriptPath + "\" + $fileNameStr

#Add a blank line for easier reading of prompts in powershell window
Write-Host

#Get a date string down to the second we can add to new markdown file we will be creating so no duplicates if we run multiple times
$fullDateStr = get-date
$dateStrName = $fullDateStr.ToString("yyyy_MM_dd-HH_mm_ss")
$newMarkdownFile = "$scriptPath\" + "$fileNameStr" + "_$dateStrName.md"

#Import .CSV file into a Variable to loop through and parse
$csvObject = Import-Csv -Delimiter $strDelim -Path "$fileNameStrPath"

#Collapse the entire table under a parent bullet with name of the CSV file
$tableCell = "TABLE IMPORT FROM CSV: " + $fileNameStr
$tableCell = $bulletType + $tableCell
Add-content $newMarkdownFile -value $tableCell

#Add {{table}}
$tableCell = "{{table}}"
$tableCell = $bulletType + $tableCell
$tableCell = $indentType + $tableCell
Add-content $newMarkdownFile -value $tableCell

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

    Add-content $newMarkdownFile -value $tableCell
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

        Add-content $newMarkdownFile -value $tableCell
        $ctr = $ctr + 1
    }
}

#Exit the script
Read-Host -Prompt "Script complete. Press any key to exit."
Exit