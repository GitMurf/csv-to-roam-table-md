#v0.1.0
#Code written by:   Murf
#Design by:         Rob Haisfield @RobertHaisfield on Twitter

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

#Loop through each row of the csv file
foreach($row in $csvObject)
{
    #For each row of csv file, loop through each column
    foreach($col in $row.psobject.properties.name)
    {
        $tableCell = $row.$col
        Add-content $newMarkdownFile -value $tableCell
    }
}

#Exit the script
Read-Host -Prompt "Script complete. Press any key to exit."
Exit