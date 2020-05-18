#Get path where script is running from so you can target JSON
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

#Try to find the JSON file to convert automatically by looking in current folder and also sorting by most recent edited
$foundJson = Get-ChildItem -Path $scriptPath -Filter *.json | Sort-Object -Property LastWriteTime -Descending | Select-Object -first 1
$foundName = $foundJson.Name

if(!$foundJson)
{
    #Exit the script
    Read-Host -Prompt "Press any key to exit. No JSON files found in the folder you are running the script from. Copy to the same folder and restart the script."
    Exit
}

#The JSON that was auto picked is the correct one (JSON in same folder as script)
#Get the file name and create the full path
$fileNameStr = $foundName
$fileNameStrPath = $foundJson.FullName

#Import JSON file into a Variable to loop through and parse
$jsonObj = ConvertFrom-JSON (Get-Content $fileNameStrPath -Raw)

#Set the Results folder to store all the outputs from the script
$resultsFolder = "$scriptPath\Results"

#Get a date string down to the second we can add to the JSON result file we will be creating so no duplicates if we run multiple times
$fullDateStr = get-date
$dateStrName = $fullDateStr.ToString("yyyyMMdd_HHmmss")
$csvResultFileName = "Roam_Json_to_CSV_" + $fileNameStr + "_" + $dateStrName
$jsonResultCsv = "$resultsFolder\" + "$csvResultFileName" + ".csv"

#Create Results folder if it doesn't already exist
if(!(Test-Path $resultsFolder)){New-Item -ItemType Directory -Force -Path $resultsFolder | Out-Null}

#$true on end for append instead of overwrite
$csvResultStream = New-Object System.IO.StreamWriter -ArgumentList "$jsonResultCsv",$true

#Clean stuff out like '"' quotes that have to be escaped with 2 of them
function Clean-String($InputString)
{
    if($InputString -eq ""){return $InputString}

    #Replace special characters like line breaks and tabs
    $InputString = $InputString -Replace '"','""' -Replace "`t","\t" -Replace "`r","\r" -Replace "`n","\n"
    return $InputString
}

Function Write-To-Result
{
    Param(
        [string]$resultLine
    )

    $csvResultStream.WriteLine($resultLine)
}

Function Loop-Block
{
    Param(
        [object]$pgBlock,
        [string]$hierStr,
        [string]$pgName,
        [int]$blDepth
    )

    foreach($childBlock in $pgBlock.children)
    {
        $newBlDepth = $blDepth + 1
        $blockStr = $childBlock.string
        $blockUID = $childBlock.uid
        $strResult = "$blockUID : $blockStr"
        $newHierStr = $hierStr + ' > ' + $strResult
        $csvString = '"' + (Clean-String $pgName) + '","' + (Clean-String $blockUID) + '","' + $newBlDepth + '","' + (Clean-String $blockStr) + '","' + (Clean-String $newHierStr) + '"'
        Write-To-Result $csvString
        Loop-Block $childBlock $newHierStr $pgName $newBlDepth
    }
}

#Creater header row for columns
$csvString = '"' + "PageName" + '","' + "BlockUID" + '","' + "BlockDepth" + '","' + "BlockString" + '","' + "Hierarchy" + '"'
Write-To-Result $csvString

#Loop through every Roam Page Name
foreach($pageObj in $jsonObj)
{
    $pageName = $pageObj.title
    #Loop through every block on each page
    foreach($pageBlock in $pageObj.children)
    {
        $blockDepth = 0
        $blockStr = $pageBlock.string
        $blockUID = $pageBlock.uid
        $strResult = "$blockUID : $blockStr"
        $hierarchyStr = $pageName + ' > ' + $strResult
        $csvString = '"' + (Clean-String $pageName) + '","' + (Clean-String $blockUID) + '","' + $blockDepth + '","' + (Clean-String $blockStr) + '","' + (Clean-String $hierarchyStr) + '"'
        Write-To-Result $csvString
        Loop-Block $pageBlock $hierarchyStr $pageName $blockDepth
    }
    break
}

$csvResultStream.Close()

#Exit the script
Read-Host -Prompt "Conversion complete. Press any key to exit."
Exit
