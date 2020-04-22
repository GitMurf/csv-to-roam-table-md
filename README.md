#Converting CSV files to Markdown for Roam Tables

Very early stages. Script is working great but non-windows folks have to figure out how to run PowerShell on their computer.

*Rob Haisfield is on a MAC and said he found a source for running PowerShell. Will wait and see when he tests it, but hopefully he can explain how it can be done.*

---

## Brief Video Demo

https://www.screencast.com/t/ak2DEaml2HYK

---

## How to Download the Script

I know everyone is used to GitHub, but the nice thing about Bitbucket is you do NOT have to signup for an account to access a Public repository like mine.

1. Go to [Downloads tab](https://bitbucket.org/murf/csv-to-roam-table-md/downloads/)
2. Click **Download repository**
3. Save the .zip package on your computer.
4. Unzip the package.

---

## How to Run the Script

This is a brand new script and I haven't spent much time on it but wanted to get it out. The ReadMe is lacking as well but I plan to spend some time later this week. Watch the Demo Video linked above which shows the script in action.

1. Move the CSV-to-Roam-table-md.ps1 PowerShell script into the same folder as whatever .CSV file you are trying to convert.
2. Right-click the script and select **Run with PowerShell**.
3. Press ENTER if you are converting a file with the Default "," (comma) separated values.
    * If you want to use another delimiter like ; or | or TAB then press 'n' and ENTER
    * Enter the delimiter you want to use. NOTE: If your file is Tab delimited then enter 'TAB'.
4. Press ENTER again at the next prompt if you want to use the Default path which is set to whatever folder you are running the script from.
    * If you want to use another folder path then press 'n' and ENTER
    * Enter the folder path where your .CSV file is that you want to convert.
5. Enter the name of the file... do **NOT** include the path (For Example: locations.csv)
6. Open the resulting .md file and simply copy the entire file contents and then paste into any block in Roam and it should auto format/create the table.

A lot more to come!