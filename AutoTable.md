---
Name: AutoTable
Description: Office Script that automatically creates a table
Date: 09/15/2024
---
<img src="Images/OSLogo.jpg" width="23"/>
## Why AutoTable?
If you are contanstly working with lenghty excel reports and need to create tables, this office script will saves you time. 


## Basic instructions

1. Open any workbook in Excel for Windows or for Mac and select the Automate tab.
2. Click on New Script.
3. On the left side, you will see the Code Editor pop up, click on Script 6 or whatever number populates, and rename the file.

## Sample files

Download <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true">convert-csv-example.zip</a> to get the Template.xlsx file and two sample .csv files. Extract the files into a folder in your OneDrive. This sample assumes the folder is named "output".

Add the following script and build a flow using the steps given to try the sample yourself!

## Sample code: Insert comma-separated values into a workbook

```TypeScrip
function main(workbook: ExcelScript.Workbook) {
	// Get the active worksheet in the current workbook
	let selectedSheet = workbook.getActiveWorksheet();
	
	// Select the specific cell B2 in the active worksheet
	let specificCell = selectedSheet.getRange("B2");

	// Get the surrouding region around the selected cell B2
	let currentRegion = specificCell.getSurroundingRegion();

	// Convert the surrounding region into a table, using the first row as headers
	selectedSheet.addTable(currentRegion.getAddress(),true);
}
