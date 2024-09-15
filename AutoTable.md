---
Name: AutoTable
Description: Office Script that automatically creates a table
Date: 09/15/2024
---

# Convert CSV files to Excel workbooks

Many services export data as comma-separated value (CSV) files. This solution automates the process of converting those CSV files to Excel workbooks in the .xlsx file format. It uses a [Power Automate](https://make.powerautomate.com) flow to find files with the .csv extension in a OneDrive folder and an Office Script to copy the data from the .csv file into a new Excel workbook.

## Solution

1. Store the .csv files and a blank "Template" .xlsx file in a OneDrive folder.
1. Create an Office Script to parse the CSV data into a range.
1. Create a Power Automate flow to read the .csv files and pass their contents to the script.

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
