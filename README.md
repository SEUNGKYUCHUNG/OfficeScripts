## AutoTable<img src="Images/OSLogo.jpg" width="23"/>9.15.2024





## Description
This Office Script automatically creates a table for you. 

## Basic Instructions
1. Open any workbook in Excel for Windows or for Mac and select the Automate tab.

    <img src="/atinstruction1.jpg" width="550"/>
3. Click on New Script.
   
   <img src="/atinstruction2.jpg.png" width="200"/>
5. On the left side, you will see the Code Editor pop up, click on Script 6 or whatever number populates, and rename the file and it will save.
   
   <img src="/atinstruction3.png" width="300"/>
7. In the Code Editor, copy this code, paste it, and Save script.
   ```TypeScript
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
   ```
Con't should look like this.    
<img src="/atinstruction4.png" width="270"/>

5. Now click Run.
   
   <img src="/atinstruction5.png" width="250"/>
