## AutoTableNH<img src="images/oslogo.jpg" width="23"/>
<p style="font-size:15px;">Project created on: September 19, 2024.</p>

## Description
This Office Script automatically creates a table with no headers. 

## Basic Instructions
1. Open any workbook in Excel for Windows or for Mac and select the Automate tab.

	<img src="/autotable/images/atinstruction1.png" width="550"/>
3. Click on New Script.

   	<img src="/autotable/images/atinstruction2.png" width="250"/>
5. On the left side, you will see the Code Editor pop up, click on Script 6 or whatever Script number populates, and rename the file in the Script name and it will save.
   
  	 <img src="/autotable/images/atinstruction3.png" width="250"/>
   
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
	workbook.addTable(currentRegion.getAddress(), false);
}
   }
   ```
	Con't should look like this.    
   	<img src="/autotable/images/atinstruction4.png" width="250"/>

5. Now click Run.
   
   	<img src="/autotable/images/atinstruction5.png" width="250"/>

	AutoTable Sample. 

  	 <img src="/autotable/images/atinstruction6.png" width="550"/>
