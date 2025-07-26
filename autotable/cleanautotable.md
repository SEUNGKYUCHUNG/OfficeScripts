## CleanAutoTable<img src="images/oslogo.gif" alt="OS Logo" width="20" height="20"> 
<p style="font-size:15px;">Project created on: May 21, 2025.</p>

## Description
This Office Script automatically Drops blank rows/columns & creates table with headers. 

## Basic Instructions
1. Open any workbook in Excel for Windows or for Mac and select the Automate tab.

	<img src="/autotable/images/atinstruction1.png" width="550"/>
2. Click on New Script.

   	<img src="/autotable/images/atinstruction2.png" width="250"/>
3. On the left side, you will see the Code Editor pop up, click on Script 8 or whatever Script number populates, and rename the file in the Script name and it will save.
   
  	 <img src="/autotable/images/atinstruction3b.png" width="250"/>
   
4. In the Code Editor, copy this code, paste it, and **Save script**.

```TypeScript
/**
 * Cleans the active worksheet by removing empty columns and rows,
 * then converts the remaining block into a table.  
 * Prompts you to choose at runtime whether the table has headers.
 *
 * @param workbook      The Excel workbook.
 * @param hasHeaders    Checkbox shown on Run: 
 *                      - true  = treat first row as headers  
 *                      - false = generate default headers (Column1, Column2, …)
 */
function main(
    workbook: ExcelScript.Workbook,
    hasHeaders: boolean = true  // ← Runtime checkbox: “hasHeaders”
) {
    // STEP 1: Get the active worksheet and its “used” block
    const sheet = workbook.getActiveWorksheet();
    let range = sheet.getUsedRange();
    if (!range) {
       // If the sheet is entirely blank, exit early.
        console.log("No data found on the sheet.");
        return;
    }

    // STEP 2: Delete empty columns, scanning RIGHT → LEFT
    // 2.1) Read all values into a 2D array to avoid in-loop fetches
    let data = range.getValues() as (string | null)[][];
    // 2.2) Loop from the last column index down to 0
    for (let c = data[0].length - 1; c >= 0; c--) {
        let isEmpty = true;
        // 2.3) For each row in column c, check if any cell is non-blank
        for (let r = 0; r < data.length; r++) {
            if (data[r][c] !== "" && data[r][c] != null) {
                isEmpty = false;
                break;
            }
        }
        // 2.4) If the entire column is blank, delete it and shift left
        if (isEmpty) {
            range.getColumn(c).delete(ExcelScript.DeleteShiftDirection.left);
        }
    }

    // STEP 3: Delete empty rows, scanning BOTTOM → TOP
    // 3.1) Refresh range & data after columns have been deleted
    range = sheet.getUsedRange();
    data = range.getValues() as (string | null)[][];
    // 3.2) Loop from the last row index down to 0
    for (let r = data.length - 1; r >= 0; r--) {
        let isEmpty = true;
        // 3.3) For each column in row r, check if any cell is non-blank
        for (let c = 0; c < data[0].length; c++) {
            if (data[r][c] !== "" && data[r][c] != null) {
                isEmpty = false;
                break;
            }
        }
        // 3.4) If the entire row is blank, delete it and shift up
        if (isEmpty) {
            range.getRow(r).delete(ExcelScript.DeleteShiftDirection.up);
        }
    }

    // STEP 4: Convert the cleaned block into a table
    range = sheet.getUsedRange();
    if (range) {
        // Use the user’s choice to decide if the first row is treated as headers
        sheet.addTable(range.getAddress(), /*hasHeaders=*/ hasHeaders);
    }
}
```
Con't should look like this.
        
<img src="/autotable/images/atinstruction8b.png" width="250"/> 
The entire coding doesn't show unless you scroll down. Make sure the entire code is all there. 


5. Now click Run.
   
   	<img src="/autotable/images/atinstruction9b.png" width="250"/>

6. Select True for Header or False for no header. Then, click run. 

   	 <img src="/autotable/images/atinstruction9c.png" width="550"/>


	CleanAutoTable Sample.

   Before:

   <img src="/autotable/images/atinstruction11.png" width="550"/>

    After: 

    <img src="/autotable/images/atinstruction6.png" width="550"/>
