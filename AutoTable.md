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