package AsposeCellsExamples.RowsAndColumns;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class CopyingRows {

	public static void main(String[] args) throws Exception {
		
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyingRows.class) + "RowsAndColumns/";

		// Create a new Workbook.
		Workbook excelWorkbook = new Workbook(dataDir + "book1.xls");

		// Get the first worksheet in the workbook.
		Worksheet wsTemplate = excelWorkbook.getWorksheets().get(0);

		// Copy the second row with data, formating, images and drawing objects to the 12th row in the worksheet.
		wsTemplate.getCells().copyRow(wsTemplate.getCells(), 2, 10);

		// Save the excel file.
		excelWorkbook.save(dataDir + "CopyingRows_out.xls");

		// Print message
		System.out.println("Row and Column copied successfully.");
		
	}
}
