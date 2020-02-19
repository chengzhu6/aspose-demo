package AsposeCellsExamples.RowsAndColumns;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class HidingRowsandColumns {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(HidingRowsandColumns.class) + "RowsAndColumns/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		// Hiding the 3rd row of the worksheet
		cells.hideRow(2);

		// Hiding the 2nd column of the worksheet
		cells.hideColumn(1);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "HidingRowsandColumns_out.xls");

		// Print message
		System.out.println("Rows and Columns hidden successfully.");

	}
}
