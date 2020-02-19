package AsposeCellsExamples.RowsAndColumns;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class DeleteARow {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(DeleteARow.class) + "RowsAndColumns/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Deleting 3rd row from the worksheet
		worksheet.getCells().deleteRows(2, 1, true);

		// Saving the modified Excel file in default (that is Excel 2000) format
		workbook.save(dataDir + "DeleteARow_out.xls");
	}
}
