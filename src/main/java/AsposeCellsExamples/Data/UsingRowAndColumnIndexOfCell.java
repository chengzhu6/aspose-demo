package AsposeCellsExamples.Data;

import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class UsingRowAndColumnIndexOfCell {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UsingRowAndColumnIndexOfCell.class) + "Data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the worksheet in the Excel file
		com.aspose.cells.Worksheet worksheet = workbook.getWorksheets().get(0);
		com.aspose.cells.Cells cells = worksheet.getCells();

		// Accessing a cell using the indices of its row and column
		com.aspose.cells.Cell cell = cells.get(0, 0);

		// Print message
		System.out.println("Cell Value: " + cell.getValue());

	}
}
