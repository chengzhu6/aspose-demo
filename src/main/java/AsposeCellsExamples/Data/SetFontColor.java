package AsposeCellsExamples.Data;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class SetFontColor {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SetFontColor.class) + "Data/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the added worksheet in the Excel file
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
		Cells cells = worksheet.getCells();

		// Adding some value to the "A1" cell
		Cell cell = cells.get("A1");
		cell.setValue("Hello Aspose!");

		// Setting the font color to blue
		Style style = cell.getStyle();
		Font font = style.getFont();
		font.setColor(Color.getBlue());
		cell.setStyle(style);

		cell.setStyle(style);

		// Saving the modified Excel file in default format
		workbook.save(dataDir + "SetFontColor_out.xls");
	}
}
