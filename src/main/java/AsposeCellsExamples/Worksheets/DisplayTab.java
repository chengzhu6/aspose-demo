package AsposeCellsExamples.Worksheets;

import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class DisplayTab {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DisplayTab.class) + "Worksheets/";

		// Instantiating a Workbook object by excel file path
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Hiding the tabs of the Excel file
		workbook.getSettings().setShowTabs(true);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "DisplayTab_out.xls");

		// Print message
		System.out.println("Tabs are now displayed, please check the output file.");

	}
}
