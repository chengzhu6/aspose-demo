package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class PreventExportingHiddenWorksheetContent {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(PreventExportingHiddenWorksheetContent.class) + "TechnicalArticles/";
		
		// Create workbook object
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Do not export hidden worksheet contents
		ImplementingIStreamProvider options = new ImplementingIStreamProvider();
		options.setExportHiddenWorksheet(false);

		// Save the workbook
		workbook.save(dataDir + "PEHWorksheetContent_out.html");

	}
}
