package AsposeCellsExamples.Worksheets;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class DisplayHideScrollBars {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DisplayHideScrollBars.class) + "Worksheets/";

		// Instantiating a Excel object by excel file path
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Hiding the vertical scroll bar of the Excel file
		workbook.getSettings().setVScrollBarVisible(false);

		// Hiding the horizontal scroll bar of the Excel file
		workbook.getSettings().setHScrollBarVisible(false);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "DisplayHideScrollBars_out.xls");

		// Print message
		System.out.println("Scroll bars are now hidden, please check the output document.");

	}
}
