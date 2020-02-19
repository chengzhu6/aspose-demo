package HTML;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ExcludeUnusedStylesInExcelToHTML {

	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	/**
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		//Create workbook
		Workbook wb = new Workbook();

		//Create an unused named style
		wb.createStyle().setName("UnusedStyle_XXXXXXXXXXXXXX");

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		wb.getWorksheets().add("mysheet");
		Worksheet sheet1 = wb.getWorksheets().get(1);
		sheet1.setName("my sheet");
		sheet1.getCells().get("C7").putValue("This is my cell");

		//Put some value in cell C7
		ws.getCells().get("C7").putValue("This is sample text.");
		ws.getCells().get("C8").putValue("This is C8 text.");

		//Specify html save options, we want to exclude unused styles
		HtmlSaveOptions opts = new HtmlSaveOptions();

		//Comment this line to include unused styles
		opts.setExcludeUnusedStyles(true);

		//Save the workbook in html format
		wb.save(outDir + "outputExcludeUnusedStylesInExcelToHTML.html", opts);

		// Print the message
		System.out.println("ExcludeUnusedStylesInExcelToHTML executed successfully.");
	}
}
