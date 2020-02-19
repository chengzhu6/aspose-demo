package SmartMarkers;

import java.util.ArrayList;

import AsposeCellsExamples.SmartMarkers.Individual;
import AsposeCellsExamples.SmartMarkers.Wife;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookDesigner;
import AsposeCellsExamples.Utils;

public class UsingNestedObjects {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UsingNestedObjects.class) + "SmartMarkers/";
		Workbook workbook = new Workbook(dataDir + "TestSmartMarkers.xlsx");

		WorkbookDesigner designer = new WorkbookDesigner();
		designer.setWorkbook(workbook);

		ArrayList<AsposeCellsExamples.SmartMarkers.Individual> list = new ArrayList<AsposeCellsExamples.SmartMarkers.Individual>();
		list.add(new AsposeCellsExamples.SmartMarkers.Individual("John", 23, new AsposeCellsExamples.SmartMarkers.Wife("Jill", 20)));
		list.add(new AsposeCellsExamples.SmartMarkers.Individual("Jack", 25, new AsposeCellsExamples.SmartMarkers.Wife("Hilly", 21)));
		list.add(new AsposeCellsExamples.SmartMarkers.Individual("James", 26, new AsposeCellsExamples.SmartMarkers.Wife("Hally", 22)));
		list.add(new Individual("Baptist", 27, new Wife("Newly", 23)));

		designer.setDataSource("Individual", list);

		designer.process(false);

		workbook.save(dataDir + "UsingNestedObjects-out.xlsx");
	}

}
