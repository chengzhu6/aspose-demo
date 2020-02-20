package WorkbookVBAProject;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class AddMarcoToWorkbook {

    public static void main(String[] args) throws Exception {
        Workbook workbook = new Workbook();

// Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

// Add VBA Module
        int idx = workbook.getVbaProject().getModules().add(worksheet);

// Access the VBA Module, set its name and codes
        VbaModule module = workbook.getVbaProject().getModules().get(idx);
        module.setName("TestModule");

        module.setCodes("Sub ShowMessage()" + "\r\n" + "    MsgBox \"Welcome to Aspose!\"" + "\r\n" + "End Sub");
        String dataDir = Utils.getSharedDataDir(AddMarcoToWorkbook.class) + "WorkbookVBAProject/";
// Save the workbook
        workbook.save( dataDir + "output.xlsm", SaveFormat.XLSM);
    }
}
