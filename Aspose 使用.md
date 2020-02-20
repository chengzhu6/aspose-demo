# Aspose.cell for java

##support java 1.6,1.7,1.8
##论坛地址：https://forum.aspose.com/c/cells
##Demo repo: https://github.com/chengzhu6/aspose-demo.git

## Start with gradle
    在build.gradle中引入
        repositories {
            maven{
                url "http://repository.aspose.com/repo/"
            }
        dependencies {
            compile group: 'com.aspose', name: 'aspose-cells', version: '20.1'
        }
##  使用License
    
    License license = new License();
    license.setLicense("Aspose.Cells.Java.lic");
     
    if (License.isLicenseSet()) {
        System.out.println("License is Set!");
    } 
## Feature
    
   ###1. 在worksheet添加水印
    
        String dataDir = Utils.getSharedDataDir(AddWordArtWatermarkToWorksheet.class) + "TechnicalArticles/";
        // Instantiate a new Workbook
        Workbook workbook = new Workbook();
        
        // Get the first default sheet
        Worksheet sheet = workbook.getWorksheets().get(0);
        
        // Add Watermark
        Shape wordart = sheet.getShapes().addTextEffect(MsoPresetTextEffect.TEXT_EFFECT_1, "CONFIDENTIAL",
        		"Arial Black", 50, false, true, 18, 8, 1, 1, 130, 800);
        
        // Get the fill format of the word art
        FillFormat wordArtFormat = wordart.getFill();
        
        // Set the color
        wordArtFormat.setOneColorGradient(Color.getRed(), 0.2, GradientStyleType.HORIZONTAL, 2);
        
        // Set the transparency
        wordArtFormat.setTransparency(0.9);
        
        // Make the line invisible
        LineFormat lineFormat = wordart.getLine();
        lineFormat.setWeight(0.0);
        
        // Save the file
        workbook.save(dataDir + "AWArtWToWorksheet_out.xls");
   ###2. 添加宏from一个模版中，修改现有的宏, 添加宏。
        
   ####添加宏
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
   
获取到的collection不支持forEach, stream。
