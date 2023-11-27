package org.example;

import com.grapecity.documents.excel.*;

//该示例，展示如何通过页面分隔符（HPageBreaks）把页面进行分割。
//同时，使用打印设置中的标题行，来保证每页有相同的内容。
//运行，并查看PDF所展示的结果。
//页面分隔符：https://www.grapecity.com.cn/developer/grapecitydocuments/excel-java/docs/Features/ConfigurePrintSettingsviaPageSetup/ConfigurePageBreaks
//顶部及底部重复行：https://www.grapecity.com.cn/developer/grapecitydocuments/excel-java/docs/Features/ConfigurePrintSettingsviaPageSetup/ConfigureRowstoRepeatatTopandBottom
public class Example2_SplitDataToEachSheet {
    public static void main(String[] args) {
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getWorksheets().get(0);
        Object data = new Object[][]{
                {"A", "A1", "A2"},
                {"A", "A1", "A2"},
                {"A", "A1", "A2"},
                {"A", "A1", "A2"},
                {"B", "B1", "B2"},
                {"B", "B1", "B2"},
                {"B", "B1", "B2"},
                {"B", "B1", "B2"},
                {"B", "B1", "B2"},
                {"C", "C1", "C2"},
                {"C", "C1", "C2"},
                {"C", "C1", "C2"},
                {"C", "C1", "C2"},
        };

        sheet.getRange("A1:C1").merge();
        sheet.getRange("A1:C1").setValue("公司信息");
        sheet.getRange("A2").setValue("公司名");
        sheet.getRange("B2").setValue("数据1");
        sheet.getRange("C2").setValue("数据2");
        sheet.getRange("A3:C15").setValue(data);

        sheet.getHPageBreaks().add(sheet.getRange("A7"));
        sheet.getHPageBreaks().add(sheet.getRange("A12"));

        sheet.getPageSetup().setPrintTitleRows("$1:$2");

        wb.save("output/SplitData.pdf");
    }
}
