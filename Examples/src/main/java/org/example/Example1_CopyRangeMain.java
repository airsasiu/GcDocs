package org.example;

import com.grapecity.documents.excel.*;

//这个例子展示了如何跨工作表复制区域
//跨工作表剪切或复制：https://www.grapecity.com.cn/developer/grapecitydocuments/excel-java/docs/Features/ManageWorkbook/CutOrCopyAcrossSheets
public class Example1_CopyRangeMain {
    public static void main(String[] args) {
        Workbook wb = new Workbook();
        wb.getWorksheets().add();

        IWorksheet sheet1 = wb.getWorksheets().get(0);
        IWorksheet sheet2 = wb.getWorksheets().get(1);
        //在Sheet1中添加测试数据
        Object[][] data = new Object[][] { { 1 }, { 3 }, { 5 }, { 7 }, { 9 } };
        sheet1.getRange("A1:A5").setValue(data);

        //把数据从sheet1上复制到sheet2中
        sheet1.getRange("A1:A5").copy(sheet2.getRange("A1:A5"));
        //或者使用 剪切
        //sheet1.getRange("A1:A5").cut(sheet2.getRange("A1:A5"));

        //保存为Excel
        wb.save("output/CopyRange.xlsx");
    }
}