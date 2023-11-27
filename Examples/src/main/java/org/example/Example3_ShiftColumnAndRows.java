package org.example;


import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;

//本例子将展示如何平移行与列，平移数据的本质，其实是删除行/列或插入行/列
//插入和删除行列：https://www.grapecity.com.cn/developer/grapecitydocuments/excel-java/docs/Features/ManageWorksheet/RangeOperations/InsertAndDeleteRowsAndColumns
public class Example3_ShiftColumnAndRows {
    public static void main(String[] args) {
        Workbook wb = new Workbook();
        wb.open("Examples/src/main/resources/Example3_ShiftColumnAndRows.xlsx");
        IWorksheet sheet = wb.getWorksheets().get(0);
        sheet.getRange("A3:A5").getEntireRow().insert();
        sheet.getRange("A3:A5").getEntireRow().delete();

        sheet.getRange("A3:C3").getEntireColumn().insert();
        sheet.getRange("A3:C3").getEntireColumn().delete();
    }
}
