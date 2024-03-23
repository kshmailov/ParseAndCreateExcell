package org.example;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class CreateExcelTable {
    public static int rowNum = 0;
    public static FileOutputStream out;
    public static XSSFWorkbook workbook= new XSSFWorkbook();
    public static XSSFSheet sheet = workbook.createSheet("table");

    public static void table(List<String> list) throws IOException {
        out = new FileOutputStream("data/Table.xlsx");
        ArrayList<String> table = new ArrayList<>(list);
        for (String string : table) {
            String[] tableString = string.split(" ");
            XSSFRow row = sheet.createRow(rowNum++);
            for (int i = 0; i < tableString.length; i++) {
                XSSFCell cell = row.createCell(i);
                cell.setCellValue(tableString[i]);
            }
        }
    }

    public static void closeTable() throws IOException {
        workbook.write(out);
        out.close();
    }
}
