package org.example;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Map;
public class CreateExcelScheme {
    public static int rowNum = 0;
    public static FileOutputStream out;
    public static XSSFWorkbook workbook= new XSSFWorkbook();
    public static XSSFSheet sheet = workbook.createSheet("Schemes");
    public static void scheme(Map<String, String> map) throws IOException {
        out = new FileOutputStream("data/Schemes.xlsx");
        ArrayList<String> nameSchemes = new ArrayList<>(map.keySet());
        nameSchemes.sort(Comparator.comparing(String::length).reversed());
        for (String nameScheme : nameSchemes) {
            XSSFRow row = sheet.createRow(rowNum++);
            XSSFCell cell1 = row.createCell(0);
            cell1.setCellValue(nameScheme);
            XSSFCell cell2 = row.createCell(1);
            cell2.setCellValue(map.get(nameScheme));
        }
    }

    public static void closeScheme() throws IOException {
        workbook.write(out);
        out.close();
    }
}
