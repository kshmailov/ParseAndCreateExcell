package org.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ParseExcelTableFpo {

    protected static HashSet<String> parseExcelTableFpo(String path, int list) throws IOException {
        FileInputStream fis = new FileInputStream(path);

        Workbook workbook;
        try {
            workbook = new XSSFWorkbook(
                    fis);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        HashSet<String> schemes = new HashSet<>();
        Sheet sheet = workbook.getSheetAt(list);
        String schemesKpr = sheet.getRow(0).getCell(3).getStringCellValue();

        String schemesGroup = "";
        if (schemesKpr.contains("Юг")) {
            schemesGroup = "{Юг}";
        } else if (schemesKpr.contains("Кубанское")) {
            schemesGroup = "{Куб}";
        } else if (schemesKpr.contains("Маныч")) {
            schemesGroup = "{Ман}";
        }
        for (Row row : sheet) {
            String cellValue = row.getCell(0).getStringCellValue();
            if (cellValue.contains("Нормальная схема")) {
                String keyNameScheme = "Нормальная_схема" + schemesGroup;
                schemes.add(keyNameScheme);
            } else if (cellValue.contains("или")) {
                String[] schemesFrag = cellValue.split("или");
                for (String frag : schemesFrag) {
                    if (!frag.isEmpty()) {
                        String nameScheme = modifiedNameSchemes(frag);
                        String keyNameScheme = nameScheme + schemesGroup;
                        schemes.add(keyNameScheme);
                    }
                }
            } else if (!modifiedNameSchemes(cellValue).isEmpty()) {
                String nameScheme = modifiedNameSchemes(cellValue);
                String keyNameScheme = nameScheme + schemesGroup;
                schemes.add(keyNameScheme);
            }

        }

        fis.close();
        return schemes;
    }

    public static String modifiedNameSchemes(String name) {
        String regex = "\\(([^)]+)\\)";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(name);
        String modifiedNameSchemes = "";
        while (matcher.find()) {
            modifiedNameSchemes = matcher.group(1);
        }
        return modifiedNameSchemes;
    }
    public static ArrayList<String> createTableFpoString(HashSet<String> firstSheet, HashSet<String> secondSheet,String schemesGroup, String schemesKpr, String sezon){

        String tsTableString = null;
        ArrayList<String> listSchemeFpo = new ArrayList<>();

        switch (schemesGroup) {
            case "{Юг}":
                tsTableString = "Шунт_\"Юг\"_Выведена";
                break;
            case "{Куб}":
                tsTableString = "Шунт_\"Кубанское\"_Выведена";
                break;
            case "{Ман}":
                tsTableString = "Шунт_\"Маныч\"_Выведена";
                break;
        }
        String po = "*ФПО1";
        for (String scheme : firstSheet){
            if (!secondSheet.contains(scheme)){
                StringBuilder tableString = new StringBuilder(String.join(" ", scheme, tsTableString,sezon, po, schemesKpr));
                for (int i =0; i<32; i++){
                    tableString.append(" []");
                }
                listSchemeFpo.add(tableString.toString());
            }
        }
        return listSchemeFpo;
    }
}
