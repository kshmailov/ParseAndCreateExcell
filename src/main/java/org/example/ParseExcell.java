package org.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ParseExcell {
    public static Map<String, String> parseExcel(String path, int count) throws IOException {
        FileInputStream fis = new FileInputStream(path);
        HashMap<String, String> schemes = new HashMap<>();
        Workbook workbook;
        try {
            workbook = new XSSFWorkbook(
                    fis);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        ArrayList<Sheet> sheets = new ArrayList<>();
        for (int i = 0; i<count; i++){
            sheets.add(workbook.getSheetAt(i));
        }
        for (Sheet list : sheets) {

            String schemesKpr = list.getRow(0).getCell(3).getStringCellValue();
            String schemesGroup = "";
            if (schemesKpr.contains("Юг")) {
                schemesGroup = "{Юг}";
            } else if (schemesKpr.contains("Кубанское")) {
                schemesGroup = "{Куб}";
            } else if (schemesKpr.contains("Маныч")) {
                schemesGroup = "{Ман}";
            }
            for (Row row : list) {
                String cellValue = row.getCell(0).getStringCellValue();
                if (cellValue.contains("Нормальная схема")){
                    String keyNameScheme = "[Нормальная_схема"+schemesGroup+"]";
                    schemes.put(keyNameScheme, "");
                }
                else if (cellValue.contains(" или ")) {
                    String[] schemesFrag = cellValue.split(" или ");
                    for (String frag : schemesFrag) {
                        if (!frag.isEmpty()){
                            String nameScheme = modifiedNameSchemes(frag);
                            String tsRemes = tsRemes(nameScheme);
                            String keyNameScheme = "[" + nameScheme + schemesGroup + "]";
                            schemes.put(keyNameScheme, tsRemes);
                        }
                    }
                } else if (!modifiedNameSchemes(cellValue).isEmpty()) {
                    String nameScheme = modifiedNameSchemes(cellValue);
                    String tsRemes = tsRemes(nameScheme);
                    String keyNameScheme = "[" + nameScheme + schemesGroup + "]";
                    schemes.put(keyNameScheme, tsRemes);
                }
            }
        }
        fis.close();
        return schemes;
    }
    public static String modifiedNameSchemes(String name){
        String regex = "\\(([^)]+)\\)";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(name);
        String modifiedNameSchemes ="";
        while (matcher.find()) {
            modifiedNameSchemes = matcher.group(1);
        }
        return modifiedNameSchemes;
    }
    public static String tsRemes(String nameSchemes){
        ArrayList<String> tsRemList = new ArrayList<>(){
            {
                add("[Р1:РоАЭС-Тихорецк_1ц]");
                add("[Р2:РоАЭС-Тихорецк_2ц]");
                add("[Р3:РоАЭС-Буденновск]");
                add("[Р4:РоАЭС-Невинномысск]");
                add("[Р5:Ростовская-Тамань]");
                add("[Р6:АТ-501_Буденновск]");
                add("[Р7:АТ-502_Буденновск]");
                add("[Р8:АТГ-1_ПС_Невинномысск]");
                add("[Р9:АТГ-2_ПС_Невинномысск]");
                add("[Р10:НчГРЭС-Тихорецк]");
                add("[Р11:Волгодонск-Сальск]");
                add("[Р12:Р-20-А-20]");
                add("[Р13:Койсуг-Крыловская]");
                add("[Р14:НчГРЭС-Койсуг_1]");
                add("[Р15:НчГРЭС-Койсуг_2]");
                add("[Р16:А-20-А-30]");
                add("[Р17:Староминская-А-30]");
                add("[Р18:Койсуг-А-20]");
                add("[Р24:Рем_тран_А20-А30-Стармин]");
                add("[Р25:СВ-220_А-30]");
                add("[Р26:1СШ_Тихорецк]");
                add("[Р27:2СШ_Тихорецк]");
                add("[Р28:Невин_Алания]");
                add("[Р29:Ея_тяговая-Песчанокопская]");
                add("[Р30:Тихорецк-Ея_тяговая]");
                add("[Р31:Сальск-Песчанокопская]");
                add("[Р32:Тихорецк-Крыловская]");
            }};
        String ts = "";
        if (nameSchemes.contains("+")) {
            String[] pairNames = nameSchemes.split("\\+");
            String tsRem1 ="";
            String tsRem2 ="";
            for (String tsRemes : tsRemList) {
                int end = tsRemes.indexOf(":");
                String substring = tsRemes.substring(1,end);
                if (substring.equals(pairNames[0])) {
                    tsRem1 = tsRemes;
                } else if (substring.equals(pairNames[1])) {
                    tsRem2 = tsRemes;
                }
            }
            String tsList = String.join(";",tsRem1,tsRem2);
            ts = "["+tsList+"]";

        }else {
            for (String tsRemes : tsRemList) {
                int end = tsRemes.indexOf(":");
                String substring = tsRemes.substring(1,end);
                if (substring.equals(nameSchemes)){
                    ts = tsRemes;
                    break;
                }
            }
        }
        return ts;
    }
}
