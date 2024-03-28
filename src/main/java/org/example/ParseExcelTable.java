package org.example;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ParseExcelTable {
    private static TreeMap<String,String> keyMapUv;
    private static TreeMap<String,String> keyMapPo;

    protected static ArrayList<String> parseExcelTable(String path, int count) throws IOException {
        FileInputStream fis = new FileInputStream(path);

        Workbook workbook;
        try {
            workbook = new XSSFWorkbook(
                    fis);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        ArrayList<Sheet> sheets = new ArrayList<>();
        for (int i = 0; i < count; i++) {
            sheets.add(workbook.getSheetAt(i));
        }
        ArrayList<String> tableUV = new ArrayList<>();
        keyMapPo = getMapPo();
        for (Sheet list : sheets) {

            String schemesKpr = list.getRow(0).getCell(3).getStringCellValue();
            String kpr = null;
            switch (schemesKpr){
                case "КПР-1_\"Юг\"":
                    kpr="КПР-1_\"Юг\"";
                    break;
                case "КПР-2_\"Юг\"":
                    kpr="КПР-2_\"Юг\"";
                    break;
                case "КПР-1_\"Маныч\"":
                    kpr="КПР-1_\"Маныч\"";
                    break;
                case "КПР-2_\"Маныч\"":
                    kpr="КПР-2_\"Маныч\"";
                    break;
                case "КПР-3_\"Маныч\"":
                    kpr="КПР-3_\"Маныч\"";
                    break;
                case "КПР-4_\"Маныч\"":
                    kpr="КПР-4_\"Маныч\"";
                    break;
                case "КПР-1_\"Кубанское\"":
                    kpr="КПР-1_\"Кубанское\"";
                    break;
                case "КПР-2_\"Кубанское\"":
                    kpr="КПР-2_\"Кубанское\"";
                    break;
                case "КПР-3_\"Кубанское\"":
                    kpr="КПР-3_\"Кубанское\"";
                    break;
            }
            String sezon = list.getRow(1).getCell(0).getStringCellValue();
            String schemesGroup = "";
            if (schemesKpr.contains("Юг")) {
                schemesGroup = "{Юг}";
            } else if (schemesKpr.contains("Кубанское")) {
                schemesGroup = "{Куб}";
            } else if (schemesKpr.contains("Маныч")) {
                schemesGroup = "{Ман}";
            }
            String sezonGroup;
            if (sezon.toLowerCase().contains("лето")){
                sezonGroup ="Текущий_сезон=ЛЕТО";
            }else {
                sezonGroup ="Текущий_сезон=ЗИМА";
            }
            keyMapUv = getMapUv(schemesGroup, sezonGroup);
            for (Row row : list) {
                String cellValue = row.getCell(0).getStringCellValue();
                if (cellValue.contains("Нормальная схема")) {
                    String keyNameScheme = "Нормальная_схема" + schemesGroup;
                    tableUV.add(createTableString(schemesGroup,sezonGroup, row, keyNameScheme,kpr));
                } else if (cellValue.contains("или")) {
                    String[] schemesFrag = cellValue.split("или");
                    for (String frag : schemesFrag) {
                        if (!frag.isEmpty()) {
                            String nameScheme = modifiedNameSchemes(frag);
                            String keyNameScheme = nameScheme + schemesGroup;
                            tableUV.add(createTableString(schemesGroup,sezonGroup, row, keyNameScheme,kpr));
                        }
                    }
                } else if (!modifiedNameSchemes(cellValue).isEmpty()) {
                    String nameScheme = modifiedNameSchemes(cellValue);
                    String keyNameScheme = nameScheme + schemesGroup;
                    tableUV.add(createTableString(schemesGroup,sezonGroup, row, keyNameScheme,kpr));
                }

            }
        }
        fis.close();
        return tableUV;
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
    public static TreeMap<String, String> getMapUv(String schemesGroup, String sezonGroup){
        TreeMap<String, String> mapUv = new TreeMap<>();
        if (schemesGroup.equals("{Юг}")&&sezonGroup.equals("Текущий_сезон=ЛЕТО")){
            mapUv.put("1", "[[УВ_ОН1-КЭ]]");
            mapUv.put("2", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ]]");
            mapUv.put("3", "[[УВ_ОН1-КЭ];[УВ_ОН4-КЭ]]");
            mapUv.put("4", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ]]");
            mapUv.put("5", "[[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ]]");
            mapUv.put("6", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ]]");
            mapUv.put("7", "[[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ]]");
            mapUv.put("8", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ]]");
            mapUv.put("9", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН6-КЭ]]");
            mapUv.put("10", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН5-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("11", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ];[УВ_ОН6-КЭ]]");
            mapUv.put("12", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН4-КЭ];[УВ_ОН6-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("13", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН6-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("14", "[[УВ_ОН1-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ];[УВ_ОН6-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("15", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ];[УВ_ОН6-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("16", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ];[УВ_ОН6-КЭ];[УВ_ОН200_ВЧ]]");
        } else if (schemesGroup.equals("{Юг}")&&sezonGroup.equals("Текущий_сезон=ЗИМА")) {
            mapUv.put("1", "[[УВ_ОН1-КЭ]]");
            mapUv.put("2", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ]]");
            mapUv.put("3", "[[УВ_ОН1-КЭ];[УВ_ОН4-КЭ]]");
            mapUv.put("4", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ]]");
            mapUv.put("5", "[[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ]]");
            mapUv.put("6", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ]]");
            mapUv.put("7", "[[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ]]");
            mapUv.put("8", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ]]");
            mapUv.put("9", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН5-КЭ];[УВ_ОН6-КЭ]]");
            mapUv.put("10", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("11", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ];[УВ_ОН6-КЭ]]");
            mapUv.put("12", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН4-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("13", "[[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("14", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("15", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН6-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("16", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("17", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН6-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("18", "[[УВ_ОН1-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ];[УВ_ОН6-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("19", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ];[УВ_ОН6-КЭ];[УВ_ОН100_ВЧ]]");
            mapUv.put("20", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ];[УВ_ОН6-КЭ];[УВ_ОН200_ВЧ]]");
        }else if (schemesGroup.equals("{Ман}")&&sezonGroup.equals("Текущий_сезон=ЛЕТО")) {
            mapUv.put("1", "[[УВ_ОН100_ВЧ]]");
            mapUv.put("2", "[[УВ_ОН100_ВЧ]]");
            mapUv.put("3", "[[УВ_ОН100_ВЧ]]");
            mapUv.put("4", "[[УВ_ОН100_ВЧ]]");
            mapUv.put("5", "[[УВ_ОН200_ВЧ]]");
            mapUv.put("6", "[[УВ_ОН200_ВЧ]]");
            mapUv.put("7", "[[УВ_ОН200_ВЧ]]");
            mapUv.put("8", "[[УВ_ОН200_ВЧ]]");
        }else if (schemesGroup.equals("{Ман}")&&sezonGroup.equals("Текущий_сезон=ЗИМА")) {
            mapUv.put("1", "[[УВ_ОН100_ВЧ]]");
            mapUv.put("2", "[[УВ_ОН100_ВЧ]]");
            mapUv.put("3", "[[УВ_ОН100_ВЧ]]");
            mapUv.put("4", "[[УВ_ОН100_ВЧ]]");
            mapUv.put("5", "[[УВ_ОН100_ВЧ]]");
            mapUv.put("6", "[[УВ_ОН100_ВЧ]]");
            mapUv.put("7", "[[УВ_ОН200_ВЧ]]");
            mapUv.put("8", "[[УВ_ОН200_ВЧ]]");
            mapUv.put("9", "[[УВ_ОН200_ВЧ]]");
            mapUv.put("10", "[[УВ_ОН200_ВЧ]]");
            mapUv.put("11", "[[УВ_ОН200_ВЧ]]");
            mapUv.put("12", "[[УВ_ОН200_ВЧ]]");
            mapUv.put("13", "[[УВ_ОН200_ВЧ]]");
        }else if (schemesGroup.equals("{Куб}")&&sezonGroup.equals("Текущий_сезон=ЛЕТО")) {
            mapUv.put("1", "[[УВ_ОН1-КЭ]]");
            mapUv.put("2", "[[УВ_ОН1-КЭ];[УВ_ОН3-КЭ]]");
            mapUv.put("3", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ]]");
            mapUv.put("4", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН5-КЭ]]");
            mapUv.put("5", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ]]");
            mapUv.put("6", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ]]");
            mapUv.put("7", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН6-КЭ]]");
            mapUv.put("8", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ];[УВ_ОН6-КЭ]]");
        }else if (schemesGroup.equals("{Куб}")&&sezonGroup.equals("Текущий_сезон=ЗИМА")) {
            mapUv.put("1", "[[УВ_ОН1-КЭ]]");
            mapUv.put("2", "[[УВ_ОН1-КЭ];[УВ_ОН3-КЭ]]");
            mapUv.put("3", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ]]");
            mapUv.put("4", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН4-КЭ]]");
            mapUv.put("5", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ]]");
            mapUv.put("6", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ]]");
            mapUv.put("7", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН6-КЭ]]");
            mapUv.put("8", "[[УВ_ОН1-КЭ];[УВ_ОН2-КЭ];[УВ_ОН3-КЭ];[УВ_ОН4-КЭ];[УВ_ОН5-КЭ];[УВ_ОН6-КЭ]]");
        }
        return mapUv;
    }
    public static TreeMap<String, String> getMapPo(){
        TreeMap<String, String> mapPo = new TreeMap<>();
        mapPo.put("1", "ПО1_ФОЛ_500кВ_РоАЭС-Тихорецк_1ц");
        mapPo.put("2", "ПО2_ФОЛ_500кВ_РоАЭС-Тихорецк_2ц");
        mapPo.put("3", "ПО3_ФОЛ_500кВ_РоАЭС-Буденновск");
        mapPo.put("4", "ПО4_ФОЛ_500кВ_РоАЭС-Невинномысск");
        mapPo.put("5", "ПО5_ФОЛ_220кВ_НчГРЭС-Койсуг1");
        mapPo.put("6", "ПО6_ФОЛ_220кВ_НчГРЭС-Койсуг2");
        mapPo.put("7", "ПО7_ФОТ_АТ501(502)_ПС_Буденновск");
        mapPo.put("8", "ПО8_ФОТ_АТГ-1(2)_ПС_Невинномысск");
        mapPo.put("9", "ПО9_ФОЛ_500кВ_Ростовская-Тамань");
        mapPo.put("10", "ПО10_ФОЛ_330кВ_НчГРЭС-Тихорецк");
        mapPo.put("11", "ПО11_ФОЛ_Песчанокопск-Ея_тяговая");
        mapPo.put("12", "ПО12_ФОЛ_Тихорецк-Ея_тяговая");
        mapPo.put("13", "ПО13_ФОЛ_Сальская-Песчанокопская");
        mapPo.put("14", "ПО14_ФОЛ_Волгодонск-Сальская");
        mapPo.put("15", "ПО15_ФОЛ_Койсуг-А-20");
        mapPo.put("16", "ПО16_ФОЛ_Р-20-А-20");
        mapPo.put("17", "ПО17_ФОЛ_Койсуг-Крыловская");
        mapPo.put("18", "ПО18_ФОЛ_Тихорецк-Крыловская");
        mapPo.put("19", "ПО19_ФОТ_АТ-1_ПС_Крыловская");
        mapPo.put("20", "ПО20_ФОДЛ_РоАЭС-Тих1_РоАЭС-Тих2");
        return mapPo;
    }
    public static TreeMap<String, String> getMapForm(String schemesGroup){
        TreeMap<String, String> mapFormPo = new TreeMap<>();
        switch (schemesGroup) {
            case "{Юг}":
                mapFormPo.put("1", "ФОЛ_РоАЭС-Тихорецк№1{Юг}");
                mapFormPo.put("2", "ФОЛ_РоАЭС-Тихорецк№2{Юг}");
                mapFormPo.put("3", "ФОЛ_РоАЭС-Буденновск{Юг}");
                mapFormPo.put("4", "ФОЛ_РоАЭС-Невинномыск{Юг}");
                mapFormPo.put("5", "ФОЛ_НчГРЭС-Койсуг№1{Юг}");
                mapFormPo.put("6", "ФОЛ_НчГРЭС-Койсуг№2{Юг}");
                mapFormPo.put("9", "ФОЛ_Ростовская-Тамань{Юг}");
                mapFormPo.put("10", "ФОЛ_НчГРЭС-Тихорецк{Юг}");
                mapFormPo.put("20", "ФОДЛ_РоАЭС-Тихорецк№1_и_№2{Юг}");
                break;
            case "{Ман}":
                mapFormPo.put("3", "ФОЛ_РоАЭС-Буденновск{Маныч}");
                mapFormPo.put("4", "ФОЛ_РоАЭС-Невинномысск{Маныч}");
                mapFormPo.put("7", "ФОТ-501/502_Буденновск{Маныч}");
                mapFormPo.put("8", "ФОТ-1/2_Невинномысск{Маныч}");
                break;
            case "{Куб}":
                mapFormPo.put("1", "ФОЛ_РоАЭС-Тихорецк№1{Куб}");
                mapFormPo.put("2", "ФОЛ_РоАЭС-Тихорецк№2{Куб}");
                mapFormPo.put("5", "ФОЛ_НчГРЭС-Койсуг№1{Куб}");
                mapFormPo.put("6", "ФОЛ_НчГРЭС-Койсуг№2{Куб}");
                mapFormPo.put("9", "ФОЛ_Ростовская-Тамань{Куб}");
                mapFormPo.put("10", "ФОЛ_НчГРЭС-Тихорецк{Куб}");
                mapFormPo.put("11", "ФОЛ_Песч-ЕяТяговая{Куб}");
                mapFormPo.put("12", "ФОЛ_Тихор-ЕяТяговая{Куб}");
                mapFormPo.put("13", "ФОЛ_Сальск-Песчан{Куб}");
                mapFormPo.put("14", "ФОЛ_Волг-Сальск{Куб}");
                mapFormPo.put("15", "ФОЛ_Койсуг-А-20{Куб}");
                mapFormPo.put("16", "ФОЛ_Р-20-А-20{Куб}");
                mapFormPo.put("17", "ФОЛ_Койсуг-Крыловск{Куб}");
                mapFormPo.put("18", "ФОЛ_Тихор-Крыловск{Куб}");
                mapFormPo.put("20", "ФОДЛ_РоАЭС-Тихорецк№1_и_№2{Куб}");
                break;
        }
        return mapFormPo;
    }
    public static String createTableString(String schemesGroup, String sezonGroup,Row row, String keyNameScheme, String schemesKpr){

        String tsTableString = null;

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
        String po = Integer.toString((int)row.getCell(1).getNumericCellValue());
        TreeMap<String,String> mapFormulaPo=getMapForm(schemesGroup);
        String boolFormula ="["+mapFormulaPo.get(po)+"&"+sezonGroup+ "]";
        String poString=keyMapPo.get(po);

        StringBuilder tableString = new StringBuilder(String.join(" ", keyNameScheme, tsTableString, boolFormula, poString, schemesKpr));
        for (int i =3; i<35; i++){
            String uv = Integer.toString((int) row.getCell(i).getNumericCellValue());
            if (uv.equals("0")){
                tableString.append(" []");
            }else {
                tableString.append(" ").append(keyMapUv.get(uv));
            }
        }
        return tableString.toString();
    }
}
