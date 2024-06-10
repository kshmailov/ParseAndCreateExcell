package org.example;


import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.HashMap;
import java.util.Map;



public class App 
{

    public static void main(String[] args) throws IOException {
//        Map<String, String> schemesUg = new HashMap<>(ParseExcelScheme.parseExcel("data/Ug.xlsx", 4));
//        CreateExcelScheme.scheme(schemesUg);
//        Map<String, String> schemesManic = new HashMap<>(ParseExcelScheme.parseExcel("data/Manich.xlsx", 5));
//        CreateExcelScheme.scheme(schemesManic);
//        Map<String, String> schemesCuban = new HashMap<>(ParseExcelScheme.parseExcel("data/Kuban.xlsx", 5));
//        CreateExcelScheme.scheme(schemesCuban);
//        CreateExcelScheme.closeScheme();
//        ArrayList<String> tableUvUg = new ArrayList<>(ParseExcelTable.parseExcelTable("data/Ug.xlsx", 4));
//        CreateExcelTable.table(tableUvUg);
//        HashSet<String> ugSheet1 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Ug.xlsx", 0));
//        HashSet<String> ugSheet2 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Ug.xlsx", 1));
//        HashSet<String> ugSheet3 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Ug.xlsx", 2));
//        HashSet<String> ugSheet4 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Ug.xlsx", 3));
//        ArrayList<String> listFpoUg =new ArrayList<>();
//        listFpoUg.addAll(ParseExcelTableFpo.createTableFpoString(ugSheet2,ugSheet1,"{Юг}","КПР-1_\"Юг\"","[Текущий_сезон=ЛЕТО]"));
//        listFpoUg.addAll(ParseExcelTableFpo.createTableFpoString(ugSheet1,ugSheet2,"{Юг}","КПР-2_\"Юг\"","[Текущий_сезон=ЛЕТО]"));
//        listFpoUg.addAll(ParseExcelTableFpo.createTableFpoString(ugSheet4,ugSheet3,"{Юг}","КПР-1_\"Юг\"","[Текущий_сезон=ЗИМА]"));
//        listFpoUg.addAll(ParseExcelTableFpo.createTableFpoString(ugSheet3,ugSheet4,"{Юг}","КПР-2_\"Юг\"","[Текущий_сезон=ЗИМА]"));
//        CreateExcelTable.table(listFpoUg);
//        ArrayList<String> tableUvManich = new ArrayList<>(ParseExcelTable.parseExcelTable("data/Manich.xlsx", 5));
//        CreateExcelTable.table(tableUvManich);
//        HashSet<String> manichSheet1 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Manich.xlsx", 0));
//        HashSet<String> manichSheet2 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Manich.xlsx", 1));
//        HashSet<String> manichSheet3 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Manich.xlsx", 2));
//        HashSet<String> manichSheet4 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Manich.xlsx", 3));
//        HashSet<String> manichSheet5 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Manich.xlsx", 4));
//        ArrayList<String> listFpoManich =new ArrayList<>();
//        listFpoManich.addAll(ParseExcelTableFpo.createTableFpoString(manichSheet2,manichSheet1,"{Ман}","КПР-1_\"Маныч\"","[Текущий_сезон=ЛЕТО]"));
//        listFpoManich.addAll(ParseExcelTableFpo.createTableFpoString(manichSheet1,manichSheet2,"{Ман}","КПР-2_\"Маныч\"","[Текущий_сезон=ЛЕТО]"));
//        listFpoManich.addAll(ParseExcelTableFpo.createTableFpoString(manichSheet3,manichSheet5,"{Ман}","КПР-2_\"Маныч\"","[Текущий_сезон=ЗИМА]"));
//        listFpoManich.addAll(ParseExcelTableFpo.createTableFpoString(manichSheet4,manichSheet5,"{Ман}","КПР-2_\"Маныч\"","[Текущий_сезон=ЗИМА]"));
//        listFpoManich.addAll(ParseExcelTableFpo.createTableFpoString(manichSheet4,manichSheet3,"{Ман}","КПР-3_\"Маныч\"","[Текущий_сезон=ЗИМА]"));
//        listFpoManich.addAll(ParseExcelTableFpo.createTableFpoString(manichSheet5,manichSheet3,"{Ман}","КПР-3_\"Маныч\"","[Текущий_сезон=ЗИМА]"));
//        listFpoManich.addAll(ParseExcelTableFpo.createTableFpoString(manichSheet3,manichSheet4,"{Ман}","КПР-4_\"Маныч\"","[Текущий_сезон=ЗИМА]"));
//        listFpoManich.addAll(ParseExcelTableFpo.createTableFpoString(manichSheet5,manichSheet4,"{Ман}","КПР-4_\"Маныч\"","[Текущий_сезон=ЗИМА]"));
//        CreateExcelTable.table(listFpoManich);
//        ArrayList<String> tableUvKuban = new ArrayList<>(ParseExcelTable.parseExcelTable("data/Kuban.xlsx", 5));
//        CreateExcelTable.table(tableUvKuban);
//        HashSet<String> kubanSheet1 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Kuban.xlsx", 0));
//        HashSet<String> kubanSheet2 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Kuban.xlsx", 1));
//        HashSet<String> kubanSheet3 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Kuban.xlsx", 2));
//        HashSet<String> kubanSheet4 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Kuban.xlsx", 3));
//        HashSet<String> kubanSheet5 =new HashSet<>(ParseExcelTableFpo.parseExcelTableFpo("data/Kuban.xlsx", 4));
//        ArrayList<String> listFpoKuban =new ArrayList<>();
//        listFpoKuban.addAll(ParseExcelTableFpo.createTableFpoString(kubanSheet2,kubanSheet1,"{Куб}","КПР-1_\"Кубанское\"","[Текущий_сезон=ЛЕТО]"));
//        listFpoKuban.addAll(ParseExcelTableFpo.createTableFpoString(kubanSheet3,kubanSheet1,"{Куб}","КПР-1_\"Кубанское\"","[Текущий_сезон=ЛЕТО]"));
//        listFpoKuban.addAll(ParseExcelTableFpo.createTableFpoString(kubanSheet1,kubanSheet2,"{Куб}","КПР-2_\"Кубанское\"","[Текущий_сезон=ЛЕТО]"));
//        listFpoKuban.addAll(ParseExcelTableFpo.createTableFpoString(kubanSheet3,kubanSheet2,"{Куб}","КПР-2_\"Кубанское\"","[Текущий_сезон=ЛЕТО]"));
//        listFpoKuban.addAll(ParseExcelTableFpo.createTableFpoString(kubanSheet1,kubanSheet3,"{Куб}","КПР-3_\"Кубанское\"","[Текущий_сезон=ЛЕТО]"));
//        listFpoKuban.addAll(ParseExcelTableFpo.createTableFpoString(kubanSheet2,kubanSheet3,"{Куб}","КПР-3_\"Кубанское\"","[Текущий_сезон=ЛЕТО]"));
//        listFpoKuban.addAll(ParseExcelTableFpo.createTableFpoString(kubanSheet5,kubanSheet4,"{Куб}","КПР-1_\"Кубанское\"","[Текущий_сезон=ЗИМА]"));
//        listFpoKuban.addAll(ParseExcelTableFpo.createTableFpoString(kubanSheet4,kubanSheet5,"{Куб}","КПР-2_\"Кубанское\"","[Текущий_сезон=ЗИМА]"));
//        CreateExcelTable.table(listFpoKuban);
        ArrayList<String> listShunt = new ArrayList<>(ParseExcelTableShunt.parseExcelTableShunt());
        CreateExcelTable.table(listShunt);
        CreateExcelTable.closeTable();

    }
}
