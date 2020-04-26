package com.company;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
    static String pathGL = "C:\\Users\\MarkusReed\\Desktop\\study\\1_Katia\\sbis\\Nakladnaya_25_10_2019-2.xls";
    static int[] cellNums = new int[] {2, 3, 5};
    static final String PATTERN_NAME = "[а-яА-Я]{3,}(\\s|\\W|$)";
    static final String PATTERN_ADJECTIVE = "[а-яА-Я]{3,}(ые|ие|ый|ий|ая|ое)";
    static final String PATTERN_CHILD = "(д/(дев|мал)[.]?)|((Д|д)ет(ск..|[.]?(\\s|\\W|$)))|(яс[.])|(ясельн([.]|..))|(подростков..)";
    static final String PATTERN_MAN = "(М|м)уж(ск..|[.]?(\\s|\\W|$))";
    static final String PATTERN_WOMAN = "(Ж|ж)ен(ск..|[.]?(\\s|\\W|$))";
    static final String PATTERN_SIZE = "([рp][.]\\d{1,3}([/]|[-])?\\d{0,3}([/]|[-])?\\d{0,3}(\\s|$))|(\\d{1,3}([/]|[-])\\d{1,3}(\\s|$))|(\\d{2,3}[XxХх*]\\d{2,3}(\\s|$))|(\\d{1,3}($|([,]$)|([.]$)))|([(]\\d{1,3}[)]($|([,]$)|([.]$)))";
    static int EMPTY_ROW = 0;

    public static void main(String[] args) throws IOException {
        int workbookSize = openWorkbook(pathGL).getSheetAt(0).getLastRowNum();
        for (int i = 0; i < workbookSize; i++) {
            writeRow(correctRow(openSheet(openWorkbook(pathGL), 0).getRow(i), cellNums), i-EMPTY_ROW);
        }
        System.out.println("Rows: " + (workbookSize-EMPTY_ROW) + "\nComplete!");
    }

    public static Workbook openWorkbook(String path) throws IOException {
        try {
            FileInputStream fileInputStream = new FileInputStream(path);
            Workbook workbook = new HSSFWorkbook(fileInputStream);
            fileInputStream.close();
            return workbook;
        } catch (FileNotFoundException e) {
            return new HSSFWorkbook();
        }
    }

    public static Sheet openSheet(Workbook workbook, int sheetNum) {
        try {
            return  workbook.getSheetAt(sheetNum);
        } catch (IllegalArgumentException e) {
            return workbook.createSheet("Sheet"+ sheetNum);
        }
    }

    public  static  boolean checkCell(Row row, int cellNum, CellType cellType) {
        return row.getCell(cellNum, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null
                && row.getCell(cellNum).getCellType().equals(cellType);
    }

    public  static boolean checkRow(Row row){
        return checkCell(row, cellNums[0], CellType.STRING) && (checkCell(row, cellNums[1], CellType.NUMERIC)
                || checkCell(row, cellNums[1], CellType.FORMULA)) && (checkCell(row, cellNums[2], CellType.NUMERIC)
                || checkCell(row, cellNums[2], CellType.FORMULA));
    }

    public static boolean matcherFound(String text, String patternText){
        Pattern pattern = Pattern.compile(patternText);
        Matcher matcher = pattern.matcher(text);
        return matcher.find();
    }

    public static void replaceSubstring(String text, String patternText) {
        Pattern pattern = Pattern.compile(patternText);
        Matcher matcher = pattern.matcher(text);
        while (matcher.find()){
            text = text.replace(text.substring(matcher.start(), matcher.end()), " ");
        }
    }

    public static String addGender(String text) {
        if (matcherFound(text, PATTERN_CHILD))
            return "дет.";
        else if (matcherFound(text, PATTERN_MAN))
            return "муж.";
        else if (matcherFound(text, PATTERN_WOMAN))
            return "жен.";
        return "";
    }

    public static String getCorrectName(String text) {
        Pattern pattern = Pattern.compile(PATTERN_NAME);
        Matcher matcher = pattern.matcher(text);
        matcher.find();
        String name = matcher.group();

        char[] repSimbols = new char[] {',', '.', '-'};
        for (int i = 0; i < repSimbols.length; i++) {
            name = name.replace(repSimbols[i], ' ');
        }
//        name = name.replace(",", " ");
//        name = name.replace(".", " ");
//        name = name.replace("-", " ");

        name = name.substring(0,1).toUpperCase() + name.substring(1);
        return name;
    }

    public static String getCorrectSize(String text) {
        Pattern pattern = Pattern.compile(PATTERN_SIZE);
        Matcher matcher = pattern.matcher(text);
        String temporary = "SIZE";
        while (matcher.find())
            temporary = text.substring(matcher.start(), matcher.end());

        String[][] repSimbols = new String[][] {{"р.", ""}, {"/", " "}, {"-", " "}, {"(", ""}, {")", ""}, {"x", "*"}, {"X", "*"}, {"х", "*"}, {"Х", "*"}};
        for (int i = 0; i < repSimbols.length; i++) {
            temporary = temporary.replace(repSimbols[i][0], repSimbols[i][1]);
        }
//        temporary = temporary.replace("р.", "");
//        temporary = temporary.replace("/", " ");
//        temporary = temporary.replace("-", " ");
//        temporary = temporary.replace("(", "");
//        temporary = temporary.replace(")", "");
//        temporary = temporary.replace("x", "*");
//        temporary = temporary.replace("X", "*");
//        temporary = temporary.replace("х", "*");
//        temporary = temporary.replace("Х", "*");

        try {
            List<Integer> sizes = new ArrayList<>();
            for (String sas : temporary.split(" "))
                sizes.add(Integer.parseInt(sas));
            return Collections.min(sizes).toString();
        } catch (NumberFormatException e) {
            return temporary;
        }
    }

    public static MyRow correctRow(Row row, int[] cellNums){
        if (checkRow(row)){
            String text = row.getCell(cellNums[0]).getStringCellValue();
            double amount = row.getCell(cellNums[1]).getNumericCellValue();
            double cost = row.getCell(cellNums[2]).getNumericCellValue();

            String name = "***";
            name += addGender(text);
            replaceSubstring(text, PATTERN_ADJECTIVE);
            name = name.replace("***", getCorrectName(text));
            name += " " + getCorrectSize(text) + ", ";
            name += (int)Math.round(cost);

            return new MyRow(name, amount, cost);
        } else {
            EMPTY_ROW++;
            return new MyRow();
        }
    }

    public  static void writeRow(MyRow myRow, int i) throws IOException {
        if (myRow.getAmount() != -1 ) {
            String path = pathGL.substring(0, pathGL.lastIndexOf("."));
            Workbook workbook = openWorkbook(path + "(NEW).xls");
            Sheet sheet = openSheet(workbook, 0);

            Row row = sheet.createRow(i);
            Cell nameCell = row.createCell(0);
            Cell amountCell = row.createCell(1);
            Cell unitsСell = row.createCell(2);
            Cell costCell = row.createCell(3);

            nameCell.setCellValue(myRow.getName());
            amountCell.setCellValue(myRow.getAmount());
            unitsСell.setCellValue("шт");
            costCell.setCellValue(myRow.getCost());

            FileOutputStream fileOutputStream = new FileOutputStream(path + "(NEW).xls");
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        }
    }
}