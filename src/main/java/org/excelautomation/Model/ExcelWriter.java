package org.excelautomation.Model;

import org.excelautomation.Controller.Controller;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

/*
    и. -> CORAL;
    м. -> LIGHT_GREEN;
    о. -> LEMON_CHIFFON;
    с. -> LIGHT_TURQUOISE;
    т. -> GREY_25_PERCENT;
    э. -> TAN;
 */

public class ExcelWriter implements Runnable {
    private int startRow = 1;
    private final Controller controller;
    private final static String EMPTY_STRING = "";
    private final static Map<Character, IndexedColors> colors = new HashMap<>();
    private final static Set<Integer> numericCells = new HashSet<>();
    private final static Set<Integer> formulaCells = new HashSet<>();
    private final static Map<Integer, String> defaultValues = new HashMap<>();
    private int progress = 0;
    private int totalRows;
    private List<Map<Integer, String>> overallData;
    private File file;

    static {
        colors.put('и', IndexedColors.CORAL);
        colors.put('м', IndexedColors.LIGHT_GREEN);
        colors.put('о', IndexedColors.LEMON_CHIFFON);
        colors.put('с', IndexedColors.LIGHT_TURQUOISE);
        colors.put('т', IndexedColors.GREY_25_PERCENT);
        colors.put('э', IndexedColors.TAN);

        defaultValues.put(22, EMPTY_STRING);
        defaultValues.put(23, "0");
        defaultValues.put(24, "Р/С");
        defaultValues.put(25, "RUR");
        defaultValues.put(26, EMPTY_STRING);
        defaultValues.put(27, "RUR");
        defaultValues.put(29, "IF(OR(B3=\"АР\",B3=\"КР\",B3=\"ОР\",B3=\"ММС\",B3=\"СКТС\",B3=\"ЛВС\",B3=\"АИ\",B3=\"новое условие\",B3=\"новое условие\",B3=\"новое условие\",B3=\"новое условие\",B3=\"новое условие\",B3=\"новое условие\",B3=\"новое условие\"),0,T3*0.3)");
        defaultValues.put(32, "T4*0.03");
        defaultValues.put(34, EMPTY_STRING);
        defaultValues.put(35, EMPTY_STRING);
        defaultValues.put(36, EMPTY_STRING);
        defaultValues.put(37, EMPTY_STRING);
        defaultValues.put(38, EMPTY_STRING);
        defaultValues.put(40, EMPTY_STRING);
        defaultValues.put(41, EMPTY_STRING);
    }

    public ExcelWriter(Controller controller) {
        this.controller = controller;
        Integer[] numericCellsArray = { 18, 22, 23, 26, 40, 41 };
        Integer[] formulaCellsArray = { 19, 20, 21, 28, 29, 30, 31, 32, 33, 39, 42, 43};
        numericCells.addAll(Arrays.asList(numericCellsArray));
        formulaCells.addAll(Arrays.asList(formulaCellsArray));
    }

    @Override
    public void run() {
        ZipSecureFile.setMinInflateRatio(0);
        totalRows = overallData.size();
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(file))) {

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            long startWritingTime = System.currentTimeMillis();
            Sheet sheet = workbook.getSheet("Спецификация");
            int rowsInFile = sheet.getLastRowNum() - 17;

            if (rowsInFile < overallData.size()) {
                controller.error(ErrorType.NOT_ENOUGH_ROWS, overallData.size() + 16);
                Thread.sleep(10000);
                return;
            }

            for (Map<Integer, String> rowData : overallData) {
                Row row = sheet.getRow(startRow++);
                progress++;
                controller.sendMessage(MessageType.PROGRESS, "\rЗапись строки: " + getProgress());
                if (controller.config.getIS_ADDING()) {
                    for (int i : defaultValues.keySet()) {
                        if (!rowData.containsKey(i - 1)) {
                            rowData.put(i - 1, defaultValues.get(i));
                        }
                    }
                }
                for (Integer i : rowData.keySet()) {
                    String data = rowData.get(i);
                    if (numericCells.contains(i + 1)) {
                        if (data.isEmpty()) row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setBlank();
                        else row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(Double.parseDouble(data));
                    } else if (formulaCells.contains(i + 1)) {
                        if (!data.matches("-?\\d+(\\.\\d+)?")) {
                            row.getCell(i).setCellFormula(data.replaceAll("(?<=[A-Z])\\d+", String.valueOf(row.getRowNum() + 1)));
                        } else {
                            row.getCell(i).setBlank();
                            row.getCell(i).setCellValue(Double.parseDouble(data));
                        }
                    } else if (i < controller.config.getATTRIBUTE() - 1) {
                        row.getCell(i).setCellValue(data);
                    } else {
                        row.createCell(i).setCellValue(data);
                        CellStyle headerStyle = workbook.createCellStyle();
                        if (data.charAt(0) != 'я' && colors.containsKey(data.charAt(0))) {
                            headerStyle.setFillForegroundColor(colors.get(data.charAt(0)).getIndex());
                            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            row.getCell(1).setCellStyle(headerStyle);
                            row.getCell(2).setCellStyle(headerStyle);
                        }
                    }
                }
            }
            controller.sendMessage(MessageType.BLANK);
            if (controller.config.getIS_ADDING()) {
                for (Row row : sheet) {
                    controller.sendMessage(MessageType.INFO, row, overallData.size());
                    if (row.getRowNum() == overallData.size()) break;
                    for (Cell cell : row) {
                        if (formulaCells.contains(cell.getColumnIndex() + 1)) evaluator.evaluateFormulaCell(cell);
                    }
                }
                controller.sendMessage(MessageType.BLANK);
            }
            try (FileOutputStream outputStream = new FileOutputStream(file)) {
                workbook.write(outputStream);
            }
            long endWritingTime = System.currentTimeMillis();
            controller.sendMessage(MessageType.WRITING_TIME, startWritingTime, endWritingTime);
        } catch (IOException | InterruptedException e) {
            controller.error(ErrorType.IOEXCEPTION, e);
        }
    }

    public void doMerge(List<Map<Integer, String>> overallData, File file) {
        this.overallData = overallData;
        this.file = file;
        new Thread(this).start();
    }


    public String getProgress() {
        return progress + "/" + totalRows + ".";
    }
}
