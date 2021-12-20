package org.excelautomation.Model;

import org.excelautomation.Controller.Controller;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/*
    и. -> CORAL;
    м. -> LIGHT_GREEN;
    о. -> LEMON_CHIFFON;
    с. -> LIGHT_TURQUOISE;
    т. -> GREY_25_PERCENT;
    э. -> TAN;
 */

public class ExcelWriter {
    private int startRow = 1;
    private final Controller controller;
    private final static Map<Character, IndexedColors> colors = new HashMap<>();

    static {
        colors.put('и', IndexedColors.CORAL);
        colors.put('м', IndexedColors.LIGHT_GREEN);
        colors.put('о', IndexedColors.LEMON_CHIFFON);
        colors.put('с', IndexedColors.LIGHT_TURQUOISE);
        colors.put('т', IndexedColors.GREY_25_PERCENT);
        colors.put('э', IndexedColors.TAN);
    }

    public ExcelWriter(Controller controller) {
        this.controller = controller;
    }

    public void doMerge(List<List<String>> overallData, File file) {
        ZipSecureFile.setMinInflateRatio(0);
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            long startWritingTime = System.currentTimeMillis();
            Sheet sheet = workbook.getSheet("Спецификация");
            int rowsInFile = sheet.getLastRowNum() - 17;
            if (rowsInFile < overallData.size()) {
                controller.error(ErrorType.NOT_ENOUGH_ROWS, overallData.size() + 18);
                Thread.sleep(10000);
                return;
            }
            for (List<String> rowData : overallData) {
                Row row = sheet.getRow(startRow++);
                for (int i = 0; i < 14; i++) {
                    String data = rowData.get(i);
                    if (!data.equals("!EMPTY!")) {
                        if (i < 2) row.getCell(i + 1).setCellValue(data);
                        if (i >= 2 && i < 12) row.getCell(5 + i).setCellValue(data);
                        if (i == 12) row.getCell(17).setCellValue(Double.parseDouble(data));
                        if (i == 13) row.getCell(33).setCellValue(data);
                    }
                }
                if (!rowData.get(14).isEmpty()) {
                    CellStyle headerStyle = workbook.createCellStyle();
                    headerStyle.setFillForegroundColor(colors.get(rowData.get(14).charAt(0)).getIndex());
                    headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    row.getCell(1).setCellStyle(headerStyle);
                    row.getCell(2).setCellStyle(headerStyle);
                }
            }
            FileOutputStream outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
            long endWritingTime = System.currentTimeMillis();
            controller.sendMessage(MessageType.WRITING_TIME, startWritingTime, endWritingTime);
        } catch (IOException | InterruptedException e) {
            controller.error(ErrorType.IOEXCEPTION, e);
        }
    }
}
