package org.excelautomation.Model;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.excelautomation.Controller.Controller;

public class ExcelParser {

    private boolean fail = false;
    private final Controller controller;

    public ExcelParser(Controller controller) {
        this.controller = controller;
    }

    private boolean isNumeric(Cell cell) {
        if (cell == null) {
            return true;
        }
        try {
            double d = cell.getNumericCellValue();
        } catch (IllegalStateException nfe) {
            return false;
        }
        return true;
    }

    public Map<Integer, List<String>> readExcel(File file) {
        Map<Integer, List<String>> data = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            int i = 0;
            for (Row row : sheet) {
                data.put(i, new ArrayList<String>());
                for (int j = 0; j < 14; j++) {
                    Cell cell = row.getCell(j);
                    checkCellFormat(cell, j, row, file);
                    if (cell == null) {
                        data.get(i).add("!EMPTY!");
                        continue;
                    }
                    switch (cell.getCellType()) {
                        case STRING -> data.get(i).add(String.valueOf(cell.getRichStringCellValue()));
                        case NUMERIC -> data.get(i).add(String.valueOf(cell.getNumericCellValue()));
                        case BLANK -> data.get(i).add("!EMPTY!");
                    }
                }
                data.get(i).add(String.valueOf(file.getName().charAt(0)));
                i++;
            }
        } catch (FileNotFoundException e) {
            System.err.println("FILE NOT FOUND: " + file);
        } catch (IOException e) {
            e.printStackTrace();
        }

        return fail ? null : data;
    }

    private void checkCellFormat(Cell cell, int col, Row row, File file) {
        if ((col < 2 || col == 10) && (cell == null || cell.getCellType().compareTo(CellType.BLANK) == 0)) {
            controller.error((col < 1 ? ErrorType.WRONG_AB_COLUMN : ErrorType.WRONG_K_COLUMN), file.getName(), row, 0);
            fail = true;
        }

        if (col == 12 && !isNumeric(cell)) {
            controller.error(ErrorType.WRONG_M_COLUMN, file.getName(), row, 0);
            fail = true;
        }
    }

}
