package org.excelautomation.Model;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

import org.excelautomation.Controller.Controller;

public class ExcelParser {

    private boolean fail = false;
    private final Controller controller;
    private static final int TEMPLATE_COLUMNS = 43;
    private static final int SPECS_COLUMNS = 14;
    private final Set<Integer> cellsToSkip = new HashSet<>();

    public ExcelParser(Controller controller) {
        this.controller = controller;
        Integer[] cellsToSkipArray = { 19, 20, 21, 29, 30, 32, 38, 41, 42 };
        cellsToSkip.addAll(Arrays.asList(cellsToSkipArray));
    }

    private boolean isNumeric(Cell cell) {
        if (cell == null) {
            return true;
        }
        try {
            cell.getNumericCellValue();
        } catch (IllegalStateException nfe) {
            return false;
        }
        return true;
    }

    public List<Map<Integer, String>> readExcelFile(File file, boolean isTemplate) {
        ZipSecureFile.setMinInflateRatio(0);
        List<Map<Integer, String>> data = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file)) {
            int startColumn = isTemplate ? 1 : 0;
            Sheet sheet;
            if (isTemplate) {
                sheet = workbook.getSheet("Спецификация");
            }
            else sheet = workbook.getSheetAt(0);
            int columnsToRead = isTemplate ? TEMPLATE_COLUMNS : SPECS_COLUMNS;
            for (Row row : sheet) {
                if (isTemplate) {
                    if (row.equals(sheet.getRow(0)) || row.getRowNum() > sheet.getLastRowNum() - 16) continue;
                    if (row.getCell(1).getCellType() == CellType.BLANK && row.getCell(2).getCellType() == CellType.BLANK)
                        break;
                }
                Map<Integer, String> rowData = new HashMap<>();
                for (int j = startColumn; j < columnsToRead; j++) {
                    Cell cell = row.getCell(j);
                    int index = j;
                    if (cell != null) {
                        if (!isTemplate) {
                            checkCellFormat(cell, j, row, file);
                            index = getIndex(j);
                        }
                        if (isTemplate && cellsToSkip.contains(j + 1)) continue;
                        switch (cell.getCellType()) {
                            case STRING -> rowData.put(index, String.valueOf(cell.getRichStringCellValue()));
                            case FORMULA -> rowData.put(index, cell.getCellFormula());
                            case NUMERIC -> rowData.put(index, String.valueOf(cell.getNumericCellValue()));
                        }
                    }
                }
                if (!isTemplate) rowData.put(getIndex(14), String.valueOf(file.getName().charAt(0)));
                data.add(rowData);
            }
        } catch (FileNotFoundException e) {
            System.err.println("FILE NOT FOUND: " + file);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }

        return fail ? null : data;
    }

    private int getIndex(int j) {
        if (j < 2) return j + 1;
        else if (j <= 12) return j + 5;
        else if (j == 13) return 33;
        else if (j == 14) return 42;
        return 0;
    }

    private void checkCellFormat(Cell cell, int col, Row row, File file) {
        if ((col < 2 || col == 10) && (cell == null || cell.getCellType().compareTo(CellType.BLANK) == 0)) {
            controller.error((col < 1 ? ErrorType.WRONG_AB_COLUMN : ErrorType.WRONG_K_COLUMN), file.getName(), row);
            fail = true;
        }

        if (col == 12 && !isNumeric(cell)) {
            controller.error(ErrorType.WRONG_M_COLUMN, file.getName(), row);
            fail = true;
        }
    }


}
