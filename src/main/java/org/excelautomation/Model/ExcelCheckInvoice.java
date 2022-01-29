package org.excelautomation.Model;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.excelautomation.Controller.Controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

public class ExcelCheckInvoice {

    private final Controller controller;
    private final Set<String> invoiceBadSet = new HashSet<>();
    private final Map<Integer, String> errors = new HashMap<>();
    private final Map<String, Map<Integer, Boolean>> checkedColumns = new HashMap<>();
    private final Map<String, Map<Integer, String>> invoiceData = new HashMap<>();
    private final int[] columnsToRead = {23, 24, 25, 27};

    public ExcelCheckInvoice(Controller controller) {
        this.controller = controller;
        errors.put(23, "В СКИДКЕ");
        errors.put(24, "В ТИПЕ ЦЕНЫ");
        errors.put(25, "В ВАЛЮТЕ МАТЕРИАЛОВ");
        errors.put(27, "В ВАЛЮТЕ СМР");
    }

    private Map<Integer, Boolean> generateNewMap() {
        Map<Integer, Boolean> result = new HashMap<>();
        result.put(23, false);
        result.put(24, false);
        result.put(25, false);
        result.put(27, false);
        return result;
    }

    public void check(File file) {
        ZipSecureFile.setMinInflateRatio(0);
        boolean allFine = true;
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
            Sheet sheet = workbook.getSheet("Спецификация");
            int rowNum = 0;
            int rowsInFile = sheet.getLastRowNum() - 17;
            for (Row row : sheet) {
                rowNum++;
                if (rowNum == 1) {
                    continue;
                }
                if (rowNum == rowsInFile) {
                    break;
                }
                row.getCell(45, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setBlank();
                Cell cell = row.getCell(33);
                String invoice = cell.getStringCellValue();
                if (!checkedColumns.containsKey(invoice)) {
                    checkedColumns.put(invoice, generateNewMap());
                }
                cell = row.getCell(44);
                if (invoice.equals("") || cell.getCellType() != CellType.BLANK) {
                    continue;
                }
                String cellValue = "";
                Map<Integer, String> invoiceValue = new HashMap<>();
                for (int index : columnsToRead) {
                    cell = row.getCell(index - 1);
                    switch (cell.getCellType()) {
                        case NUMERIC -> cellValue = String.valueOf(cell.getNumericCellValue());
                        case STRING -> cellValue = cell.getStringCellValue();
                    }
                    invoiceValue.put(index, cellValue);
                }
                if (invoiceData.containsKey(invoice)) {
                    StringBuilder error = new StringBuilder();
                    for (int index : columnsToRead) {
                        if (!invoiceData.get(invoice).get(index).equals(invoiceValue.get(index))) {
                            error.append(errors.get(index)).append(";");
                            allFine = false;
                        }
                    }
                    if (!error.toString().equals("")) {
                        row.getCell(45, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(error.toString());
                    }
                } else {
                    invoiceData.put(invoice, invoiceValue);
                }
            }
            if (allFine) {
                controller.sendMessage(MessageType.INFO, "Ошибок нет.");
            } else {
                controller.sendMessage(MessageType.INFO,String.format("В файле: %s есть ошибки.%n", file.getName()));
            }
            try (FileOutputStream outputStream = new FileOutputStream(file)) {
                workbook.write(outputStream);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
