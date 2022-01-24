package org.excelautomation.Model;

import org.excelautomation.Controller.Controller;
import java.io.File;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

public class Model {
    private final String path;
    private final Controller controller;
    private final List<Map<Integer, String>> overallData = new ArrayList<>();
    private final ExcelWriter excelWriter;
    private final ExcelParser excelParser;
    private boolean fail = false;

    public Model(Controller controller) {
        this.controller = controller;
        excelWriter = new ExcelWriter(controller);
        excelParser = new ExcelParser(controller);
        this.path = "./";
    }

    public void readFolder(boolean isAdding) {
        long startReadingTime = System.currentTimeMillis();
        File folder = new File(path);
        File[] files = folder.listFiles();
        if (checkFolder(files)) {
            File template = new File(controller.config.getEMPTY_TEMPLATE());
            File workingTemplate = new File(controller.config.getWORKING_TEMPLATE());
            int n = 0;
            for (File f : files) {
                String name = f.getName();
                if (!name.startsWith("~$") && name.endsWith(".xlsx")) {
                    readFile(f, ++n, false);
                }
            }
            if (n == 0) {
                controller.error(ErrorType.FILE_NOT_FOUND);
                return;
            }
            if (isAdding) {
                readFile(workingTemplate, ++n, true);
            }
            long endReadingTime = System.currentTimeMillis();
            controller.sendMessage(MessageType.READING_TIME, startReadingTime, endReadingTime);
            if (fail) return;
            writeTemplate(template);
        }
    }

    private void writeTemplate(File template) {
        overallData.sort(Comparator.comparing(l -> l.get(42)));
        excelWriter.doMerge(overallData, template);
    }

    private void readFile(File f, int n, boolean isTemplate) {
        controller.sendMessage(MessageType.INFO, f.getName(), n);
        List<Map<Integer, String>> excelFile = excelParser.readExcelFile(f, isTemplate);
        if (excelFile == null) {
            fail = true;
            return;
        }
        overallData.addAll(excelFile);
    }

    private boolean checkFolder(File[] files) {
        boolean template = false;
        boolean excelFiles = false;
        if (files == null || files.length < 2) {
            controller.error(ErrorType.FILE_NOT_FOUND);
            return false;
        }
        for (File f : files) {
            String name = f.getName();
            if (name.equals(controller.config.getEMPTY_TEMPLATE()) || name.equals(controller.config.getWORKING_TEMPLATE())) {
                template = true;
            } else if (!name.startsWith("~$") && name.endsWith(".xlsx")) {
                if (!(name.startsWith("и.") || name.startsWith("с.") || name.startsWith("т")
                 || name.startsWith("м.") || name.startsWith("э") || name.startsWith("о.") || name.startsWith("я."))) {
                    controller.error(ErrorType.NO_PREFIX_ITEMS);
                    return false;
                }
                excelFiles = true;
            }
        }
        if (!template) {
            controller.error(ErrorType.TEMPLATE_NOT_FOUND);
            return false;
        }
        if (!excelFiles) {
            controller.error(ErrorType.FILE_NOT_FOUND);
            return false;
        }
        return true;
    }
}
