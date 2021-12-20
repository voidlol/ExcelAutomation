package org.excelautomation.Model;

import org.excelautomation.Controller.Controller;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class Model {
    private final String path;
    private final Controller controller;
    private final List<List<String>> overallData = new ArrayList<>();
    private final ExcelWriter excelWriter;
    private final ExcelParser excelParser;
    private boolean fail = false;

    public Model(String path, Controller controller) {
        this.controller = controller;
        excelWriter = new ExcelWriter(controller);
        excelParser = new ExcelParser(controller);
        this.path = path == null ? "./" : path;
    }

    public void readFolder() {
        long startReadingTime = System.currentTimeMillis();
        File folder = new File(path);
        File[] files = folder.listFiles();
        if (checkFolder(files)) {
            File template = null;
            int n = 0;
            for (File f : files) {
                String name = f.getName();
                if (name.endsWith(".xlsm")) template = f;
                if (!name.startsWith("~$") && name.endsWith(".xlsx")) {
                    readFile(f, ++n);
                }
            }
            long endReadingTime = System.currentTimeMillis();
            controller.sendMessage(MessageType.READING_TIME, startReadingTime, endReadingTime);
            if (fail) return;
            writeTemplate(template);
        }
    }

    private void writeTemplate(File template) {
        excelWriter.doMerge(overallData, template);
    }

    private void readFile(File f, int n) {
        controller.sendMessage(MessageType.INFO, f.getName(), n);
        Map<Integer, List<String>> map = excelParser.readExcel(f);
        if (map == null) {
            fail = true;
            return;
        }
        overallData.addAll(map.values());
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
            if (name.endsWith(".xlsm")) {
                template = true;
            } else if (!name.startsWith("~$") && name.endsWith(".xlsx")) {
                if (!(name.startsWith("и.") || name.startsWith("с.") || name.startsWith("т")
                 || name.startsWith("м.") || name.startsWith("э") || name.startsWith("о."))) {
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
