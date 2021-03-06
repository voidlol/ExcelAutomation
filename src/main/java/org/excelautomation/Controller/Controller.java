package org.excelautomation.Controller;

import org.apache.poi.ss.usermodel.Row;
import org.excelautomation.Model.ErrorType;
import org.excelautomation.Model.ExcelCheckInvoice;
import org.excelautomation.Model.MessageType;
import org.excelautomation.Model.Model;
import org.excelautomation.View.View;

import java.io.File;
import java.util.concurrent.TimeUnit;

public class Controller {
    private final View display;
    private final Model model;
    private final ExceptionHandler exceptionHandler;
    public final Config config = Config.readConfig();

    public Controller(View view) {
        this.display = view;
        this.model = new Model(this);
        exceptionHandler = new ExceptionHandler(this.display);
    }

    public void error(ErrorType ep, Exception e) {
        exceptionHandler.handleError(ep, e);
    }

    public void error(ErrorType ep, int rows) {
        exceptionHandler.handleError(ep, rows);
    }

    public void error(ErrorType ep, String fileName, Row row) {
        exceptionHandler.handleError(ep, fileName, row.getRowNum() + 1);
    }

    public void error(ErrorType ep) {
        exceptionHandler.handleError(ep);
    }

    public void init() {
        if (config.isCHECKING()) {
            sendMessage(MessageType.INFO, "Проверяем файл: " + config.getWORKING_TEMPLATE());
            ExcelCheckInvoice excelCheckInvoice = new ExcelCheckInvoice(this);
            excelCheckInvoice.check(new File(config.getWORKING_TEMPLATE()));
        } else {
            sendMessage(MessageType.INFO, "Working template: " + config.getWORKING_TEMPLATE());
            sendMessage(MessageType.INFO, "Empty Template: " + config.getEMPTY_TEMPLATE());
            sendMessage(MessageType.INFO, "Are we adding? " + (config.getIS_ADDING() ? "YES" : "NO"));
            try {
                Thread.sleep(1000);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            model.readFolder(config.getIS_ADDING());
        }
    }

    public void sendMessage(MessageType mp, String fileName, int count) {
        display.println("№" + count + ": " + fileName, mp);
    }

    public void sendMessage(MessageType mp, String message) {
        if (mp == MessageType.PROGRESS) {
            display.print(message);
        } else {
            display.println(message, mp);
        }
    }

    public void sendMessage(MessageType mp) {
        if (mp == MessageType.BLANK) {
            display.print("\n");
        }
    }

    public void sendMessage(MessageType mp, long start, long end) {
        switch (mp) {
            case READING_TIME -> display.println("ЧТЕНИЕ ЗАНЯЛО: " + getTime(start, end), MessageType.INFO);
            case WRITING_TIME -> display.println("ЗАПИСЬ ЗАНЯЛА: " + getTime(start, end), MessageType.INFO);
        }
    }

    public void sendMessage(MessageType mp, Row row, int size) {
        display.print("\rПерерасчет формул: строка " + (row.getRowNum()) + "/" + size + ".", mp);
    }

    private String getTime(long start, long end) {
        return String.format("%02d:%02d:%02d",
                TimeUnit.MILLISECONDS.toHours(end - start),
                TimeUnit.MILLISECONDS.toMinutes(end - start) -
                        TimeUnit.HOURS.toMinutes(TimeUnit.MILLISECONDS.toHours(end - start)),
                TimeUnit.MILLISECONDS.toSeconds(end - start) -
                        TimeUnit.MINUTES.toSeconds(TimeUnit.MILLISECONDS.toMinutes(end - start)));
    }
}
