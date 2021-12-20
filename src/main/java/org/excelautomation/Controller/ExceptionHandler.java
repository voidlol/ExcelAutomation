package org.excelautomation.Controller;

import org.apache.poi.ss.usermodel.Row;
import org.excelautomation.Model.ErrorType;
import org.excelautomation.Model.MessageType;
import org.excelautomation.View.View;

public class ExceptionHandler {

    private final View display;

    public ExceptionHandler(View display) {
        this.display = display;
    }

    public void handleError(ErrorType ep, Exception e) {
        if (ep == ErrorType.IOEXCEPTION) {
            display.println("ЧТО-ТО ПОШЛО НЕ ТАК :(", MessageType.ERROR);
            if (e != null) {
                display.println(e.getMessage(), MessageType.ERROR);
            }
        }
    }

    public void handleError(ErrorType ep, int rows) {
        switch (ep) {
            case NOT_ENOUGH_ROWS -> display.println("НЕДОСТАТОЧНО СТРОК. ДОЛЖНО БЫТЬ: " + rows, MessageType.ERROR);
        }
    }

    public void handleError(ErrorType ep) {
        switch (ep) {
            case FILE_NOT_FOUND -> display.println("НЕТ ФАЙЛОВ ДЛЯ РАБОТЫ", MessageType.ERROR);
            case NO_PREFIX_ITEMS -> display.println("ОШИБКА В НАЗВАНИИ ФАЙЛОВ. ОТСУТСВТУЕТ ПРЕФИКС", MessageType.ERROR);
            case TEMPLATE_NOT_FOUND -> display.println("ОТСУТСТВЕТ ШАБЛОН С МАКРОСОМ", MessageType.ERROR);
        }
    }

    public void handleError(ErrorType ep, String fileName, int row) {
        switch (ep) {
            case WRONG_AB_COLUMN -> display.println("ПРОПУЩЕН ШИФР ИЛИ НАИМЕНОВАНИЕ РАЗДЕЛА В ФАЙЛЕ: "
                    + fileName + " В СТРОКЕ: " + row, MessageType.ERROR);
            case WRONG_K_COLUMN -> display.println("ПРОПУЩЕН ВЕНДОР В ФАЙЛЕ: " + fileName
                    + " В СТРОКЕ: " + row, MessageType.ERROR);
            case WRONG_M_COLUMN -> display.println("ОШИБКА В СТОЛБЦЕ С ОБЪЕМАМИ В ФАЙЛЕ: "
                    + fileName + " В СТРОКЕ: " + row, MessageType.ERROR);
        }
    }

}
