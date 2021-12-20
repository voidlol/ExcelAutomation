package org.excelautomation.View;

import org.excelautomation.Model.MessageType;

public interface View {
    void print(String message);
    void print(String message, MessageType mp);
    void println(String message);
    void println(String message, MessageType mp);
}
