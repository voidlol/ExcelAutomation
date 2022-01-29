package org.excelautomation.View;

import org.excelautomation.Model.MessageType;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class ConsoleOutput implements View {

    private String appendTime() {
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH:mm:ss ");
        return dtf.format(LocalDateTime.now());
    }

    @Override
    public void print(String message) {
        this.print(message, MessageType.INFO);
    }

    @Override
    public void print(String message, MessageType mp) {
        switch (mp) {
            case INFO -> System.out.print(message.equals("\n") ? message : appendTime() + message);
            case ERROR -> System.err.print(appendTime() + message);
        }
    }

    @Override
    public void println(String message) {
        this.println(message, MessageType.INFO);
    }

    @Override
    public void println(String message, MessageType mp) {
        switch (mp) {
            case INFO -> System.out.println(appendTime() + message);
            case ERROR -> System.err.println(appendTime() + message);
        }
    }
}
