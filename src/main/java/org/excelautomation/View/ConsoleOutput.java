package org.excelautomation.View;

import org.excelautomation.Model.MessageType;

public class ConsoleOutput implements View {

    @Override
    public void print(String message) {
        this.print(message, MessageType.INFO);
    }

    @Override
    public void print(String message, MessageType mp) {
        switch (mp) {
            case INFO -> System.out.print(message);
            case ERROR -> System.err.print(message);
        }
    }

    @Override
    public void println(String message) {
        this.println(message, MessageType.INFO);
    }

    @Override
    public void println(String message, MessageType mp) {
        switch (mp) {
            case INFO -> System.out.println(message);
            case ERROR -> System.err.println(message);
        }
    }
}
