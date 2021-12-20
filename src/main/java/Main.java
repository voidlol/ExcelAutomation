import org.excelautomation.Controller.Controller;
import org.excelautomation.View.ConsoleOutput;
import org.excelautomation.View.View;

public class Main {

    public static void main(String[] args) {
        View display = new ConsoleOutput();
        Controller controller;
        if (args.length != 0) {
            controller = new Controller(display, args[0]);
        } else
            controller = new Controller(display, null);

        controller.init();
    }
}
