import org.excelautomation.Controller.Controller;
import org.excelautomation.View.ConsoleOutput;
import org.excelautomation.View.View;

public class Main {

    public static void main(String[] args) {
        View display = new ConsoleOutput();
        Controller controller = new Controller(display);
        controller.init();
    }
}
