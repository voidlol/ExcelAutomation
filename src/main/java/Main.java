import org.excelautomation.Controller.Controller;
import org.excelautomation.View.ConsoleOutput;
import org.excelautomation.View.View;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class Main {

    public static void main(String[] args) {
        System.out.println("Excel Automation Version: " + new Main().getVersion());
        View display = new ConsoleOutput();
        Controller controller = new Controller(display);
        controller.init();
    }

    private String getVersion() {
        Properties p = new Properties();
        InputStream is = getClass().getResourceAsStream(".properties");
        try {
            p.load(is);
            return p.getProperty("version");
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
}
