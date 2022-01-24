package org.excelautomation.Controller;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Properties;

public class Config {

    private final String EMPTY_TEMPLATE;
    private final String WORKING_TEMPLATE;
    private final boolean IS_ADDING;

    public Config(String EMPTY_TEMPLATE, String WORKING_TEMPLATE, boolean IS_ADDING) {
        this.EMPTY_TEMPLATE = EMPTY_TEMPLATE;
        this.WORKING_TEMPLATE = WORKING_TEMPLATE;
        this.IS_ADDING = IS_ADDING;
    }

    public String getEMPTY_TEMPLATE() {
        return EMPTY_TEMPLATE;
    }

    public String getWORKING_TEMPLATE() {
        return WORKING_TEMPLATE;
    }

    public boolean getIS_ADDING() {
        return IS_ADDING;
    }

    public static Config readConfig() {
        Properties properties = new Properties();
        File config = new File("./config.cfg");
        try (Reader reader = new InputStreamReader(new FileInputStream(config), StandardCharsets.UTF_8)) {
            properties.load(reader);
            if (!(properties.containsKey("EMPTY_TEMPLATE")
                    || properties.containsKey("WORKING_TEMPLATE")
                    || properties.containsKey("IS_ADDING"))) {
                throw new FileNotFoundException("ПАРАМЕТРЫ ЗАДАНЫ НЕВЕРНО");
            }
        } catch (FileNotFoundException e) {
            System.out.println(e.getMessage());
        } catch (IOException e) {
            e.printStackTrace();
        }

        return new Config(properties.getProperty("EMPTY_TEMPLATE")
                        , properties.getProperty("WORKING_TEMPLATE")
                        , !"0".equals(properties.getProperty("IS_ADDING")));
    }
}
