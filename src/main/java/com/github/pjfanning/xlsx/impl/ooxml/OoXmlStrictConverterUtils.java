package com.github.pjfanning.xlsx.impl.ooxml;

import org.apache.poi.util.Beta;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Properties;

@Beta
public class OoXmlStrictConverterUtils {

    private OoXmlStrictConverterUtils() {}

    public static boolean isBlank(final String str) {
        return str == null || str.trim().length() == 0;
    }

    public static boolean isNotBlank(final String str) {
        return !isBlank(str);
    }

    public static Properties readMappings() {
        Properties props = new Properties();
        try(InputStream is = OoXmlStrictConverterUtils.class.getResourceAsStream("/ooxml-strict-mappings.properties");
                BufferedReader reader = new BufferedReader(new InputStreamReader(is, StandardCharsets.ISO_8859_1))) {
            String line;
            while((line = reader.readLine()) != null) {
                String[] vals = line.split("=");
                if(vals.length >= 2) {
                    props.setProperty(vals[0], vals[1]);
                } else if(vals.length == 1) {
                    props.setProperty(vals[0], "");
                }

            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return props;
    }

}
