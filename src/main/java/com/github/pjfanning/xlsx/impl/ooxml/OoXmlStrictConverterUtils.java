package com.github.pjfanning.xlsx.impl.ooxml;

import java.io.BufferedReader;
import java.io.FilterInputStream;
import java.io.FilterOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.util.Properties;

public class OoXmlStrictConverterUtils {

    public static boolean isBlank(final String str) {
        return str == null || str.trim().length() == 0;
    }

    public static boolean isNotBlank(final String str) {
        return !isBlank(str);
    }

    public static boolean isXml(final String fileName) {
        if(isNotBlank(fileName)) {
            int pos = fileName.lastIndexOf(".");
            if(pos != -1) {
                String ext = fileName.substring(pos + 1).toLowerCase();
                return ext.equals("xml") || ext.equals("vml") || ext.equals("rels");
            }
        }
        return false;
    }

    public static InputStream disableClose(InputStream inputStream) {
        return new FilterInputStream(inputStream) {
            @Override
            public void close() throws IOException {
            }
        };
    }

    public static OutputStream disableClose(OutputStream outputStream) {
        return new FilterOutputStream(outputStream) {
            @Override
            public void close() throws IOException {
            }
        };
    }

    public static Properties readMappings() {
        Properties props = new Properties();
        try(InputStream is = OoXmlStrictConverterUtils.class.getResourceAsStream("/ooxml-strict-mappings.properties");
                BufferedReader reader = new BufferedReader(new InputStreamReader(is, "ISO-8859-1"))) {
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
