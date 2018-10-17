package com.monitorjbl.xlsx.impl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;

public class TempFileUtil {
    public static File writeInputStreamToFile(InputStream is, int bufferSize) throws IOException {
        File f = Files.createTempFile("tmp-", ".xlsx").toFile();
        try(FileOutputStream fos = new FileOutputStream(f)) {
            int read;
            byte[] bytes = new byte[bufferSize];
            while((read = is.read(bytes)) != -1) {
                fos.write(bytes, 0, read);
            }
            return f;
        } finally {
            is.close();
        }
    }
}