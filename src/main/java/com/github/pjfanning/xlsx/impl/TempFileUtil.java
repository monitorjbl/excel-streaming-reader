package com.github.pjfanning.xlsx.impl;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;

public class TempFileUtil {
  private static final Logger log = LoggerFactory.getLogger(TempFileUtil.class);

  public static File writeInputStreamToFile(InputStream is, int bufferSize) throws IOException {
    if (is == null) throw new NullPointerException("InputStream is null");
    File f = Files.createTempFile("tmp-", ".xlsx").toFile();
    try (FileOutputStream fos = new FileOutputStream(f)) {
      int read;
      byte[] bytes = new byte[bufferSize];
      while ((read = is.read(bytes)) != -1) {
        fos.write(bytes, 0, read);
      }
      return f;
    } catch (IOException | RuntimeException | Error e) {
      try {
        f.delete();
      } catch (Throwable t) {
        log.warn("Failed to delete temp file {}: {}", f.getAbsolutePath(), t.toString());
      }
      throw e;
    } finally {
      is.close();
    }
  }
}
