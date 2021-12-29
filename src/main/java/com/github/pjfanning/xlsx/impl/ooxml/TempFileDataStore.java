package com.github.pjfanning.xlsx.impl.ooxml;

import org.apache.poi.util.TempFile;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;

class TempFileDataStore implements TempDataStore {

  private static final Logger log = LoggerFactory.getLogger(TempFileDataStore.class);
  private File tempFile;

  @Override
  public OutputStream getOutputStream() throws IOException {
    if (tempFile != null) {
      throw new IOException("temp file already created");
    }
    tempFile = TempFile.createTempFile("excel-streaming-reader", ".xml");
    return new FileOutputStream(tempFile);
  }

  @Override
  public InputStream getInputStream() throws IOException {
    if (tempFile == null) {
      throw new IOException("temp file was never populated");
    }
    return new FileInputStream(tempFile);
  }

  @Override
  public void close() throws IOException {
    if (tempFile != null && !tempFile.delete()) {
      log.debug("failed to delete temp file");
    }
  }
}
