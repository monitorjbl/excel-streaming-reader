package com.github.pjfanning.xlsx.impl.ooxml;

import org.apache.poi.util.TempFile;

import java.io.*;

class TempFileDataStore implements TempDataStore {

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
    if (tempFile != null) tempFile.delete();
  }
}
