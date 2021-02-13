package com.github.pjfanning.xlsx.impl.ooxml;

import java.io.*;

class TempMemoryDataStore implements TempDataStore {

  private final ByteArrayOutputStream bos = new ByteArrayOutputStream(4096);

  @Override
  public OutputStream getOutputStream() throws IOException {
    return bos;
  }

  @Override
  public InputStream getInputStream() throws IOException {
    return new ByteArrayInputStream(bos.toByteArray());
  }

  @Override
  public void close() throws IOException {
    if (bos != null) bos.close();
  }
}
