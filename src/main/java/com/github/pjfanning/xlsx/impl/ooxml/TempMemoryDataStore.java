package com.github.pjfanning.xlsx.impl.ooxml;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

class TempMemoryDataStore implements TempDataStore {

  private final UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream(4096);

  @Override
  public OutputStream getOutputStream() throws IOException {
    return bos;
  }

  @Override
  public InputStream getInputStream() throws IOException {
    return bos.toInputStream();
  }

  @Override
  public void close() throws IOException {
    if (bos != null) bos.close();
  }
}
