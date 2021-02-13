package com.github.pjfanning.xlsx.impl.ooxml;

import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Use <code>getOutputStream</code> and insert data using the stream. Read the data back using
 * <code>getOutputStream</code>. The implementations are not designed to be thread-safe.
 */
interface TempDataStore extends Closeable {
  OutputStream getOutputStream() throws IOException;
  InputStream getInputStream() throws IOException;
}
