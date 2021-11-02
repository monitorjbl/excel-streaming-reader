package com.github.pjfanning.xlsx.impl.ooxml;

import java.io.Closeable;
import java.io.IOException;

public class ResourceWithTrackedCloseable<T> implements Closeable {
  private final T resource;
  private final ResourceCloser closeable;

  public ResourceWithTrackedCloseable(T resource, ResourceCloser closeable) {
    this.resource = resource;
    this.closeable = closeable;
  }

  public T getResource() {
    return resource;
  }

  @Override
  public void close() throws IOException {
    closeable.close();
  }
}
