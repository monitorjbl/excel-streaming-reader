package com.github.pjfanning.xlsx.impl;

final class LazySupplier<T> implements Supplier {
  private T content;
  private final java.util.function.Supplier<T> functionalSupplier;

  LazySupplier(java.util.function.Supplier<T> functionalSupplier) {
    this.functionalSupplier = functionalSupplier;
  }

  @Override
  public Object getContent() {
    if (content == null) {
      content = functionalSupplier.get();
    }
    return content;
  }
}
