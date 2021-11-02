package com.github.pjfanning.xlsx.impl.ooxml;

@FunctionalInterface
public interface ResourceCloser {
  void close();
}