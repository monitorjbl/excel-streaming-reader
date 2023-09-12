package com.github.pjfanning.xlsx.impl;

final class NumberUtil {

  static int parseInt(final String s) {
    return Integer.parseInt(s.trim());
  }

  static double parseDouble(final String s) {
    return Double.parseDouble(s.trim());
  }
}
