package com.github.pjfanning.xlsx.impl;

import java.time.LocalDate;
import java.time.LocalDateTime;

class DateTimeUtil {
  static LocalDateTime parseDateTime(String dt) {
    try {
      return LocalDateTime.parse(dt);
    } catch (Exception e) {
      try {
        return LocalDate.parse(dt).atStartOfDay();
      } catch (Exception e2) {
        throw new IllegalStateException("Failed to parse `" + dt + "` as LocalDateTime");
      }
    }
  }
}
