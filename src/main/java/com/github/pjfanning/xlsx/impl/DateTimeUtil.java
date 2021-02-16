package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.usermodel.DateUtil;

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

  static double convertTime(String input) {
    int dotIndex = input.lastIndexOf(".");
    if (dotIndex >= 0) {
      //POI DateUtil does not handle milliseconds in time
      return DateUtil.convertTime(input.substring(0, dotIndex));
    }
    return DateUtil.convertTime(input);
  }
}
