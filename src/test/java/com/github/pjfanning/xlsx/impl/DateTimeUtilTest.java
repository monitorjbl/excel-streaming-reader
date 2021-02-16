package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.usermodel.DateUtil;
import org.junit.Assert;
import org.junit.Test;

import java.time.LocalDate;

public class DateTimeUtilTest {
  @Test
  public void testConvertTime() {
    Assert.assertEquals(DateUtil.convertTime("12:00"),
            DateTimeUtil.convertTime("12:00:00.000"), 0.001);
  }

  @Test
  public void testParse() {
    Assert.assertEquals(LocalDate.parse("2021-02-28").atStartOfDay(), DateTimeUtil.parseDateTime("2021-02-28"));
  }
}
