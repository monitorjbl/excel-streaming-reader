package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.XMLHelper;
import org.junit.Assert;
import org.junit.Test;

import javax.xml.stream.XMLEventReader;
import java.io.FileInputStream;
import java.time.LocalDate;

public class StreamingSheetReaderTest {
  @Test
  public void testStrictDates() throws Exception {
    XMLEventReader xer = XMLHelper.newXMLInputFactory().createXMLEventReader(
            new FileInputStream("src/test/resources/strict.dates.xml"));
    StreamingSheetReader reader = new StreamingSheetReader(
            null, null, null, null, null, xer, true, 100);
    try {
      Assert.assertEquals(0, reader.getFirstRowNum());
      Assert.assertEquals(0, reader.getLastRowNum());
      Row firstRow = reader.iterator().next();
      Assert.assertEquals(CellType.NUMERIC, firstRow.getCell(0).getCellType());
      Assert.assertEquals("2021-02-28", firstRow.getCell(0).getStringCellValue());
      Assert.assertEquals(LocalDate.parse("2021-02-28").atStartOfDay(),
              firstRow.getCell(0).getLocalDateTimeCellValue());
      Assert.assertEquals(java.sql.Date.valueOf(LocalDate.parse("2021-02-28")),
              firstRow.getCell(0).getDateCellValue());
      Assert.assertEquals(CellType.NUMERIC, firstRow.getCell(1).getCellType());
      Assert.assertEquals("12:00:00.000", firstRow.getCell(1).getStringCellValue());
      Assert.assertEquals(DateUtil.convertTime("12:00"), firstRow.getCell(1).getNumericCellValue(), 0.001);
    } finally {
      xer.close();
      reader.close();
    }
  }
}
