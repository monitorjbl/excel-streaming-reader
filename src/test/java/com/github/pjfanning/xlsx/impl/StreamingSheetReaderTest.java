package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.XMLHelper;
import org.junit.Assert;
import org.junit.Test;

import javax.xml.stream.XMLEventReader;
import java.io.FileInputStream;
import java.time.LocalDate;
import java.time.ZoneOffset;
import java.util.Date;

public class StreamingSheetReaderTest {
  @Test
  public void testStrictDates() throws Exception {
    XMLEventReader xer = XMLHelper.newXMLInputFactory().createXMLEventReader(
            new FileInputStream("src/test/resources/strict.dates.xml"));
    StreamingSheetReader reader = new StreamingSheetReader(null, null, null, xer, true, 100);
    try {
      Assert.assertEquals(0, reader.getFirstRowNum());
      Assert.assertEquals(0, reader.getLastRowNum());
      Row firstRow = reader.iterator().next();
      Assert.assertEquals(CellType.STRING, firstRow.getCell(0).getCellType());
      Assert.assertEquals("2021-02-28", firstRow.getCell(0).getStringCellValue());
      Assert.assertEquals(LocalDate.parse("2021-02-28").atStartOfDay(),
              firstRow.getCell(0).getLocalDateTimeCellValue());
      Assert.assertEquals(Date.from(LocalDate.parse("2021-02-28").atStartOfDay().toInstant(ZoneOffset.UTC)),
              firstRow.getCell(0).getDateCellValue());
      Assert.assertEquals(CellType.STRING, firstRow.getCell(1).getCellType());
      Assert.assertEquals("12:00:00.000", firstRow.getCell(1).getStringCellValue());
    } finally {
      xer.close();
      reader.close();
    }
  }
}
