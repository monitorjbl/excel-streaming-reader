package com.monitorjbl.xlsx;

import java.io.*;
import java.util.Iterator;
import java.util.Locale;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

public class StreamingSheetTest {
  @BeforeAll
  public static void init() {
    Locale.setDefault(Locale.ENGLISH);
  }

  @Test
  public void testLastRowNum() throws Exception {
    try(
        InputStream is = new FileInputStream(new File("src/test/resources/large.xlsx"));
        Workbook workbook = StreamingReader.builder().open(is);
    ) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      assertEquals(24, sheet.getLastRowNum());
    }

    try(
        InputStream is = new FileInputStream(new File("src/test/resources/empty_sheet.xlsx"));
        Workbook workbook = StreamingReader.builder().open(is);
    ) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      assertEquals(0, sheet.getLastRowNum());
    }
  }

  @Test
  public void testEmptyCellShouldHaveGeneralStyle() throws Exception {
    try(
        InputStream is = new FileInputStream(new File("src/test/resources/large.xlsx"));
        Workbook workbook = StreamingReader.builder().open(is);
    ) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      Row row = sheet.iterator().next();
      assertEquals(CellType.NUMERIC, row.getCell(0).getCellType());
      assertNotNull(row.getCell(0).getCellStyle());
    }
  }

  @Test
  public void testCellWithLineBreak() throws IOException {
    final String testValue = "1\n2\r\n3";
    try (
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
            ByteArrayOutputStream bos = new ByteArrayOutputStream()
    ) {
      Sheet xssfSheet = xssfWorkbook.createSheet();
      xssfSheet.createRow(0).createCell(0).setCellValue(testValue);

      xssfWorkbook.write(bos);

      try (Workbook wb = StreamingReader.builder().open(new ByteArrayInputStream(bos.toByteArray()))) {
        Sheet sheet = wb.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.rowIterator();
        if (rowIterator.hasNext()) {
          Row row = rowIterator.next();
          if (row.getRowNum() == 0) {
            Cell cell0 = row.getCell(0);
            assertNotNull(cell0);
            assertEquals(testValue, cell0.getStringCellValue());
          }
        }
      }
    }
  }

  @Test
  public void testCellWithLineBreakNoSharedStrings() throws IOException {
    //SXSSFWorkbook does not use SharedStrings by default
    final String testValue = "1\n2\r\n3";
    try (
            SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
            ByteArrayOutputStream bos = new ByteArrayOutputStream()
    ) {
      Sheet xssfSheet = sxssfWorkbook.createSheet();
      xssfSheet.createRow(0).createCell(0).setCellValue(testValue);

      sxssfWorkbook.write(bos);

      try (Workbook wb = StreamingReader.builder().open(new ByteArrayInputStream(bos.toByteArray()))) {
        Sheet sheet = wb.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.rowIterator();
        if (rowIterator.hasNext()) {
          Row row = rowIterator.next();
          if (row.getRowNum() == 0) {
            Cell cell0 = row.getCell(0);
            assertNotNull(cell0);
            assertEquals(testValue, cell0.getStringCellValue());
          }
        }
      }
    }
  }

}
