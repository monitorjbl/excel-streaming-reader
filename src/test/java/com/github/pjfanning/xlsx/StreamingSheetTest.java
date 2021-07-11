package com.github.pjfanning.xlsx;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Locale;
import java.util.NoSuchElementException;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

public class StreamingSheetTest {
  @BeforeClass
  public static void init() {
    Locale.setDefault(Locale.ENGLISH);
  }

  @Test
  public void testLastRowNum() throws Exception {
    try(
        InputStream is = new FileInputStream("src/test/resources/large.xlsx");
        Workbook workbook = StreamingReader.builder().open(is);
    ) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      assertEquals(0, sheet.getFirstRowNum());
      assertEquals(24, sheet.getLastRowNum());
    }

    try(
        InputStream is = new FileInputStream("src/test/resources/empty_sheet.xlsx");
        Workbook workbook = StreamingReader.builder().open(is);
    ) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      assertEquals(0, sheet.getFirstRowNum());
      assertEquals(0, sheet.getLastRowNum());
    }
  }

  @Test
  public void testEmptyCellShouldHaveGeneralStyle() throws Exception {
    try(
        InputStream is = new FileInputStream("src/test/resources/large.xlsx");
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
  public void testMergedRegion() throws IOException {
    try (UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream()) {
      try (XSSFWorkbook wb = new XSSFWorkbook()) {
        XSSFSheet sheet = wb.createSheet();
        CellRangeAddress region = new CellRangeAddress(1, 1, 1, 2);
        assertEquals(0, sheet.addMergedRegion(region));
        wb.write(bos);
      }
      try (Workbook workbook = StreamingReader.builder().open(bos.toInputStream())) {
        Sheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {
          //need to iterate over all rows before merged region data is read (it is at end of sheet data)
        }
        assertEquals(1, sheet.getMergedRegions().size());
        assertEquals(1, sheet.getNumMergedRegions());
      }
    }
  }

  @Test
  public void testRowIteratorNext() throws Exception {
    try(
            InputStream is = new FileInputStream("src/test/resources/large.xlsx");
            Workbook workbook = StreamingReader.builder().rowCacheSize(5).open(is);
    ) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> iter = sheet.rowIterator();
      int count = 0;
      while(nextRow(iter) != null) {
        count++;
      }
      assertEquals(25, count);
    }
  }

  private Row nextRow(Iterator<Row> iter) {
    try {
      return iter.next();
    } catch (NoSuchElementException nsee) {
      return null;
    }
  }

}
