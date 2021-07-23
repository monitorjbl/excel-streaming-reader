package com.github.pjfanning.xlsx;

import com.github.pjfanning.xlsx.impl.XlsxHyperlink;
import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

import static com.github.pjfanning.xlsx.TestUtils.*;
import static org.junit.Assert.*;

public class StreamingSheetTest {
  @BeforeClass
  public static void init() {
    Locale.setDefault(Locale.ENGLISH);
  }

  @Test
  public void testLastRowNum() throws Exception {
    try(
        InputStream is = getInputStream("large.xlsx");
        Workbook workbook = StreamingReader.builder().open(is);
    ) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      assertEquals(0, sheet.getFirstRowNum());
      assertEquals(24, sheet.getLastRowNum());
    }

    try(
        InputStream is = getInputStream("empty_sheet.xlsx");
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
        InputStream is = getInputStream("large.xlsx");
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
        for (Row row : sheet) {
          //iterate again to make sure we don't duplicate the merged region data
        }
        assertEquals(1, sheet.getMergedRegions().size());
        assertEquals(1, sheet.getNumMergedRegions());
      }
    }
  }

  @Test
  public void testRowIteratorNext() throws Exception {
    try(
            InputStream is = getInputStream("large.xlsx");
            Workbook workbook = StreamingReader.builder().rowCacheSize(5).open(is);
    ) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> iter = sheet.rowIterator();
      int count = 0;
      while(iter.hasNext()) {
        iter.next();
        count++;
      }
      assertEquals(25, count);
    }
  }

  @Test
  public void testHyperlinksEnabled() throws Exception {
    try (
            InputStream is = getInputStream("59775.xlsx");
            Workbook workbook = StreamingReader.builder().setReadHyperlinks(true).open(is);
    ) {
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();
      while(rowIterator.hasNext()) {
        nextRow(rowIterator);
        //ignore - just need to read through all rows
      }
      List<? extends Hyperlink> hps = sheet.getHyperlinkList();
      assertEquals(4, hps.size());

      CellAddress A2 = new CellAddress("A2");
      CellAddress A3 = new CellAddress("A3");
      CellAddress A4 = new CellAddress("A4");
      CellAddress A7 = new CellAddress("A7");

      XlsxHyperlink link1 = (XlsxHyperlink)sheet.getHyperlink(A2);
      assertEquals("A2", link1.getCellRef());
      assertEquals(HyperlinkType.URL, link1.getType());
      assertEquals("http://twitter.com/#!/apacheorg", link1.getAddress());
      assertTrue(hps.contains(link1));

      XlsxHyperlink link2 = (XlsxHyperlink)sheet.getHyperlink(A3);
      assertEquals("A3", link2.getCellRef());
      assertEquals(HyperlinkType.URL, link2.getType());
      assertEquals("http://www.bailii.org/databases.html#ie", link2.getAddress());
      assertTrue(hps.contains(link2));

      XlsxHyperlink link3 = (XlsxHyperlink)sheet.getHyperlink(A4);
      assertEquals("A4", link3.getCellRef());
      assertEquals(HyperlinkType.URL, link3.getType());
      assertEquals("https://en.wikipedia.org/wiki/Apache_POI#See_also", link3.getAddress());
      assertTrue(hps.contains(link3));

      XlsxHyperlink link4 = (XlsxHyperlink)sheet.getHyperlink(A7);
      assertEquals("A7", link4.getCellRef());
      assertEquals(HyperlinkType.DOCUMENT, link4.getType());
      assertEquals("Sheet1", link4.getAddress());
      assertTrue(hps.contains(link4));

      assertEquals(hps, sheet.getHyperlinkList());
    }
  }

  @Test
  public void testHyperlinksDisabled() throws Exception {
    try (
            InputStream is = getInputStream("59775.xlsx");
            Workbook workbook = StreamingReader.builder().open(is);
    ) {
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();
      while(rowIterator.hasNext()) {
        nextRow(rowIterator);
        //ignore - just need to read through all rows
      }
      try {
        sheet.getHyperlinkList();
        fail("expected IllegalStateException");
      } catch (IllegalStateException ise) {
        //expected
      }
    }
  }

  @Test
  public void testGetActiveCell() throws Exception {
    try (
            InputStream is = getInputStream("59775.xlsx");
            Workbook workbook = StreamingReader.builder().open(is);
    ) {
      Sheet sheet = workbook.getSheetAt(0);
      nextRow(sheet.rowIterator()); //need to force a read of first row before getActiveCell works
      assertEquals(new CellAddress("A1"), sheet.getActiveCell());
    }
  }
}
