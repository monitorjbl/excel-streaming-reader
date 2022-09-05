package com.github.pjfanning.xlsx;

import com.github.pjfanning.xlsx.impl.XlsxHyperlink;
import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

import static com.github.pjfanning.xlsx.TestUtils.getInputStream;
import static com.github.pjfanning.xlsx.TestUtils.nextRow;
import static org.junit.Assert.*;

public class StreamingSheetTest {
  private static Locale defaultLocale;

  @BeforeClass
  public static void init() {
    defaultLocale = Locale.getDefault();
    Locale.setDefault(Locale.ENGLISH);
  }

  @AfterClass
  public static void tearDown() {
    Locale.setDefault(defaultLocale);
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
      assertEquals(workbook, sheet.getWorkbook());
    }

    try(
        InputStream is = getInputStream("empty_sheet.xlsx");
        Workbook workbook = StreamingReader.builder().open(is);
    ) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      assertEquals(0, sheet.getFirstRowNum());
      assertEquals(0, sheet.getLastRowNum());
      assertEquals(workbook, sheet.getWorkbook());
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
            Workbook workbook = StreamingReader.builder().rowCacheSize(5).open(is)
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
            Workbook workbook = StreamingReader.builder().setReadHyperlinks(true).open(is)
    ) {
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();
      Cell a2 = null, a3 = null, a4 = null, a7 = null;
      while(rowIterator.hasNext()) {
        Row row = nextRow(rowIterator);
        if (row.getRowNum() == 1) {
          a2 = row.getCell(0);
        } else if (row.getRowNum() == 2) {
          a3 = row.getCell(0);
        } else if (row.getRowNum() == 3) {
          a4 = row.getCell(0);
        } else if (row.getRowNum() == 6) {
          a7 = row.getCell(0);
        }
      }
      assertNotNull( "a2 found", a2);
      assertNotNull( "a3 found", a3);
      assertNotNull( "a4 found", a4);
      assertNotNull( "a7 found", a7);
      assertEquals("http://twitter.com/#!/apacheorg", a2.getStringCellValue());
      assertEquals("http://www.bailii.org/databases.html#ie", a3.getStringCellValue());
      assertEquals("https://en.wikipedia.org/wiki/Apache_POI#See_also", a4.getStringCellValue());
      assertEquals("#Sheet1", a7.getStringCellValue());

      for (Row row : sheet) {
        //iterate again to make sure we don't duplicate the hyperlink data
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
      assertEquals("!/apacheorg", link1.getLocation());
      assertTrue(hps.contains(link1));

      XlsxHyperlink link2 = (XlsxHyperlink)sheet.getHyperlink(A3);
      assertEquals("A3", link2.getCellRef());
      assertEquals(HyperlinkType.URL, link2.getType());
      assertEquals("http://www.bailii.org/databases.html#ie", link2.getAddress());
      assertEquals("ie", link2.getLocation());
      assertTrue(hps.contains(link2));

      XlsxHyperlink link3 = (XlsxHyperlink)sheet.getHyperlink(A4);
      assertEquals("A4", link3.getCellRef());
      assertEquals(HyperlinkType.URL, link3.getType());
      assertEquals("https://en.wikipedia.org/wiki/Apache_POI#See_also", link3.getAddress());
      assertEquals("See_also", link3.getLocation());
      assertTrue(hps.contains(link3));

      XlsxHyperlink link4 = (XlsxHyperlink)sheet.getHyperlink(A7);
      assertEquals("A7", link4.getCellRef());
      assertEquals(HyperlinkType.DOCUMENT, link4.getType());
      assertEquals("Sheet1", link4.getAddress());
      assertEquals("Sheet1", link4.getLocation());
      assertTrue(hps.contains(link4));

      assertEquals(hps, sheet.getHyperlinkList());

      XlsxHyperlink link1a = (XlsxHyperlink) link1.copy();
      assertEquals(link1, link1a);
      assertEquals(link1.hashCode(), link1a.hashCode());

      XSSFHyperlink link1b = link1.createXSSFHyperlink();
      assertEquals(link1.getAddress(), link1b.getAddress() + "#" + link1b.getLocation());
      assertEquals(link1.getLocation(), link1b.getLocation());
      assertEquals(link1.getCellRef(), link1b.getCellRef());
      assertEquals(link1.getType(), link1b.getType());
      assertEquals(link1.getLabel(), link1b.getLabel());
      assertEquals(link1.getTooltip(), link1b.getTooltip());
    }
  }

  @Test
  public void testXSSFHyperlinks() throws Exception {
    try (
            InputStream is = getInputStream("59775.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(is)
    ) {
      XSSFSheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();
      Cell a2 = null, a3 = null, a4 = null, a7 = null;
      while(rowIterator.hasNext()) {
        Row row = nextRow(rowIterator);
        if (row.getRowNum() == 1) {
          a2 = row.getCell(0);
        } else if (row.getRowNum() == 2) {
          a3 = row.getCell(0);
        } else if (row.getRowNum() == 3) {
          a4 = row.getCell(0);
        } else if (row.getRowNum() == 6) {
          a7 = row.getCell(0);
        }
      }
      assertNotNull( "a2 found", a2);
      assertNotNull( "a3 found", a3);
      assertNotNull( "a4 found", a4);
      assertNotNull( "a7 found", a7);
      assertEquals("http://twitter.com/#!/apacheorg", a2.getStringCellValue());
      assertEquals("http://www.bailii.org/databases.html#ie", a3.getStringCellValue());
      assertEquals("https://en.wikipedia.org/wiki/Apache_POI#See_also", a4.getStringCellValue());
      assertEquals("#Sheet1", a7.getStringCellValue());

      for (Row row : sheet) {
        //iterate again to make sure we don't duplicate the hyperlink data
      }

      List<? extends Hyperlink> hps = sheet.getHyperlinkList();
      assertEquals(4, hps.size());

      CellAddress A2 = new CellAddress("A2");
      CellAddress A3 = new CellAddress("A3");
      CellAddress A4 = new CellAddress("A4");
      CellAddress A7 = new CellAddress("A7");

      XSSFHyperlink link1 = sheet.getHyperlink(A2);
      assertEquals("A2", link1.getCellRef());
      assertEquals(HyperlinkType.URL, link1.getType());
      assertEquals("http://twitter.com/#!/apacheorg", link1.getAddress());
      assertEquals("!/apacheorg", link1.getLocation());
      assertTrue(hps.contains(link1));

      XSSFHyperlink link2 = sheet.getHyperlink(A3);
      assertEquals("A3", link2.getCellRef());
      assertEquals(HyperlinkType.URL, link2.getType());
      assertEquals("http://www.bailii.org/databases.html#ie", link2.getAddress());
      assertEquals("ie", link2.getLocation());
      assertTrue(hps.contains(link2));

      XSSFHyperlink link3 = sheet.getHyperlink(A4);
      assertEquals("A4", link3.getCellRef());
      assertEquals(HyperlinkType.URL, link3.getType());
      assertEquals("https://en.wikipedia.org/wiki/Apache_POI#See_also", link3.getAddress());
      assertEquals("See_also", link3.getLocation());
      assertTrue(hps.contains(link3));

      XSSFHyperlink link4 = sheet.getHyperlink(A7);
      assertEquals("A7", link4.getCellRef());
      assertEquals(HyperlinkType.DOCUMENT, link4.getType());
      assertEquals("Sheet1", link4.getAddress());
      assertEquals("Sheet1", link4.getLocation());
      assertTrue(hps.contains(link4));

      assertEquals(hps, sheet.getHyperlinkList());
    }
  }


  @Test
  public void testHyperlinksDisabled() throws Exception {
    try (
            InputStream is = getInputStream("59775.xlsx");
            Workbook workbook = StreamingReader.builder().open(is)
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
  public void testSharedHyperlink() throws Exception {
    try (
            InputStream is = getInputStream("sharedhyperlink.xlsx");
            Workbook workbook = StreamingReader.builder().setReadHyperlinks(true).open(is)
    ) {
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();
      while(rowIterator.hasNext()) {
        nextRow(rowIterator);
      }

      Hyperlink hyperlink3 = sheet.getHyperlink(new CellAddress("A3"));
      Hyperlink hyperlink4 = sheet.getHyperlink(new CellAddress("A4"));
      Hyperlink hyperlink5 = sheet.getHyperlink(new CellAddress("A5"));
      assertNotNull("hyperlink found?", hyperlink3);
      assertEquals(hyperlink3, hyperlink4);
      assertEquals(hyperlink3, hyperlink5);
    }
  }

  @Test
  public void testGetActiveCell() throws Exception {
    try (
            InputStream is = getInputStream("59775.xlsx");
            Workbook workbook = StreamingReader.builder().open(is)
    ) {
      Sheet sheet = workbook.getSheetAt(0);
      nextRow(sheet.rowIterator()); //need to force a read of first row before getActiveCell works
      assertEquals(new CellAddress("A1"), sheet.getActiveCell());
    }
  }

  @Test
  public void testCustomWidthAndHeight() throws IOException {
    try (
            InputStream is = getInputStream("WidthsAndHeights.xlsx");
            Workbook wb = StreamingReader.builder().open(is)
    ) {
      Sheet sheet = wb.getSheetAt(0);
      Row row0 = null, row1 = null, row2 = null;
      for (Row row : sheet) {
        if (row.getRowNum() == 0) {
          row0 = row;
        } else if (row.getRowNum() == 1) {
          row1 = row;
        } else if (row.getRowNum() == 2) {
          row2 = row;
        }
      }
      assertEquals(8, sheet.getDefaultColumnWidth());
      assertEquals(300, sheet.getDefaultRowHeight());
      assertEquals(15.0, sheet.getDefaultRowHeightInPoints(), 0.00001);
      assertEquals(5120, sheet.getColumnWidth(0));
      assertEquals(2048, sheet.getColumnWidth(1));
      assertEquals(0, sheet.getColumnWidth(2));
      assertEquals(140.034, sheet.getColumnWidthInPixels(0), 0.00001);
      assertEquals(56.0136, sheet.getColumnWidthInPixels(1), 0.00001);
      assertEquals(0.0, sheet.getColumnWidthInPixels(2), 0.00001);
      assertFalse(sheet.isColumnHidden(0));
      assertFalse(sheet.isColumnHidden(1));
      assertTrue(sheet.isColumnHidden(2));
      assertNotNull(row0);
      assertNotNull(row1);
      assertNotNull(row2);
      assertEquals(750, row0.getHeight());
      assertEquals(37.5, row0.getHeightInPoints(), 0.00001);
      assertFalse(row0.getZeroHeight());
      assertEquals(300, row1.getHeight());
      assertEquals(15.0, row1.getHeightInPoints(), 0.00001);
      assertFalse(row1.getZeroHeight());
      assertEquals(15, row2.getHeight());
      assertEquals(0.75, row2.getHeightInPoints(), 0.00001);
      assertTrue(row2.getZeroHeight());
    }
  }

  @Test
  public void testPaneInformation() throws IOException {
    try (
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
            UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream()
    ) {
      int leftmostColumn = 3;
      int topRow = 4;

      Sheet s = xssfWorkbook.createSheet();

      // Populate
      for (int rn = 0; rn <= topRow; rn++) {
        Row r = s.createRow(rn);
        for (int cn = 0; cn < leftmostColumn; cn++) {
          Cell c = r.createCell(cn, CellType.NUMERIC);
          c.setCellValue(100 * rn + cn);
        }
      }

      // Now a column only freezepane
      s.createFreezePane(4, 0);
      PaneInformation paneInfo = s.getPaneInformation();

      assertEquals(4, paneInfo.getVerticalSplitPosition());
      assertEquals(0, paneInfo.getHorizontalSplitPosition());
      assertEquals(4, paneInfo.getVerticalSplitLeftColumn());
      assertEquals(0, paneInfo.getHorizontalSplitTopRow());
      assertTrue(paneInfo.isFreezePane());
      assertEquals(1, paneInfo.getActivePane());

      xssfWorkbook.write(bos);

      try (Workbook wb = StreamingReader.builder().open(bos.toInputStream())) {
        Sheet sheet = wb.getSheetAt(0);
        PaneInformation streamPane = sheet.getPaneInformation();
        assertEquals(4, streamPane.getVerticalSplitPosition());
        assertEquals(0, streamPane.getHorizontalSplitPosition());
        assertEquals(4, streamPane.getVerticalSplitLeftColumn());
        assertEquals(0, streamPane.getHorizontalSplitTopRow());
        assertTrue(streamPane.isFreezePane());
        assertEquals(1, streamPane.getActivePane());
        Iterator<Row> rowIterator = sheet.rowIterator();
        if (rowIterator.hasNext()) {
          rowIterator.next();
        }
      }
    }
  }

  @Test
  public void testRowStyle() throws IOException {
    try (
        InputStream is = getInputStream("row-style.xlsx");
        Workbook wb = StreamingReader.builder().open(is)
    ) {
      Sheet sheet = wb.getSheetAt(0);
      Row row0 = null, row1 = null;
      for (Row row : sheet) {
        if (row.getRowNum() == 0) {
          row0 = row;
        } else if (row.getRowNum() == 1) {
          row1 = row;
        }
      }
      assertNotNull(row0);
      assertNotNull(row1);
      assertFalse(row0.isFormatted());
      assertTrue(row1.isFormatted());
    }
  }

  @Test
  public void testCellWithLineBreak() throws IOException {
    final String testValue = "1\n2\r\n3";
    try (
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
            UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream()
    ) {
      Sheet xssfSheet = xssfWorkbook.createSheet();
      xssfSheet.createRow(0).createCell(0).setCellValue(testValue);

      xssfWorkbook.write(bos);

      try (Workbook wb = StreamingReader.builder().open(bos.toInputStream())) {
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
            UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream()
    ) {
      Sheet xssfSheet = sxssfWorkbook.createSheet();
      xssfSheet.createRow(0).createCell(0).setCellValue(testValue);

      sxssfWorkbook.write(bos);

      try (Workbook wb = StreamingReader.builder().open(bos.toInputStream())) {
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
