package com.github.pjfanning.xlsx;

import com.github.pjfanning.xlsx.exceptions.OpenException;
import com.github.pjfanning.xlsx.exceptions.ParseException;
import com.github.pjfanning.xlsx.impl.StreamingWorkbookReader;
import com.github.pjfanning.xlsx.impl.XlsxPictureData;
import fi.iki.elonen.NanoHTTPD;
import org.apache.commons.io.IOUtils;
import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.openxml4j.opc.ZipPackage;
import org.apache.poi.openxml4j.util.ZipInputStreamZipEntrySource;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;
import org.xml.sax.SAXException;

import java.io.*;
import java.util.*;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Consumer;

import static com.github.pjfanning.xlsx.TestUtils.*;
import static org.apache.poi.ss.usermodel.CellType.*;
import static org.junit.Assert.*;

public class StreamingWorkbookTest {
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
  public void testIterateSheets() throws Exception {
    try(
            InputStream is = new FileInputStream("src/test/resources/sheets.xlsx");
            Workbook workbook = StreamingReader.builder().open(is);
    ) {
      testIteration(workbook);
    }
  }

  @Test
  public void testIterateSheetsUsingAvoidTempFiles() throws Exception {
    StreamingReader.Builder builder = StreamingReader.builder().setAvoidTempFiles(true);
    try(
            InputStream is = new FileInputStream("src/test/resources/sheets.xlsx");
            Workbook workbook = builder.open(is);
    ) {
      testIteration(workbook);
    }
  }

  @Test
  public void testIterateSheetsUsingFile() throws Exception {
    try(Workbook workbook = StreamingReader.builder().open(new File("src/test/resources/sheets.xlsx"))) {
      testIteration(workbook);
    }
  }

  private void testIteration(Workbook workbook) {
    assertEquals(2, workbook.getNumberOfSheets());

    Sheet alpha = workbook.getSheetAt(0);
    Sheet zulu = workbook.getSheetAt(1);
    assertEquals("SheetAlpha", alpha.getSheetName());
    assertEquals("SheetZulu", zulu.getSheetName());

    Row rowA = alpha.rowIterator().next();
    Row rowZ = zulu.rowIterator().next();

    assertEquals("stuff", rowA.getCell(0).getStringCellValue());
    assertEquals("yeah", rowZ.getCell(0).getStringCellValue());
  }

  @Test
  public void testMissingCells() throws Exception {
    try(UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream()) {
      try(XSSFWorkbook xssfWorkbook = new XSSFWorkbook()) {
        Sheet sheet = xssfWorkbook.createSheet();
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(10);
        cell.setCellValue("value123");
        xssfWorkbook.write(bos);
      }
      try(Workbook workbook = StreamingReader.builder().open(bos.toInputStream())) {
        assertEquals(1, workbook.getNumberOfSheets());
        Sheet sheet = workbook.getSheetAt(0);
        int rowCount = 0;
        for (Row row : sheet) {
          rowCount++;
          assertEquals("rowNum matches", 0, row.getRowNum());
          int cellCount = 0;
          for(Cell cell : row) {
            cellCount++;
            assertEquals("column matches", 10, cell.getColumnIndex());
            assertEquals("value123", cell.getStringCellValue());
          }
          assertEquals("cellCount matches", 1, cellCount);
          Cell cell0 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
          assertEquals("cell0 column matches", 0, cell0.getColumnIndex());
          assertEquals("cell0 cell type", BLANK, cell0.getCellType());

          assertNull("null cell when no MissingCellPolicy?", row.getCell(0));
        }
        assertEquals("rowCount matches", 1, rowCount);
      }
    }
  }

  @Test
  public void testHiddenCells() throws Exception {
    try(
            InputStream is = new FileInputStream("src/test/resources/hidden.xlsx");
            Workbook workbook = StreamingReader.builder().open(is)
    ) {
      assertEquals(3, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);

      assertFalse("Column 0 should not be hidden", sheet.isColumnHidden(0));
      assertTrue("Column 1 should be hidden", sheet.isColumnHidden(1));
      assertFalse("Column 2 should not be hidden", sheet.isColumnHidden(2));

      Iterator<Row> rowIterator = sheet.rowIterator();

      assertFalse("Row 0 should not be hidden", rowIterator.next().getZeroHeight());
      assertTrue("Row 1 should be hidden", rowIterator.next().getZeroHeight());
      assertFalse("Row 2 should not be hidden", rowIterator.next().getZeroHeight());
    }
  }

  @Test
  public void testHiddenSheets() throws Exception {
    try(
            InputStream is = new FileInputStream("src/test/resources/hidden.xlsx");
            Workbook workbook = StreamingReader.builder().open(is)
    ) {
      assertEquals(3, workbook.getNumberOfSheets());
      assertFalse(workbook.isSheetHidden(0));

      assertTrue(workbook.isSheetHidden(1));
      assertFalse(workbook.isSheetVeryHidden(1));

      assertFalse(workbook.isSheetHidden(2));
      assertTrue(workbook.isSheetVeryHidden(2));
    }
  }

  @Test
  public void testFormulaCells() throws Exception {
    try(Workbook workbook = openWorkbook("formula_cell.xlsx")) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);

      Iterator<Row> rowIterator = sheet.rowIterator();
      Cell A1 = getCellFromNextRow(rowIterator, 0);
      Cell A2 = getCellFromNextRow(rowIterator, 0);
      Cell A3 = getCellFromNextRow(rowIterator, 0);

      expectType(A3, FORMULA);
      expectCachedType(A3, NUMERIC);
      expectFormula(A3, "SUM(A1:A2)");

      expectStringContent(A1, "1");
      expectStringContent(A2, "2");
      expectStringContent(A3, "3");
    }
  }

  @Test
  public void testNumericFormattedFormulaCell() throws Exception {
    try(Workbook workbook = openWorkbook("formula_cell.xlsx")) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();

      Cell C1 = getCellFromNextRow(rowIterator, 2);
      Cell C2 = getCellFromNextRow(rowIterator, 2);

      expectType(C2, FORMULA);
      expectCachedType(C2, NUMERIC);
      expectFormula(C2, "C1");
      expectSameStringContent(C2, C1);
      expectStringContent(C2, "May 11 2018");
    }
  }

  @Test
  public void testBooleanFormattedFormulaCell() throws Exception {
    try(Workbook workbook = openWorkbook("formula_cell.xlsx")) {
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();

      Cell D1 = getCellFromNextRow(rowIterator, 3);
      Cell D2 = getCellFromNextRow(rowIterator, 3);

      expectType(D1, FORMULA);
      expectCachedType(D1, BOOLEAN);
      assertTrue(D1.getBooleanCellValue());

      expectType(D2, FORMULA);
      expectCachedType(D2, BOOLEAN);
      assertFalse(D2.getBooleanCellValue());

    }
  }

  @Test
  public void testStringFormattedFormulaCell() throws Exception {
    try(Workbook workbook = openWorkbook("formula_cell.xlsx")) {
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();

      Cell B1 = getCellFromNextRow(rowIterator, 1);
      nextRow(rowIterator);
      Cell B3 = getCellFromNextRow(rowIterator, 1);

      expectType(B3, FORMULA);
      expectCachedType(B3, STRING);
      expectFormula(B3, "B1");
      expectSameStringContent(B1, B3);
      expectStringContent(B3, "a");
      expectRichStringContent(B3, "a");
    }
  }

  @Test
  public void testQuotedStringFormattedFormulaCell() throws Exception {
    try(Workbook workbook = openWorkbook("formula_cell.xlsx")) {
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();

      nextRow(rowIterator);
      Cell B2 = getCellFromNextRow(rowIterator, 1);
      nextRow(rowIterator);
      Cell B4 = getCellFromNextRow(rowIterator, 1);

      expectType(B4, FORMULA);
      expectCachedType(B4, STRING);
      try {
        expectFormula(B4, "B2");
      } catch (IllegalStateException ise) {
        //expected
      }
      expectSameStringContent(B2, B4);
      expectStringContent(B4, "\"a\"");
    }
  }

  @Test
  public void testQuotedStringFormattedFormulaCellWithSharedFormulaSupport() throws Exception {
    try (
            InputStream stream = getInputStream("formula_cell.xlsx");
            Workbook workbook = StreamingReader.builder().setReadSharedFormulas(true).open(stream)
    ) {
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();

      nextRow(rowIterator);
      Cell B2 = getCellFromNextRow(rowIterator, 1);
      nextRow(rowIterator);
      Cell B4 = getCellFromNextRow(rowIterator, 1);

      expectType(B4, FORMULA);
      expectCachedType(B4, STRING);
      expectFormula(B4, "B2"); // this only works if setReadSharedFormulas(true) set on builder
      expectSameStringContent(B2, B4);
      expectStringContent(B4, "\"a\"");
    }
  }

  @Test
  public void testInlineString() throws Exception {
    //https://bz.apache.org/bugzilla/show_bug.cgi?id=65096
    try(Workbook workbook = openWorkbook("InlineString.xlsx")) {
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();

      Cell A1 = getCellFromNextRow(rowIterator, 0);

      expectType(A1, STRING);
      expectStringContent(A1, "\uD83D\uDE1Cmore text");
      expectRichStringContent(A1, "\uD83D\uDE1Cmore text");
    }
  }

  @Test
  public void testMissingRattrs() throws Exception {
    try(Workbook workbook = openWorkbook("missing-r-attrs.xlsx")) {
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();
      Row row = rowIterator.next();
      assertEquals(0, row.getRowNum());
      assertEquals("1", row.getCell(0).getStringCellValue());
      assertEquals("5", row.getCell(4).getStringCellValue());
      row = rowIterator.next();
      assertEquals(1, row.getRowNum());
      assertEquals("6", row.getCell(0).getStringCellValue());
      assertEquals("10", row.getCell(4).getStringCellValue());
      row = rowIterator.next();
      assertEquals(6, row.getRowNum());
      assertEquals("11", row.getCell(0).getStringCellValue());
      assertEquals("15", row.getCell(4).getStringCellValue());

      assertFalse(rowIterator.hasNext());
    }
  }

  @Test
  public void testGetAllPicturesWorksWithNoPictures() throws Exception {
    try (Workbook workbook = openWorkbook("missing-r-attrs.xlsx")) {
      List<? extends PictureData> pictureList = workbook.getAllPictures();
      assertEquals(0, pictureList.size());
    }
  }

  @Test
  public void testGetPicturesWithoutReadShapesEnabled() throws Exception {
    try (Workbook workbook = openWorkbook("WithDrawing.xlsx")) {
      List<? extends PictureData> pictureList = workbook.getAllPictures();
      assertEquals(5, pictureList.size());
      for(PictureData picture : pictureList) {
        XlsxPictureData xlsxPictureData = (XlsxPictureData)picture;
        assertTrue("picture data is not empty", picture.getData().length > 0);
        assertArrayEquals(picture.getData(), IOUtils.toByteArray(xlsxPictureData.getInputStream()));
      }
      Sheet sheet0 = workbook.getSheetAt(0);
      sheet0.rowIterator().hasNext();
      try {
        sheet0.getDrawingPatriarch();
        fail("getDrawingPatriarch expected to fail with IllegalStateException");
      } catch (IllegalStateException ise) {
        //expected
      }
    }
  }

  @Test
  public void testGetPicturesWithReadShapesEnabled() throws Exception {
    try (
            InputStream stream = getInputStream("WithDrawing.xlsx");
            Workbook workbook = StreamingReader.builder().setReadShapes(true).open(stream)
    ) {
      List<? extends PictureData> pictureList = workbook.getAllPictures();
      assertEquals(5, pictureList.size());
      for(PictureData picture : pictureList) {
        XlsxPictureData xlsxPictureData = (XlsxPictureData)picture;
        assertTrue("picture data is not empty", picture.getData().length > 0);
        assertArrayEquals(picture.getData(), IOUtils.toByteArray(xlsxPictureData.getInputStream()));
      }
      Sheet sheet0 = workbook.getSheetAt(0);
      sheet0.rowIterator().hasNext();
      Drawing<?> drawingPatriarch = sheet0.getDrawingPatriarch();
      assertNotNull("drawingPatriarch should not be null", drawingPatriarch);
      List<XSSFPicture> pictures = new ArrayList<>();
      for (Shape shape : drawingPatriarch) {
        if (shape instanceof XSSFPicture) {
          pictures.add((XSSFPicture)shape);
        } else {
          //there is one text box and 5 pictures on the sheet
          XSSFSimpleShape textBox = (XSSFSimpleShape)shape;
          String text = textBox.getText().replace("\r", "").replace("\n", "");
          assertEquals("Sheet with various pictures(jpeg, png, wmf, emf and pict)", text);
        }
        assertTrue("shape is an XSSFShape", shape instanceof XSSFShape);
        assertNotNull("shape has anchor", shape.getAnchor());
      }
      assertEquals(5, pictures.size());
      Sheet sheet1 = workbook.getSheetAt(1);
      sheet1.rowIterator().hasNext();
      assertNull("sheet1 should have no drawing patriarch", sheet1.getDrawingPatriarch());
    }
  }

  @Test(expected = OpenException.class)
  public void testEntityExpansionWithPoiDefaultSst() throws Exception {
    ExploitServer.withServer(s -> fail("Should not have made request"), () -> {
      try(
              InputStream stream = getInputStream("entity-expansion-exploit-poc-file.xlsx");
              Workbook workbook = StreamingReader.builder()
                      .setSharedStringsImplementationType(SharedStringsImplementationType.POI_DEFAULT)
                      .open(stream)
      ) {
        Sheet sheet = workbook.getSheetAt(0);
        for(Row row : sheet) {
          for(Cell cell : row) {
            System.out.println(cell.getStringCellValue());
          }
        }
      } catch(IOException e) {
        throw new UncheckedIOException(e);
      }
    });
  }

  @Test(expected = OpenException.class)
  public void testEntityExpansionWithReadOnlySst() throws Exception {
    ExploitServer.withServer(s -> fail("Should not have made request"), () -> {
      try (
              InputStream stream = getInputStream("entity-expansion-exploit-poc-file.xlsx");
              Workbook workbook = StreamingReader.builder().setUseSstReadOnly(true)
                      .open(stream)
      ) {
        Sheet sheet = workbook.getSheetAt(0);
        for(Row row : sheet) {
          for(Cell cell : row) {
            System.out.println(cell.getStringCellValue());
          }
        }
      } catch(IOException e) {
        throw new UncheckedIOException(e);
      }
    });
  }

  @Test
  public void testDataFormatter() throws IOException {
    try(Workbook workbook = openWorkbook("formats.xlsx")) {
      Sheet sheet = workbook.getSheetAt(0);
      validateFormatsSheet(sheet);
    }
  }

  @Test
  public void testWithTempFileZipInputStream() throws IOException {
    //this test cannot be run in parallel with other tests because it changes static configs
    ZipInputStreamZipEntrySource.setThresholdBytesForTempFiles(0);
    try(Workbook workbook = openWorkbook("formats.xlsx")) {
      Sheet sheet = workbook.getSheetAt(0);
      validateFormatsSheet(sheet);
    } finally {
      ZipInputStreamZipEntrySource.setThresholdBytesForTempFiles(-1);
    }
  }

  @Test
  public void testWithTempFileZipPackage() throws IOException {
    //this test cannot be run in parallel with other tests because it changes static configs
    ZipPackage.setUseTempFilePackageParts(true);
    try(Workbook workbook = openWorkbook("formats.xlsx")) {
      Sheet sheet = workbook.getSheetAt(0);
      validateFormatsSheet(sheet);
    } finally {
      ZipPackage.setUseTempFilePackageParts(false);
    }
  }

  @Test
  public void testCallingSheetIteratorTwice() throws IOException {
    try(Workbook workbook = openWorkbook("formats.xlsx")) {
      Iterator<Sheet> iter1 = workbook.sheetIterator();
      List<Sheet> sheetList1 = new ArrayList<>();
      iter1.forEachRemaining(sheetList1::add);
      assertEquals(1, sheetList1.size());
      validateFormatsSheet(sheetList1.get(0));

      Iterator<Sheet> iter2 = workbook.sheetIterator();
      List<Sheet> sheetList2 = new ArrayList<>();
      iter2.forEachRemaining(sheetList2::add);
      assertEquals(1, sheetList2.size());
      assertEquals(sheetList1.get(0).hashCode(), sheetList2.get(0).hashCode());
      validateFormatsSheet(sheetList2.get(0));
    }
  }

  @Test
  public void testCallingSheetSpliteratorTwice() throws IOException {
    try(Workbook workbook = openWorkbook("formats.xlsx")) {
      Spliterator<Sheet> iter1 = workbook.spliterator();
      List<Sheet> sheetList1 = new ArrayList<>();
      iter1.forEachRemaining(sheetList1::add);
      assertEquals(1, sheetList1.size());
      validateFormatsSheet(sheetList1.get(0));

      Spliterator<Sheet> iter2 = workbook.spliterator();
      List<Sheet> sheetList2 = new ArrayList<>();
      iter2.forEachRemaining(sheetList2::add);
      assertEquals(1, sheetList2.size());
      assertEquals(sheetList1.get(0).hashCode(), sheetList2.get(0).hashCode());
      validateFormatsSheet(sheetList2.get(0));
    }
  }

  @Test
  public void testWithGetSheetAtAndThenIterator() throws IOException {
    try(Workbook workbook = openWorkbook("formats.xlsx")) {
      Sheet sheet = workbook.getSheetAt(0);
      validateFormatsSheet(sheet);
      Iterator<Sheet> iter1 = workbook.sheetIterator();
      List<Sheet> sheetList1 = new ArrayList<>();
      iter1.forEachRemaining(sheetList1::add);
      assertEquals(1, sheetList1.size());
      assertEquals(sheet.hashCode(), sheetList1.get(0).hashCode());
      validateFormatsSheet(sheetList1.get(0));
    }
  }

  @Test
  public void testNoSuchElementExceptionOnSheetIterator() throws IOException {
    try (Workbook workbook = openWorkbook("formats.xlsx")) {
      Iterator<Sheet> iter1 = workbook.sheetIterator();
      assertTrue(iter1.hasNext());
      assertNotNull(iter1.next());
      assertFalse(iter1.hasNext());
      assertThrows(NoSuchElementException.class, () -> iter1.next());
    }
  }

  @Test
  public void testRightToLeft() throws IOException {
    try(
            InputStream stream = getInputStream("right-to-left.xlsx");
            Workbook workbook = StreamingReader.builder()
                    .setReadComments(true)
                    .setAdjustLegacyComments(true)
                    .open(stream)
    ){
      Sheet sheet = workbook.getSheet("عربى");
      Iterator<Row> rowIterator = sheet.rowIterator();

      Cell A1 = getCellFromNextRow(rowIterator, 0);
      Cell A2 = getCellFromNextRow(rowIterator, 0);
      Cell A3 = getCellFromNextRow(rowIterator, 0);
      Cell A4 = getCellFromNextRow(rowIterator, 0);

      expectFormattedContent(A1, "نص");
      expectFormattedContent(A2, "123"); //this should really be ۱۲۳
      expectFormattedContent(A3, "text with comment");
      expectFormattedContent(A4, " עִבְרִית and اَلْعَرَبِيَّةُ");

      Comment a3Comment = sheet.getCellComment(new CellAddress("A3"));
      assertEquals(0, a3Comment.getString().numFormattingRuns());
      assertEquals("تعليق الاختبار", a3Comment.getString().getString());
    }
  }

  @Test
  public void testRightToLeftMapBackedSst() throws IOException {
    try(
            InputStream stream = getInputStream("right-to-left.xlsx");
            Workbook workbook = StreamingReader.builder()
                    .setSharedStringsImplementationType(SharedStringsImplementationType.CUSTOM_MAP_BACKED)
                    .setReadComments(true)
                    .setAdjustLegacyComments(true)
                    .open(stream)
    ){
      Sheet sheet = workbook.getSheet("عربى");
      Iterator<Row> rowIterator = sheet.rowIterator();

      Cell A1 = getCellFromNextRow(rowIterator, 0);
      Cell A2 = getCellFromNextRow(rowIterator, 0);
      Cell A3 = getCellFromNextRow(rowIterator, 0);
      Cell A4 = getCellFromNextRow(rowIterator, 0);

      expectFormattedContent(A1, "نص");
      expectFormattedContent(A2, "123"); //this should really be ۱۲۳
      expectFormattedContent(A3, "text with comment");
      expectFormattedContent(A4, " עִבְרִית and اَلْعَرَبِيَّةُ");

      Comment a3Comment = sheet.getCellComment(new CellAddress("A3"));
      assertEquals(0, a3Comment.getString().numFormattingRuns());
      assertEquals("تعليق الاختبار", a3Comment.getString().getString());
    }
  }

  @Test
  public void testAdjustLegacyCommentsDisabled() throws IOException {
    try(
            InputStream stream = getInputStream("right-to-left.xlsx");
            Workbook workbook = StreamingReader.builder().setReadComments(true).open(stream)
    ){
      Sheet sheet = workbook.getSheet("عربى");
      Iterator<Row> rowIterator = sheet.rowIterator();

      Comment a3Comment = sheet.getCellComment(new CellAddress("A3"));
      String expectedComment = "تعليق الاختبار";
      assertNotEquals(expectedComment, a3Comment.getString().getString());
      assertTrue("legacy comment ends with expected comment?", a3Comment.getString().getString().endsWith(expectedComment));
    }
  }

  @Test
  public void testGetErrorCellValue() throws IOException {
    try(UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream()) {
      try(XSSFWorkbook workbook = new XSSFWorkbook()) {
        XSSFSheet sheet = workbook.createSheet("sheet1");
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell0 = row.createCell(0);
        cell0.setCellValue("");
        XSSFCell cell1 = row.createCell(1);
        cell1.setCellErrorValue(FormulaError.DIV0);
        XSSFCell cell2 = row.createCell(2);
        cell2.setCellErrorValue(FormulaError.FUNCTION_NOT_IMPLEMENTED);
        workbook.write(bos);
      }
      try(Workbook wb = StreamingReader.builder().open(bos.toInputStream())) {
        Sheet sheet = wb.getSheet("sheet1");
        Row row0 = sheet.rowIterator().next();
        try {
          row0.getCell(0).getErrorCellValue();
          fail("expected IllegalStateException");
        } catch (IllegalStateException re) {
          //expected
        }
        assertEquals(FormulaError.DIV0.getCode(), row0.getCell(1).getErrorCellValue());
        assertEquals(FormulaError.FUNCTION_NOT_IMPLEMENTED.getCode(), row0.getCell(2).getErrorCellValue());
      }
    }
  }

  @Test
  public void testBug65676() throws Exception {
    try (UnsynchronizedByteArrayOutputStream output = new UnsynchronizedByteArrayOutputStream()) {
      try(Workbook wb = new SXSSFWorkbook()) {
        Row r = wb.createSheet("Sheet").createRow(0);
        r.createCell(0).setCellValue(1.2); /* A1: Number 1.2 */
        r.createCell(1).setCellValue("ABC"); /* B1: Inline string "ABC" */
        wb.write(output);
      }
      try(Workbook wb = StreamingReader.builder().open(output.toInputStream())) {
        Sheet sheet = wb.getSheet("Sheet");
        Cell a1 = null;
        Cell b1 = null;
        for (Row row : sheet) {
          if (row.getRowNum() == 0) {
            a1 = row.getCell(0);
            b1 = row.getCell(1);
          }
        }
        assertNotNull("a1 should be found", a1);
        assertNotNull("b1 should be found", b1);
        assertEquals(1.2, a1.getNumericCellValue(), 0.00000001);
        assertEquals("ABC", b1.getStringCellValue());
      }
    }
  }

  @Test
  public void testSheetNameCaseInsensitivity() throws IOException {
    final String sheetName1 = "sheetWithCamelCaseName";
    final String sheetName2 = "SHEET_WITH_CAPS_NAME";
    try (
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
            UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream()
    ) {
      XSSFSheet xssfSheet1 = xssfWorkbook.createSheet(sheetName1);
      xssfSheet1.createRow(0).createCell(0).setCellValue(sheetName1);
      XSSFSheet xssfSheet2 = xssfWorkbook.createSheet(sheetName2);
      xssfSheet2.createRow(0).createCell(0).setCellValue(sheetName2);
      xssfWorkbook.write(bos);
      try (Workbook wb = StreamingReader.builder().open(bos.toInputStream())) {
        Sheet sheet1 = wb.getSheetAt(0);
        Sheet sheet1a = wb.getSheet(sheetName1);
        Sheet sheet1b = wb.getSheet(sheetName1.toLowerCase(Locale.ROOT));
        Sheet sheet1c = wb.getSheet(sheetName1.toUpperCase(Locale.ROOT));
        assertNotNull(sheet1);
        assertEquals(sheet1, sheet1a);
        assertEquals(sheet1, sheet1b);
        assertEquals(sheet1, sheet1c);
        assertEquals(sheetName1, sheet1a.getSheetName());
        assertEquals(sheetName1, sheet1a.rowIterator().next().getCell(0).getStringCellValue());
        assertEquals(0, wb.getSheetIndex(sheet1c));
        assertEquals(0, wb.getSheetIndex(sheetName1));
        assertEquals(0, wb.getSheetIndex(sheetName1.toLowerCase(Locale.ROOT)));
        assertEquals(0, wb.getSheetIndex(sheetName1.toUpperCase(Locale.ROOT)));

        Sheet sheet2 = wb.getSheetAt(1);
        Sheet sheet2a = wb.getSheet(sheetName2);
        Sheet sheet2b = wb.getSheet(sheetName2.toLowerCase(Locale.ROOT));
        Sheet sheet2c = wb.getSheet(sheetName2.toUpperCase(Locale.ROOT));
        assertNotNull(sheet2);
        assertNotEquals(sheet1a, sheet2a);
        assertEquals(sheet2, sheet2a);
        assertEquals(sheet2, sheet2b);
        assertEquals(sheet2, sheet2c);
        assertEquals(sheetName2, sheet2a.getSheetName());
        assertEquals(sheetName2, sheet2a.rowIterator().next().getCell(0).getStringCellValue());
        assertEquals(1, wb.getSheetIndex(sheet2c));
        assertEquals(1, wb.getSheetIndex(sheetName2));
        assertEquals(1, wb.getSheetIndex(sheetName2.toLowerCase(Locale.ROOT)));
        assertEquals(1, wb.getSheetIndex(sheetName2.toUpperCase(Locale.ROOT)));
      }
    }
  }

  @Test
  public void testConcurrentSheetRead() throws Exception {
    try (
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
            UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream()
    ) {
      Sheet sheet1 = xssfWorkbook.createSheet("s1");
      Sheet sheet2 = xssfWorkbook.createSheet("s2");
      Random rnd = new Random();
      final int rowCount1 = rnd.nextInt(20);
      final int rowCount2 = rnd.nextInt(20);
      final AtomicInteger total1 = new AtomicInteger();
      final AtomicInteger total2 = new AtomicInteger();
      for (int i = 0; i < rowCount1; i++) {
        int value = rnd.nextInt(1000);
        total1.addAndGet(value);
        sheet1.createRow(i).createCell(0).setCellValue(value);
      }
      for (int i = 0; i < rowCount2; i++) {
        int value = rnd.nextInt(1000);
        total2.addAndGet(value);
        sheet2.createRow(i).createCell(0).setCellValue(value);
      }
      xssfWorkbook.write(bos);

      try (Workbook wb = StreamingReader.builder().open(bos.toInputStream())) {
        assertEquals(2, wb.getNumberOfSheets());
        final Sheet wbSheet1 = wb.getSheet("s1");
        final Sheet wbSheet2 = wb.getSheet("s2");
        final ExecutorService executorService = Executors.newCachedThreadPool();
        final CompletableFuture<Boolean> cf1 = new CompletableFuture<>();
        final CompletableFuture<Boolean> cf2 = new CompletableFuture<>();

        executorService.submit(() -> {
          int rowCount = 0;
          double total = 0.0;
          for (Row row : wbSheet1) {
            rowCount++;
            total += row.getCell(0).getNumericCellValue();
          }
          assertEquals(rowCount1, rowCount);
          assertEquals(total1.get(), (int)total);
          cf1.complete(Boolean.TRUE);
        });

        executorService.submit(() -> {
          int rowCount = 0;
          double total = 0.0;
          for (Row row : wbSheet2) {
            rowCount++;
            total += row.getCell(0).getNumericCellValue();
          }
          assertEquals(rowCount2, rowCount);
          assertEquals(total2.get(), (int)total);
          cf2.complete(Boolean.TRUE);
        });

        assertTrue(cf1.get(30, TimeUnit.SECONDS));
        assertTrue(cf2.get(30, TimeUnit.SECONDS));
      }
    }
  }

  @Test
  public void testSheetReadWrongOrder() throws Exception {
    try (
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
            UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream()
    ) {
      Sheet sheet1 = xssfWorkbook.createSheet("s1");
      Sheet sheet2 = xssfWorkbook.createSheet("s2");
      Random rnd = new Random();
      final int rowCount1 = rnd.nextInt(20);
      final int rowCount2 = rnd.nextInt(20);
      final AtomicInteger total1 = new AtomicInteger();
      final AtomicInteger total2 = new AtomicInteger();
      for (int i = 0; i < rowCount1; i++) {
        int value = rnd.nextInt(1000);
        total1.addAndGet(value);
        sheet1.createRow(i).createCell(0).setCellValue(value);
      }
      for (int i = 0; i < rowCount2; i++) {
        int value = rnd.nextInt(1000);
        total2.addAndGet(value);
        sheet2.createRow(i).createCell(0).setCellValue(value);
      }
      xssfWorkbook.write(bos);

      try (Workbook wb = StreamingReader.builder().open(bos.toInputStream())) {
        assertEquals(2, wb.getNumberOfSheets());
        final Sheet wbSheet2 = wb.getSheet("s2");
        assertEquals("s2", wbSheet2.getSheetName());
        final Sheet wbSheet1 = wb.getSheet("s1");
        assertEquals("s1", wbSheet1.getSheetName());
        Iterator<Sheet> siter = wb.sheetIterator();
        ArrayList<Sheet> sheets = new ArrayList<>();
        siter.forEachRemaining(sheets::add);
        assertEquals(2, sheets.size());
        assertEquals("s1", sheets.get(0).getSheetName());
        assertEquals("s2", sheets.get(1).getSheetName());
        assertEquals(wbSheet1.hashCode(), sheets.get(0).hashCode());
        assertEquals(wbSheet2.hashCode(), sheets.get(1).hashCode());

        int readRowCount1 = 0;
        double readTotal1 = 0.0;
        for (Row row : wbSheet1) {
          readRowCount1++;
          readTotal1 += row.getCell(0).getNumericCellValue();
        }
        assertEquals(rowCount1, readRowCount1);
        assertEquals(total1.get(), (int)readTotal1);
        int readRowCount2 = 0;
        double readTotal2 = 0.0;
        for (Row row : wbSheet2) {
          readRowCount2++;
          readTotal2 += row.getCell(0).getNumericCellValue();
        }
        assertEquals(rowCount2, readRowCount2);
        assertEquals(total2.get(), (int)readTotal2);
      }
    }
  }

  private void validateFormatsSheet(Sheet sheet) throws IOException {
    Iterator<Row> rowIterator = sheet.rowIterator();

    Cell A1 = getCellFromNextRow(rowIterator, 0);
    Cell A2 = getCellFromNextRow(rowIterator, 0);
    Cell A3 = getCellFromNextRow(rowIterator, 0);

    expectFormattedContent(A1, "1234.6");
    expectFormattedContent(A2, "1918-11-11");
    expectFormattedContent(A3, "50%");

    assertTrue(rowIterator instanceof Closeable);
    ((Closeable)rowIterator).close();
  }

  private static class ExploitServer extends NanoHTTPD implements AutoCloseable {
    private final Consumer<IHTTPSession> onRequest;

    public ExploitServer(Consumer<IHTTPSession> onRequest) throws IOException {
      super(61932);
      this.onRequest = onRequest;
    }

    @Override
    public Response serve(IHTTPSession session) {
      onRequest.accept(session);
      return newFixedLengthResponse("<!ENTITY % data SYSTEM \"file://pom.xml\">\n");
    }

    public static void withServer(Consumer<IHTTPSession> onRequest, Runnable func) {
      try(ExploitServer server = new ExploitServer(onRequest)) {
        server.start(NanoHTTPD.SOCKET_READ_TIMEOUT, false);
        func.run();
      } catch(IOException e) {
        throw new UncheckedIOException(e);
      }
    }

    @Override
    public void close() {
      this.stop();
    }
  }
}
