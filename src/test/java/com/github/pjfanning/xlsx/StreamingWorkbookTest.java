package com.github.pjfanning.xlsx;

import com.github.pjfanning.xlsx.exceptions.ParseException;
import com.github.pjfanning.xlsx.impl.XlsxPictureData;
import fi.iki.elonen.NanoHTTPD;
import org.apache.commons.io.IOUtils;
import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.*;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.function.Consumer;

import static com.github.pjfanning.xlsx.TestUtils.*;
import static org.apache.poi.ss.usermodel.CellType.*;
import static org.junit.Assert.*;

public class StreamingWorkbookTest {
  @BeforeClass
  public static void init() {
    Locale.setDefault(Locale.ENGLISH);
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

      assertFalse("Row 0 should not be hidden", sheet.rowIterator().next().getZeroHeight());
      assertTrue("Row 1 should be hidden", sheet.rowIterator().next().getZeroHeight());
      assertFalse("Row 2 should not be hidden", sheet.rowIterator().next().getZeroHeight());
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
//      expectCachedType(B4, STRING); // this can't return FUNCTION as cached type as per javadoc ! fix in future work
//      expectFormula(B4, "B2"); // returning wrong formula type? this needs to be fixed in future work
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

  @Test(expected = ParseException.class)
  public void testEntityExpansion() throws Exception {
    ExploitServer.withServer(s -> fail("Should not have made request"), () -> {
      try(Workbook workbook = openWorkbook("entity-expansion-exploit-poc-file.xlsx")) {
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
      Iterator<Row> rowIterator = sheet.rowIterator();

      Cell A1 = getCellFromNextRow(rowIterator, 0);
      Cell A2 = getCellFromNextRow(rowIterator, 0);
      Cell A3 = getCellFromNextRow(rowIterator, 0);

      expectFormattedContent(A1, "1234.6");
      expectFormattedContent(A2, "1918-11-11");
      expectFormattedContent(A3, "50%");
    }
  }

  @Test
  public void testRightToLeft() throws IOException {
    try(
            InputStream stream = getInputStream("right-to-left.xlsx");
            Workbook workbook = StreamingReader.builder().setReadComments(true).open(stream)
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
      assertTrue(a3Comment.getString().getString().contains("تعليق الاختبا"));
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