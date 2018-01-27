package com.monitorjbl.xlsx;

import org.apache.poi.ss.usermodel.*;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Locale;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

public class StreamingWorkbookTest {
  @BeforeClass
  public static void init() {
    Locale.setDefault(Locale.ENGLISH);
  }

  @Test
  public void testIterateSheets() throws Exception {
    try(
        InputStream is = new FileInputStream(new File("src/test/resources/sheets.xlsx"));
        Workbook workbook = StreamingReader.builder().open(is);
    ) {

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
  }

  @Test
  public void testHiddenCells() throws Exception {
    try(
        InputStream is = new FileInputStream(new File("src/test/resources/hidden.xlsx"));
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
        InputStream is = new FileInputStream(new File("src/test/resources/hidden.xlsx"));
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
    try(
        InputStream is = new FileInputStream(new File("src/test/resources/formula_cell.xlsx"));
        Workbook workbook = StreamingReader.builder().open(is)
    ) {
      assertEquals(1, workbook.getNumberOfSheets());
      Sheet sheet = workbook.getSheetAt(0);

      Iterator<Row> rowIterator = sheet.rowIterator();
      rowIterator.next();
      rowIterator.next();
      Row row3 = rowIterator.next();
      Cell A3 = row3.getCell(0);

      assertEquals("Cell A3 should be of type formula", CellType.FORMULA, A3.getCellTypeEnum());
      assertEquals("Cell A3's value should be of type numeric", CellType.NUMERIC, A3.getCachedFormulaResultTypeEnum());
      assertEquals("Wrong formula", "SUM(A1:A2)", A3.getCellFormula());
    }
  }

  @Test
  public void testCellComments() throws Exception {
    try(
            InputStream is = new FileInputStream(new File("src/test/resources/read_cell_comments.xlsx"));
            Workbook workbook = StreamingReader.builder().readComments().open(is)
    ) {
      assertEquals(3, workbook.getNumberOfSheets());
      Sheet sheet1 = workbook.getSheetAt(0);

      Iterator<Row> rowIterator = sheet1.rowIterator();
      Row row1 = rowIterator.next();
      Cell A1 = row1.getCell(0);

      assertEquals("Cell A1 should have data", "A1", A1.getStringCellValue());
      assertNotNull("Cell A1 should have a comment", A1.getCellComment());
      String A1Author = A1.getCellComment().getAuthor();
      assertEquals("Invalid comment author", "BBonev", A1Author);
      // the author is visible in the comment
      assertEquals("Invalid comment text", A1Author + ":\nA1 comment\nhere on the second line", A1.getCellComment().getString().getString());

      Sheet sheet3 = workbook.getSheetAt(2);

      rowIterator = sheet3.rowIterator();
      rowIterator.next();
      Row row2 = rowIterator.next();
      Cell B2 = row2.getCell(1);

      assertEquals("Cell B2 should have data", "B2S3", B2.getStringCellValue());
      assertNotNull("Cell B2 should have a comment", B2.getCellComment());
      assertEquals("Invalid comment author", "BBonev", B2.getCellComment().getAuthor());
      // the author is not visible in the comment
      assertEquals("Invalid comment text", "Comment from B2 sheet 3", B2.getCellComment().getString().getString());
    }
  }
}
