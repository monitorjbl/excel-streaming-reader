package com.monitorjbl.xlsx;

import org.apache.poi.ss.usermodel.*;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Locale;

import static com.monitorjbl.xlsx.TestUtils.expectCachedType;
import static com.monitorjbl.xlsx.TestUtils.expectFormula;
import static com.monitorjbl.xlsx.TestUtils.expectSameStringContent;
import static com.monitorjbl.xlsx.TestUtils.expectStringContent;
import static com.monitorjbl.xlsx.TestUtils.expectType;
import static com.monitorjbl.xlsx.TestUtils.getCellFromNextRow;
import static com.monitorjbl.xlsx.TestUtils.nextRow;
import static com.monitorjbl.xlsx.TestUtils.openWorkbook;
import static org.apache.poi.ss.usermodel.CellType.FORMULA;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
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
    try (Workbook workbook = openWorkbook("formula_cell.xlsx")) {
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
    try (Workbook workbook = openWorkbook("formula_cell.xlsx")) {
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
  public void testStringFormattedFormulaCell() throws Exception {
    try (Workbook workbook = openWorkbook("formula_cell.xlsx")) {
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();

      Cell B1 = getCellFromNextRow(rowIterator, 1);
      nextRow(rowIterator);
      Cell B3 = getCellFromNextRow(rowIterator, 1);

      expectType(B3, FORMULA);
//      expectCachedType(B3, STRING); // this can't return FUNCTION as cached type as per javadoc ! fix in future work
      expectFormula(B3, "B1");
      expectSameStringContent(B1, B3);
      expectStringContent(B3, "a");
    }
  }

  @Test
  public void testQuotedStringFormattedFormulaCell() throws Exception {
    try (Workbook workbook = openWorkbook("formula_cell.xlsx")) {
      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.rowIterator();

      nextRow(rowIterator);
      Cell B2 = getCellFromNextRow(rowIterator, 1);
      nextRow(rowIterator);
      Cell B4 = getCellFromNextRow(rowIterator, 1);

      expectType(B4, FORMULA);
//      expectCachedType(B4, STRING); // this can't return FUNCTION as cached type as per javadoc ! fix in future work
//      expectFormula(B4, "B2"); // returning wrong forumla type? this needs to be fixed in future work
      expectSameStringContent(B2, B4);
      expectStringContent(B4, "\"a\"");
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
