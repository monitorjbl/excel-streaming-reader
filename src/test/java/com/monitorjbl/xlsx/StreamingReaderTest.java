package com.monitorjbl.xlsx;

import com.monitorjbl.xlsx.exceptions.MissingSheetException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Ignore;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

public class StreamingReaderTest {

  @Ignore
  public void testFullSet() throws Exception {
    Workbook wb = new XSSFWorkbook(new FileInputStream("src/test/resources/gaps.xlsx"));
    Sheet s = wb.getSheet("Sheet1");

    for (Row r : s) {
      for (Cell c : r) {
        System.out.println(c.getRowIndex() + ":" + c.getColumnIndex());
      }
    }
  }

  @Test
  public void testTypes() throws Exception {
    SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy");
    try (
        InputStream is = new FileInputStream(new File("src/test/resources/data_types.xlsx"));
        StreamingReader reader = StreamingReader.builder().read(is);
    ) {

      List<List<Cell>> obj = new ArrayList<>();

      for (Row r : reader) {
        List<Cell> o = new ArrayList<>();
        for (Cell c : r) {
          o.add(c);
        }
        obj.add(o);
      }

      assertEquals(6, obj.size());
      List<Cell> row;

      row = obj.get(0);
      assertEquals(2, row.size());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(0).getCellType());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(1).getCellType());
      assertEquals("Type", row.get(0).getStringCellValue());
      assertEquals("Value", row.get(1).getStringCellValue());

      row = obj.get(1);
      assertEquals(2, row.size());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(0).getCellType());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(1).getCellType());
      assertEquals("string", row.get(0).getStringCellValue());
      assertEquals("jib-jab", row.get(1).getStringCellValue());

      row = obj.get(2);
      assertEquals(2, row.size());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(0).getCellType());
      assertEquals(Cell.CELL_TYPE_NUMERIC, row.get(1).getCellType());
      assertEquals("int", row.get(0).getStringCellValue());
      assertEquals(10, row.get(1).getNumericCellValue(), 0);

      row = obj.get(3);
      assertEquals(2, row.size());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(0).getCellType());
      assertEquals(Cell.CELL_TYPE_NUMERIC, row.get(1).getCellType());
      assertEquals("double", row.get(0).getStringCellValue());
      assertEquals(3.14, row.get(1).getNumericCellValue(), 0);

      row = obj.get(4);
      assertEquals(2, row.size());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(0).getCellType());
      assertEquals(Cell.CELL_TYPE_NUMERIC, row.get(1).getCellType());
      assertEquals("date", row.get(0).getStringCellValue());
      assertEquals(df.parse("1/1/2014"), row.get(1).getDateCellValue());

      row = obj.get(5);
      assertEquals(7, row.size());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(0).getCellType());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(1).getCellType());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(2).getCellType());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(3).getCellType());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(4).getCellType());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(5).getCellType());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(6).getCellType());
      assertEquals("long", row.get(0).getStringCellValue());
      assertEquals("ass", row.get(1).getStringCellValue());
      assertEquals("row", row.get(2).getStringCellValue());
      assertEquals("look", row.get(3).getStringCellValue());
      assertEquals("at", row.get(4).getStringCellValue());
      assertEquals("it", row.get(5).getStringCellValue());
      assertEquals("go", row.get(6).getStringCellValue());
    }
  }

  @Test
  public void testGaps() throws Exception {
    try (
        InputStream is = new FileInputStream(new File("src/test/resources/gaps.xlsx"));
        StreamingReader reader = StreamingReader.builder().read(is);
    ) {
      List<List<Cell>> obj = new ArrayList<>();

      for (Row r : reader) {
        List<Cell> o = new ArrayList<>();
        for (Cell c : r) {
          o.add(c);
        }
        obj.add(o);
      }

      assertEquals(2, obj.size());
      List<Cell> row;

      row = obj.get(0);
      assertEquals(2, row.size());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(0).getCellType());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(1).getCellType());
      assertEquals("Dat", row.get(0).getStringCellValue());
      assertEquals(0, row.get(0).getColumnIndex());
      assertEquals(0, row.get(0).getRowIndex());
      assertEquals("gap", row.get(1).getStringCellValue());
      assertEquals(2, row.get(1).getColumnIndex());
      assertEquals(0, row.get(1).getRowIndex());

      row = obj.get(1);
      assertEquals(2, row.size());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(0).getCellType());
      assertEquals(Cell.CELL_TYPE_STRING, row.get(1).getCellType());
      assertEquals("guuurrrrrl", row.get(0).getStringCellValue());
      assertEquals(0, row.get(0).getColumnIndex());
      assertEquals(6, row.get(0).getRowIndex());
      assertEquals("!", row.get(1).getStringCellValue());
      assertEquals(6, row.get(1).getColumnIndex());
      assertEquals(6, row.get(1).getRowIndex());
    }
  }

  @Test
  public void testMultipleSheets_alpha() throws Exception {
    try (
        InputStream is = new FileInputStream(new File("src/test/resources/sheets.xlsx"));
        StreamingReader reader = StreamingReader.builder()
            .sheetIndex(0)
            .read(is);
    ) {
      List<List<Cell>> obj = new ArrayList<>();

      for (Row r : reader) {
        List<Cell> o = new ArrayList<>();
        for (Cell c : r) {
          o.add(c);
        }
        obj.add(o);
      }

      assertEquals(1, obj.size());
      List<Cell> row;

      row = obj.get(0);
      assertEquals(1, row.size());
      assertEquals("stuff", row.get(0).getStringCellValue());
    }
  }

  @Test
  public void testMultipleSheets_zulu() throws Exception {
    try (
        InputStream is = new FileInputStream(new File("src/test/resources/sheets.xlsx"));
        StreamingReader reader = StreamingReader.builder()
            .sheetIndex(1)
            .read(is);
    ) {

      List<List<Cell>> obj = new ArrayList<>();

      for (Row r : reader) {
        List<Cell> o = new ArrayList<>();
        for (Cell c : r) {
          o.add(c);
        }
        obj.add(o);
      }

      assertEquals(1, obj.size());
      List<Cell> row;

      row = obj.get(0);
      assertEquals(1, row.size());
      assertEquals("yeah", row.get(0).getStringCellValue());
    }
  }

  @Test
  public void testSheetName_zulu() throws Exception {
    try (
        InputStream is = new FileInputStream(new File("src/test/resources/sheets.xlsx"));
        StreamingReader reader = StreamingReader.builder()
            .sheetName("SheetZulu")
            .read(is);
    ) {

      List<List<Cell>> obj = new ArrayList<>();

      for (Row r : reader) {
        List<Cell> o = new ArrayList<>();
        for (Cell c : r) {
          o.add(c);
        }
        obj.add(o);
      }

      assertEquals(1, obj.size());
      List<Cell> row;

      row = obj.get(0);
      assertEquals(1, row.size());
      assertEquals("yeah", row.get(0).getStringCellValue());
    }
  }

  @Test
  public void testSheetName_alpha() throws Exception {
    try (
        InputStream is = new FileInputStream(new File("src/test/resources/sheets.xlsx"));
        StreamingReader reader = StreamingReader.builder()
            .sheetName("SheetAlpha")
            .read(is);
    ) {
      List<List<Cell>> obj = new ArrayList<>();

      for (Row r : reader) {
        List<Cell> o = new ArrayList<>();
        for (Cell c : r) {
          o.add(c);
        }
        obj.add(o);
      }

      assertEquals(1, obj.size());
      List<Cell> row;

      row = obj.get(0);
      assertEquals(1, row.size());
      assertEquals("stuff", row.get(0).getStringCellValue());
    }
  }

  @Test(expected = MissingSheetException.class)
  public void testSheetName_missingInStream() throws Exception {
    try (
        InputStream is = new FileInputStream(new File("src/test/resources/sheets.xlsx"));
        StreamingReader reader = StreamingReader.builder()
            .sheetName("adsfasdfasdfasdf")
            .read(is);
    ) {
      fail("Should have failed");
    }
  }

  @Test
  public void testSheetName_missingInFile() throws Exception {
    File f = new File("src/test/resources/sheets.xlsx");
    try (
        StreamingReader reader = StreamingReader.builder()
            .sheetName("adsfasdfasdfasdf")
            .read(f);
    ) {
      fail("Should have failed");
    } catch (MissingSheetException e) {
      assertTrue(f.exists());
    }
  }
}
