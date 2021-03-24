package com.monitorjbl.xlsx;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Locale;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

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

}
