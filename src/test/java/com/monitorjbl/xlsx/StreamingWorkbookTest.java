package com.monitorjbl.xlsx;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Locale;

import static org.junit.Assert.assertEquals;

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

}
