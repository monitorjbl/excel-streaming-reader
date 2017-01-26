package com.monitorjbl.xlsx;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.monitorjbl.xlsx.sst.CTRstImpl;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Ignore;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.security.SecureRandom;

public class PerformanceTest {

  @Ignore
  @Test
  public void testPerformance() throws Exception {
    long start = System.currentTimeMillis();
    try(
        InputStream is = new FileInputStream(new File("/tmp/test.xlsx"));
        Workbook wb = StreamingReader.builder()
            .sstCacheSize(100)
            .open(is);
    ) {
      Sheet sheet = wb.getSheetAt(0);
      long count = 0;
      for(Row row : sheet) {
        count++;
      }
      System.out.println("Read "+count+" rows in "+(System.currentTimeMillis() - start)+"ms");
    }
  }

  @Ignore
  @Test
  public void testDataCreate() {
    Workbook wb = new XSSFWorkbook();
    CreationHelper createHelper = wb.getCreationHelper();
    Sheet sheet = wb.createSheet("new sheet");
    SecureRandom random = new SecureRandom();

    for(int i = 0; i < 100_000; i++) {
      Row row = sheet.createRow(i);
      for(int j = 0; j < 10; j++) {
        Cell cell = row.createCell(j);
        cell.setCellValue(new BigInteger(130, random).toString(32));
      }
    }

    try(OutputStream fos = new FileOutputStream("/tmp/test.xlsx")) {
      wb.write(fos);
    } catch(IOException e) {
      e.printStackTrace();
    }

  }
}
