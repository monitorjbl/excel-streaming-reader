package com.monitorjbl.xlsx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class PerformanceTest {

  public static void main(String[] args) throws IOException {
    for(int i = 0; i < 10; i++) {
      long start = System.currentTimeMillis();
      InputStream is = new FileInputStream(new File("/Users/thundermoose/Downloads/SampleXLSFile_6800kb.xlsx"));
      try(Workbook workbook = StreamingReader.builder()
          .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
          .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
          .open(is)) {

        for(Row r : workbook.getSheet("test")) {
          for(Cell c : r) {
            //do nothing
          }
        }
      }

      long end = System.currentTimeMillis();
      System.out.println("Time: " + (end - start) + "ms");
    }
  }
}
