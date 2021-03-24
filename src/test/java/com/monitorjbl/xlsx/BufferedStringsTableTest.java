package com.monitorjbl.xlsx;

import com.monitorjbl.xlsx.sst.BufferedStringsTable;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.junit.jupiter.api.Test;

import java.io.File;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;


public class BufferedStringsTableTest {

  /**
   * Verifies a bug where BufferedStringsTable was only looking at the first Characters xml element
   * in a text sequence.
   */
  @Test
  public void testStringsWithMultipleXmlElements() throws Exception {
    File file = new File("src/test/resources/blank_cells.xlsx");
    File sstCache = File.createTempFile("cache", ".sst");
    sstCache.deleteOnExit();
    try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ);
         BufferedStringsTable sst = BufferedStringsTable.getSharedStringsTable(sstCache, 1000, pkg)) {
      assertNotNull(sst);
      assertEquals("B1 is Blank --->", sst.getItemAt(0).getString());
    }
  }

  /**
   * Verifies a bug where BufferedStringsTable was dropping text enclosed in formatting
   * instructions.
   */
  @Test
  public void testStringsWrappedInFormatting() throws Exception {
    File file = new File("src/test/resources/shared_styled_string.xlsx");
    File sstCache = File.createTempFile("cache", ".sst");
    sstCache.deleteOnExit();
    try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ);
         BufferedStringsTable sst = BufferedStringsTable.getSharedStringsTable(sstCache, 1000, pkg)) {
      assertNotNull(sst);
      assertEquals("shared styled string", sst.getItemAt(0).getString());
    }
  }
}
