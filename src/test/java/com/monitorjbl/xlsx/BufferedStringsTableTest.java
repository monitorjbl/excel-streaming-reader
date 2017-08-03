package com.monitorjbl.xlsx;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import com.monitorjbl.xlsx.sst.BufferedStringsTable;
import java.io.File;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.junit.Test;

public class BufferedStringsTableTest {

  /**
   * Verifies a bug where BufferedStringsTable was only looking at the first Characters element in a
   * text sequence.
   */
  @Test
  public void testMultiElementCharacters() throws Exception {
    File file = new File("src/test/resources/blank_cells.xlsx");
    OPCPackage pkg = OPCPackage.open(file);

    File sstCache = File.createTempFile("xlsx-strings", "tmp");
    sstCache.deleteOnExit();
    SharedStringsTable sst = BufferedStringsTable.getSharedStringsTable(sstCache, 1000, pkg);
    assertNotNull(sst);
    assertEquals("B1 is Blank --->", sst.getEntryAt(0).getT());
  }
}
