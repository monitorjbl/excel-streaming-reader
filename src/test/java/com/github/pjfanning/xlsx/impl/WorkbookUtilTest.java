package com.github.pjfanning.xlsx.impl;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.junit.Test;

import java.io.File;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

public class WorkbookUtilTest {

  @Test
  public void testUse1904Dates() throws Exception {
    File dates1904File = getFile("1904Dates.xlsx");
    File dates1904TrueFile = getFile("1904Dates_true.xlsx");
    File emptySheetFile = getFile("empty_sheet.xlsx");

    try (OPCPackage pkg = OPCPackage.open(dates1904File, PackageAccess.READ)) {
      assertTrue(WorkbookUtil.use1904Dates(open(pkg)));
    }
    try (OPCPackage pkg = OPCPackage.open(dates1904TrueFile, PackageAccess.READ)) {
      assertTrue(WorkbookUtil.use1904Dates(open(pkg)));
    }
    try (OPCPackage pkg = OPCPackage.open(emptySheetFile, PackageAccess.READ)) {
      assertFalse(WorkbookUtil.use1904Dates(open(pkg)));
    }
  }

  private File getFile(String file) throws Exception {
    return new File("src/test/resources/" + file);
  }

  private XSSFReader open(OPCPackage pkg) throws Exception {
    return new XSSFReader(pkg);
  }

}
