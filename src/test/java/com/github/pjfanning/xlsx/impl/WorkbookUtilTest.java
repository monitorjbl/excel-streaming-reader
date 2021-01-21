package com.github.pjfanning.xlsx.impl;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.junit.Test;

import java.io.File;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

public class WorkbookUtilTest {

  @Test
  public void testUse1904Dates() throws Exception {
    assertTrue(WorkbookUtil.use1904Dates(open("1904Dates.xlsx")));
    assertTrue(WorkbookUtil.use1904Dates(open("1904Dates_true.xlsx")));
    assertFalse(WorkbookUtil.use1904Dates(open("empty_sheet.xlsx")));
  }

  private XSSFReader open(String file) throws Exception {
    return new XSSFReader(OPCPackage.open(new File("src/test/resources/" + file)));
  }

}
