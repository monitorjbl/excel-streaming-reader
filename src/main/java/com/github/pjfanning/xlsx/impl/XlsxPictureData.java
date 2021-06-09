package com.github.pjfanning.xlsx.impl;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.usermodel.XSSFPictureData;

import java.io.IOException;
import java.io.InputStream;

public class XlsxPictureData extends XSSFPictureData {
  public XlsxPictureData(PackagePart part) {
    super(part);
  }

  public InputStream getInputStream() throws IOException {
    return getPackagePart().getInputStream();
  }
}
