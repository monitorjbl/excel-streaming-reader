package com.github.pjfanning.xlsx.impl.ooxml;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.internal.MemoryPackagePart;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFRelation;

import javax.xml.stream.XMLStreamException;
import java.io.*;
import java.util.List;

public class OoxmlStrictHelper {
  public static ThemesTable getThemesTable(OPCPackage pkg) throws IOException, XMLStreamException, InvalidFormatException {
    List<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.THEME.getContentType());
    if (parts.isEmpty()) {
      return null;
    } else {
      PackagePart part = parts.get(0);
      File tempFile = TempFile.createTempFile("ooxml-strict-themes", ".xml");
      try {
        try(
                InputStream is = part.getInputStream();
                OutputStream os = new FileOutputStream(tempFile);
                OoXmlStrictConverter converter = new OoXmlStrictConverter(is, os)
        ) {
          while (converter.convertNextElement()) {
            //continue
          }
        }
        MemoryPackagePart newPart = new MemoryPackagePart(pkg, part.getPartName(), part.getContentType());
        try(InputStream is = new FileInputStream(tempFile)) {
          newPart.load(is);
        }
        return new ThemesTable(newPart);
      } finally {
        tempFile.delete();
      }
    }
  }

  public static StylesTable getStylesTable(OPCPackage pkg) throws IOException, XMLStreamException, InvalidFormatException {
    List<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.STYLES.getContentType());
    if (parts.isEmpty()) {
      return null;
    } else {
      PackagePart part = parts.get(0);
      File tempFile = TempFile.createTempFile("ooxml-strict-styles", ".xml");
      try {
        try(
                InputStream is = part.getInputStream();
                OutputStream os = new FileOutputStream(tempFile);
                OoXmlStrictConverter converter = new OoXmlStrictConverter(is, os)
        ) {
          while (converter.convertNextElement()) {
            //continue
          }
        }
        MemoryPackagePart newPart = new MemoryPackagePart(pkg, part.getPartName(), part.getContentType());
        try(InputStream is = new FileInputStream(tempFile)) {
          newPart.load(is);
        }
        return new StylesTable(newPart);
      } finally {
        tempFile.delete();
      }
    }
  }
}
