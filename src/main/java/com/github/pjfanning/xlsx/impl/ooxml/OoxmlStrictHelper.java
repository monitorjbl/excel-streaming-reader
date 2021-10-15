package com.github.pjfanning.xlsx.impl.ooxml;

import com.github.pjfanning.xlsx.StreamingReader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.openxml4j.opc.internal.MemoryPackagePart;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFRelation;

import javax.xml.stream.XMLStreamException;
import java.io.*;
import java.util.List;

public class OoxmlStrictHelper {
  public static ThemesTable getThemesTable(StreamingReader.Builder builder, OPCPackage pkg) throws IOException, XMLStreamException, InvalidFormatException {
    List<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.THEME.getContentType());
    if (parts.isEmpty()) {
      return null;
    } else {
      PackagePart part = parts.get(0);
      try(TempDataStore tempData = createTempDataStore(builder)) {
        try(
                InputStream is = part.getInputStream();
                OutputStream os = tempData.getOutputStream();
                OoXmlStrictConverter converter = new OoXmlStrictConverter(is, os)
        ) {
          while (converter.convertNextElement()) {
            //continue
          }
        }
        MemoryPackagePart newPart = new MemoryPackagePart(pkg, part.getPartName(), part.getContentType());
        try(InputStream is = tempData.getInputStream()) {
          newPart.load(is);
        }
        return new ThemesTable(newPart);
      }
    }
  }

  public static StylesTable getStylesTable(StreamingReader.Builder builder, OPCPackage pkg) throws IOException, XMLStreamException, InvalidFormatException {
    List<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.STYLES.getContentType());
    if (parts.isEmpty()) {
      return null;
    } else {
      PackagePart part = parts.get(0);
      try(TempDataStore tempData = createTempDataStore(builder)) {
        try(
                InputStream is = part.getInputStream();
                OutputStream os = tempData.getOutputStream();
                OoXmlStrictConverter converter = new OoXmlStrictConverter(is, os)
        ) {
          while (converter.convertNextElement()) {
            //continue
          }
        }
        //TODO when POI 5.1.0 is ready, support using TempFilePackagePart
        MemoryPackagePart newPart = new MemoryPackagePart(pkg, part.getPartName(), part.getContentType());
        try(InputStream is = tempData.getInputStream()) {
          newPart.load(is);
        }
        return new StylesTable(newPart);
      }
    }
  }

  public static SharedStringsTable getSharedStringsTable(StreamingReader.Builder builder, OPCPackage pkg) throws IOException, XMLStreamException, InvalidFormatException {
    List<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.SHARED_STRINGS.getContentType());
    if (parts.isEmpty()) {
      return null;
    } else {
      PackagePart part = parts.get(0);
      try(TempDataStore tempData = createTempDataStore(builder)) {
        try(
                InputStream is = part.getInputStream();
                OutputStream os = tempData.getOutputStream();
                OoXmlStrictConverter converter = new OoXmlStrictConverter(is, os)
        ) {
          while (converter.convertNextElement()) {
            //continue
          }
        }
        //TODO when POI 5.1.0 is ready, support using TempFilePackagePart
        MemoryPackagePart newPart = new MemoryPackagePart(pkg, part.getPartName(), part.getContentType());
        try(InputStream is = tempData.getInputStream()) {
          newPart.load(is);
        }
        return new SharedStringsTable(newPart);
      }
    }
  }

  //TODO OPCPackage has this method in POI 5.0.1
  public static boolean isStrictOoxmlFormat(OPCPackage pkg) {
    PackageRelationshipCollection coreDocRelationships = pkg.getRelationshipsByType(
            PackageRelationshipTypes.STRICT_CORE_DOCUMENT);
    return coreDocRelationships.size() > 0;
  }

  private static TempDataStore createTempDataStore(StreamingReader.Builder builder) {
    if (builder.avoidTempFiles()) {
      return new TempMemoryDataStore();
    } else {
      return new TempFileDataStore();
    }
  }
}
