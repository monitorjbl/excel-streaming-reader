package com.github.pjfanning.xlsx.impl.ooxml;

import com.github.pjfanning.xlsx.StreamingReader;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.*;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.xml.sax.SAXException;

import javax.xml.stream.XMLStreamException;
import java.io.*;
import java.util.List;

public class OoxmlStrictHelper {

  private OoxmlStrictHelper() {}

  public static ThemesTable getThemesTable(StreamingReader.Builder builder, OPCPackage pkg)
          throws IOException, XMLStreamException {
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
        try(InputStream is = tempData.getInputStream()) {
          return new ThemesTable(is);
        }
      }
    }
  }

  public static StylesTable getStylesTable(StreamingReader.Builder builder, OPCPackage pkg)
          throws IOException, XMLStreamException {
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
        try(InputStream is = tempData.getInputStream()) {
          return new StylesTable(is);
        }
      }
    }
  }

  public static SharedStrings getSharedStringsTable(StreamingReader.Builder builder, OPCPackage pkg)
          throws IOException, XMLStreamException, SAXException {
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
        try(InputStream is = tempData.getInputStream()) {
          if (builder.useSstReadOnly()) {
            return new ReadOnlySharedStringsTable(is);
          } else {
            SharedStringsTable sst = new SharedStringsTable();
            try {
              sst.readFrom(is);
            } catch (IOException|RuntimeException e) {
              sst.close();
              throw e;
            }
            return sst;
          }
        }
      }
    }
  }

  public static CommentsTable getCommentsTable(StreamingReader.Builder builder, PackagePart part)
          throws IOException, XMLStreamException {
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
      try(InputStream is = tempData.getInputStream()) {
        CommentsTable commentsTable = new CommentsTable();
        commentsTable.readFrom(is);
        return commentsTable;
      }
    }
  }

  private static TempDataStore createTempDataStore(StreamingReader.Builder builder) {
    if (builder.avoidTempFiles()) {
      return new TempMemoryDataStore();
    } else {
      return new TempFileDataStore();
    }
  }
}
