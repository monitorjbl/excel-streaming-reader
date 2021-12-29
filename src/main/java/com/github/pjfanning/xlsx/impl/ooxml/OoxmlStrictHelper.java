package com.github.pjfanning.xlsx.impl.ooxml;

import com.github.pjfanning.xlsx.StreamingReader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.internal.MemoryPackagePart;
import org.apache.poi.openxml4j.opc.internal.TempFilePackagePart;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.*;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.xml.sax.SAXException;

import javax.xml.stream.XMLStreamException;
import java.io.*;
import java.util.List;

public class OoxmlStrictHelper {
  public static ThemesTable getThemesTable(StreamingReader.Builder builder, OPCPackage pkg)
          throws IOException, XMLStreamException, InvalidFormatException {
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
        //remove newPart as part of https://github.com/pjfanning/excel-streaming-reader/issues/88
        PackagePart newPart = createTempPackagePart(builder, pkg, part);
        try(InputStream is = tempData.getInputStream()) {
          newPart.load(is);
          return new ThemesTable(newPart);
        } finally {
          newPart.close();
        }
      }
    }
  }

  public static StylesTable getStylesTable(StreamingReader.Builder builder, OPCPackage pkg)
          throws IOException, XMLStreamException, InvalidFormatException {
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
        PackagePart newPart = createTempPackagePart(builder, pkg, part);
        try(InputStream is = tempData.getInputStream()) {
          newPart.load(is);
          return new StylesTable(newPart);
        } finally {
          newPart.close();
        }
      }
    }
  }

  public static SharedStrings getSharedStringsTable(StreamingReader.Builder builder, OPCPackage pkg)
          throws IOException, XMLStreamException, SAXException, InvalidFormatException {
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
            PackagePart newPart = createTempPackagePart(builder, pkg, part);
            try {
              newPart.load(is);
              return new ReadOnlySharedStringsTable(newPart);
            } finally {
              newPart.close();
            }
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
          throws IOException, XMLStreamException, InvalidFormatException {
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

  private static PackagePart createTempPackagePart(StreamingReader.Builder builder, OPCPackage pkg,
                                                   PackagePart part) throws IOException, InvalidFormatException {
    if (builder.avoidTempFiles()) {
      return new MemoryPackagePart(pkg, part.getPartName(), part.getContentType());
    } else {
      return new TempFilePackagePart(pkg, part.getPartName(), part.getContentType());
    }
  }
}
