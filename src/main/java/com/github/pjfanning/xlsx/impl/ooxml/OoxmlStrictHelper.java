package com.github.pjfanning.xlsx.impl.ooxml;

import com.github.pjfanning.poi.xssf.streaming.MapBackedCommentsTable;
import com.github.pjfanning.poi.xssf.streaming.MapBackedSharedStringsTable;
import com.github.pjfanning.poi.xssf.streaming.TempFileCommentsTable;
import com.github.pjfanning.poi.xssf.streaming.TempFileSharedStringsTable;
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
          switch (builder.getSharedStringsImplementationType()) {
            case POI_DEFAULT:
              SharedStringsTable sst = new SharedStringsTable();
              try {
                sst.readFrom(is);
              } catch (IOException|RuntimeException e) {
                sst.close();
                throw e;
              }
              return sst;
            case TEMP_FILE_BACKED:
              TempFileSharedStringsTable tfst = new TempFileSharedStringsTable(
                      builder.encryptSstTempFile(), builder.fullFormatRichText());
              try {
                tfst.readFrom(is);
              } catch (IOException|RuntimeException e) {
                tfst.close();
                throw e;
              }
              return tfst;
            case CUSTOM_MAP_BACKED:
              MapBackedSharedStringsTable mbst = new MapBackedSharedStringsTable(builder.fullFormatRichText());
              try {
                mbst.readFrom(is);
              } catch (IOException|RuntimeException e) {
                mbst.close();
                throw e;
              }
              return mbst;
            default:
              return new ReadOnlySharedStringsTable(is);
          }
        }
      }
    }
  }

  public static Comments getCommentsTable(StreamingReader.Builder builder, PackagePart part)
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
        switch (builder.getCommentsImplementationType()) {
          case TEMP_FILE_BACKED:
            TempFileCommentsTable tfct = new TempFileCommentsTable(
                    builder.encryptCommentsTempFile(),
                    builder.fullFormatRichText());
            try {
              tfct.readFrom(is);
            } catch (IOException|RuntimeException e) {
              tfct.close();
              throw e;
            }
            return tfct;
          case CUSTOM_MAP_BACKED:
            MapBackedCommentsTable mbct = new MapBackedCommentsTable(builder.fullFormatRichText());
            try {
              mbct.readFrom(is);
            } catch (IOException|RuntimeException e) {
              mbct.close();
              throw e;
            }
            return mbct;
          default:
            CommentsTable commentsTable = new CommentsTable();
            commentsTable.readFrom(is);
            return commentsTable;
        }
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
