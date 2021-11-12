package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.poi.xssf.streaming.TempFileSharedStringsTable;
import com.github.pjfanning.xlsx.StreamingReader.Builder;
import com.github.pjfanning.xlsx.XmlUtils;
import com.github.pjfanning.xlsx.exceptions.NotSupportedException;
import com.github.pjfanning.xlsx.exceptions.OpenException;
import com.github.pjfanning.xlsx.exceptions.ParseException;
import com.github.pjfanning.xlsx.exceptions.ReadException;
import com.github.pjfanning.xlsx.impl.ooxml.OoxmlStrictHelper;
import com.github.pjfanning.xlsx.impl.ooxml.OoxmlReader;
import com.github.pjfanning.xlsx.impl.ooxml.ResourceWithTrackedCloseable;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Date1904Support;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.Internal;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.model.*;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.util.*;

import static com.github.pjfanning.xlsx.XmlUtils.searchForNodeList;

public class StreamingWorkbookReader implements Iterable<Sheet>, Date1904Support, AutoCloseable {
  private static final Logger log = LoggerFactory.getLogger(StreamingWorkbookReader.class);
  private static XMLInputFactory xmlInputFactory;

  private final List<StreamingSheet> sheets;
  private final List<Map<String, String>> sheetProperties = new ArrayList<>();
  private final Map<String, List<XSSFShape>> shapeMap = new HashMap<>();
  private final Builder builder;
  private File tmp;
  private OPCPackage pkg;
  private SharedStringsTable sst;
  private boolean use1904Dates = false;
  private StreamingWorkbook workbook = null;
  private POIXMLProperties.CoreProperties coreProperties = null;
  private final List<ResourceWithTrackedCloseable<?>> trackedCloseables = new ArrayList<>();

  public StreamingWorkbookReader(Builder builder) {
    this.sheets = new ArrayList<>();
    this.builder = builder;
  }

  public void init(InputStream is) {
    if (builder.avoidTempFiles()) {
      try {
        if(builder.getPassword() != null) {
          POIFSFileSystem poifs = new POIFSFileSystem(is);
          pkg = decryptWorkbook(poifs);
        } else {
          pkg = OPCPackage.open(is);
        }
        loadPackage(pkg);
      } catch(SAXException | ParserConfigurationException e) {
        IOUtils.closeQuietly(pkg);
        throw new ParseException("Failed to parse stream", e);
      } catch(IOException e) {
        IOUtils.closeQuietly(pkg);
        throw new OpenException("Failed to open stream", e);
      } catch(OpenXML4JException | XMLStreamException e) {
        IOUtils.closeQuietly(pkg);
        throw new ReadException("Unable to read workbook", e);
      } catch(GeneralSecurityException e) {
        IOUtils.closeQuietly(pkg);
        throw new ReadException("Unable to read workbook - Decryption failed", e);
      }
    } else {
      File f = null;
      try {
        f = TempFileUtil.writeInputStreamToFile(is, builder.getBufferSize());
        log.debug("Created temp file [" + f.getAbsolutePath() + "]");
        init(f);
        tmp = f;
      } catch(IOException e) {
        if (f != null) {
          f.delete();
        }
        throw new ReadException("Unable to read input stream", e);
      } catch(RuntimeException e) {
        if (f != null) {
          f.delete();
        }
        throw e;
      }
    }
  }

  public void init(File f) {
    try {
      if(builder.getPassword() != null) {
        POIFSFileSystem poifs = new POIFSFileSystem(f);
        pkg = decryptWorkbook(poifs);
      } else {
        pkg = OPCPackage.open(f);
      }
      loadPackage(pkg);
    } catch(SAXException | ParserConfigurationException e) {
      IOUtils.closeQuietly(pkg);
      throw new ParseException("Failed to parse file", e);
    } catch(IOException e) {
      IOUtils.closeQuietly(pkg);
      throw new OpenException("Failed to open file", e);
    } catch(OpenXML4JException | XMLStreamException e) {
      IOUtils.closeQuietly(pkg);
      throw new ReadException("Unable to read workbook", e);
    } catch(GeneralSecurityException e) {
      IOUtils.closeQuietly(pkg);
      throw new ReadException("Unable to read workbook - Decryption failed", e);
    }
  }

  private OPCPackage decryptWorkbook(POIFSFileSystem poifs) throws IOException, GeneralSecurityException, InvalidFormatException {
    // Based on: https://poi.apache.org/encryption.html
    EncryptionInfo info = new EncryptionInfo(poifs);
    Decryptor d = Decryptor.getInstance(info);
    d.verifyPassword(builder.getPassword());
    return OPCPackage.open(d.getDataStream(poifs));
  }

  private void loadPackage(OPCPackage pkg) throws IOException, OpenXML4JException, ParserConfigurationException, SAXException, XMLStreamException {
    boolean strictFormat = pkg.isStrictOoxmlFormat();
    OoxmlReader reader = new OoxmlReader(this, pkg, strictFormat);
    if (strictFormat) {
      log.info("file is in strict OOXML format");
    }
    if(builder.useSstTempFile()) {
      log.debug("Created sst cache file");
      sst = new TempFileSharedStringsTable(pkg, builder.encryptSstTempFile(), builder.fullFormatRichText());
    } else if(strictFormat) {
      sst = OoxmlStrictHelper.getSharedStringsTable(builder, pkg);
    } else {
      sst = reader.getSharedStringsTable();
    }

    if (builder.readCoreProperties()) {
      try {
        POIXMLProperties xmlProperties = new POIXMLProperties(pkg);
        coreProperties = xmlProperties.getCoreProperties();
      } catch (Exception e) {
        log.warn("Failed to read coreProperties", e);
      }
    }

    StylesTable styles;
    if(strictFormat) {
      ResourceWithTrackedCloseable<ThemesTable> themesTable = OoxmlStrictHelper.getThemesTable(builder, pkg);
      ResourceWithTrackedCloseable<StylesTable> stylesTable = OoxmlStrictHelper.getStylesTable(builder, pkg);
      styles = stylesTable.getResource();
      styles.setTheme(themesTable.getResource());
    } else {
      styles = reader.getStylesTable();
    }

    use1904Dates = WorkbookUtil.use1904Dates(reader);

    loadSheets(reader, sst, styles, builder.getRowCacheSize());
  }

  void setWorkbook(StreamingWorkbook workbook) {
    this.workbook = workbook;
    workbook.setCoreProperties(coreProperties);
  }

  Workbook getWorkbook() {
    return workbook;
  }

  void loadSheets(OoxmlReader reader, SharedStringsTable sst, StylesTable stylesTable, int rowCacheSize) throws IOException, InvalidFormatException,
      XMLStreamException {
    lookupSheetNames(reader);

    //Some workbooks have multiple references to the same sheet. Need to filter
    //them out before creating the XMLEventReader by keeping track of their URIs.
    //The sheets are listed in order, so we must keep track of insertion order.
    OoxmlReader.OoxmlSheetIterator iter = reader.getSheetsData();
    Map<PackagePart, InputStream> sheetStreams = new LinkedHashMap<>();
    Map<PackagePart, Comments> sheetComments = new HashMap<>();
    while(iter.hasNext()) {
      InputStream is = iter.next();
      if (builder.readShapes()) {
        shapeMap.put(iter.getSheetName(), iter.getShapes());
      }
      PackagePart part = iter.getSheetPart();
      sheetStreams.put(part, is);
      if (builder.readComments()) {
        sheetComments.put(part, iter.getSheetComments(builder));
      }
    }

    //Iterate over the loaded streams
    int i = 0;
    for(PackagePart packagePart : sheetStreams.keySet()) {
      XMLEventReader parser = getXmlInputFactory().createXMLEventReader(sheetStreams.get(packagePart));
      sheets.add(new StreamingSheet(
              workbook,
              sheetProperties.get(i++).get("name"),
              new StreamingSheetReader(this, packagePart, sst, stylesTable,
                      sheetComments.get(packagePart), parser, use1904Dates, rowCacheSize)));
    }
  }

  void lookupSheetNames(OoxmlReader reader) throws IOException, InvalidFormatException {
    sheetProperties.clear();
    try {
      NodeList nl = searchForNodeList(XmlUtils.readDocument(reader.getWorkbookData()), "/ss:workbook/ss:sheets/ss:sheet");
      for(int i = 0; i < nl.getLength(); i++) {
        Map<String, String> props = new HashMap<>();
        props.put("name", nl.item(i).getAttributes().getNamedItem("name").getTextContent());

        Node state = nl.item(i).getAttributes().getNamedItem("state");
        props.put("state", state == null ? "visible" : state.getTextContent());
        sheetProperties.add(props);
      }
    } catch (SAXException|ParserConfigurationException e) {
      throw new ParseException("Failed to parse file", e);
    }
  }

  List<? extends Sheet> getSheets() {
    return sheets;
  }

  public List<Map<String, String>> getSheetProperties() {
    return sheetProperties;
  }

  @Override
  public Iterator<Sheet> iterator() {
    return new StreamingSheetIterator(sheets.iterator());
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public boolean isDate1904() {
    return use1904Dates;
  }

  @Override
  public void close() throws IOException {
    try {
      for(StreamingSheet sheet : sheets) {
        sheet.getReader().close();
      }
      pkg.revert();
    } finally {
      if(tmp != null) {
        if (log.isDebugEnabled()) {
          log.debug("Deleting tmp file [" + tmp.getAbsolutePath() + "]");
        }
        tmp.delete();
      }
      if(sst != null) {
        sst.close();
      }
      for(ResourceWithTrackedCloseable<?> trackedCloseable : trackedCloseables) {
        trackedCloseable.close();
      }
    }
  }

  Builder getBuilder() {
    return builder;
  }

  OPCPackage getOPCPackage() {
    return pkg;
  }

  List<XSSFShape> getShapes(String sheetName) {
    return shapeMap.get(sheetName);
  }

  private static XMLInputFactory getXmlInputFactory() {
    if (xmlInputFactory == null) {
      try {
        xmlInputFactory = XMLHelper.newXMLInputFactory();
      } catch (Throwable t) {
        log.error("Issue creating XMLInputFactory", t);
        throw t;
      }
    }
    return xmlInputFactory;
  }

  /**
   * Internal use only. To track resources that should be closed when this reader instance is closed.
   * @param trackedCloseable resource to close (later)
   */
  @Internal
  public void addTrackableCloseable(ResourceWithTrackedCloseable<?> trackedCloseable) {
    this.trackedCloseables.add(trackedCloseable);
  }

  static class StreamingSheetIterator implements Iterator<Sheet> {
    private final Iterator<StreamingSheet> iterator;

    public StreamingSheetIterator(Iterator<StreamingSheet> iterator) {
      this.iterator = iterator;
    }

    @Override
    public boolean hasNext() {
      return iterator.hasNext();
    }

    @Override
    public Sheet next() {
      return iterator.next();
    }

    @Override
    public void remove() {
      throw new NotSupportedException();
    }
  }
}
