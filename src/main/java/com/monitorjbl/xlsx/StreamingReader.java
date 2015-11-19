package com.monitorjbl.xlsx;

import com.monitorjbl.xlsx.exceptions.CloseException;
import com.monitorjbl.xlsx.exceptions.MissingSheetException;
import com.monitorjbl.xlsx.exceptions.OpenException;
import com.monitorjbl.xlsx.exceptions.ReadException;
import com.monitorjbl.xlsx.impl.StreamingCell;
import com.monitorjbl.xlsx.impl.StreamingRow;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.Characters;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

import static com.monitorjbl.xlsx.XmlUtils.document;
import static com.monitorjbl.xlsx.XmlUtils.searchForNodeList;
import java.security.GeneralSecurityException;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * Streaming Excel workbook implementation. Most advanced features of POI are not supported.
 * Use this only if your application can handle iterating through an entire workbook, row by
 * row.
 */
public class StreamingReader implements Iterable<Row>, AutoCloseable {
  private static final Logger log = LoggerFactory.getLogger(StreamingReader.class);

  private final SharedStringsTable sst;
  private final StylesTable stylesTable;
  private final XMLEventReader parser;
  private final DataFormatter dataFormatter = new DataFormatter();

  private int rowCacheSize;
  private List<Row> rowCache = new ArrayList<>();
  private Iterator<Row> rowCacheIterator;

  private String lastContents;
  private StreamingRow currentRow;
  private StreamingCell currentCell;

  private File tmp;

  private StreamingReader(SharedStringsTable sst, StylesTable stylesTable, XMLEventReader parser, int rowCacheSize) {
    this.sst = sst;
    this.stylesTable = stylesTable;
    this.parser = parser;
    this.rowCacheSize = rowCacheSize;
  }

  /**
   * Read through a number of rows equal to the rowCacheSize field or until there is no more data to read
   *
   * @return true if data was read
   */
  private boolean getRow() {
    try {
      rowCache.clear();
      while(rowCache.size() < rowCacheSize && parser.hasNext()) {
        handleEvent(parser.nextEvent());
      }
      rowCacheIterator = rowCache.iterator();
      return rowCacheIterator.hasNext();
    } catch(XMLStreamException | SAXException e) {
      log.debug("End of stream");
    }
    return false;
  }

  /**
   * Handles a SAX event.
   *
   * @param event
   * @throws SAXException
   */
  private void handleEvent(XMLEvent event) throws SAXException {
    if(event.getEventType() == XMLStreamConstants.CHARACTERS) {
      Characters c = event.asCharacters();
      lastContents += c.getData();
    } else if(event.getEventType() == XMLStreamConstants.START_ELEMENT) {
      StartElement startElement = event.asStartElement();
      String tagLocalName = startElement.getName().getLocalPart();

      if("row".equals(tagLocalName)) {
        Attribute rowIndex = startElement.getAttributeByName(new QName("r"));
        currentRow = new StreamingRow(Integer.parseInt(rowIndex.getValue())-1);
      } else if("c".equals(tagLocalName)) {
        Attribute ref = startElement.getAttributeByName(new QName("r"));

        String[] coord = ref.getValue().split("(?<=\\D)(?=\\d)|(?<=\\d)(?=\\D)");
        currentCell = new StreamingCell(CellReference.convertColStringToIndex(coord[0]), Integer.parseInt(coord[1]) - 1);
        setFormatString(startElement, currentCell);

        Attribute type = startElement.getAttributeByName(new QName("t"));
        if(type != null) {
          currentCell.setType(type.getValue());
        }
      }

      // Clear contents cache
      lastContents = "";
    } else if(event.getEventType() == XMLStreamConstants.END_ELEMENT) {
      EndElement endElement = event.asEndElement();
      String tagLocalName = endElement.getName().getLocalPart();

      if("v".equals(tagLocalName)) {
        currentCell.setRawContents(unformattedContents());
        currentCell.setContents(formattedContents());
      } else if("row".equals(tagLocalName) && currentRow != null) {
        rowCache.add(currentRow);
      } else if("c".equals(tagLocalName)) {
        currentRow.getCellMap().put(currentCell.getColumnIndex(), currentCell);
      }

    }
  }

  /**
   * Read the numeric format string out of the styles table for this cell. Stores
   * the result in the Cell.
   *
   * @param startElement
   * @param cell
   */
  void setFormatString(StartElement startElement, StreamingCell cell) {
    Attribute cellStyle = startElement.getAttributeByName(new QName("s"));
    String cellStyleString = (cellStyle != null) ? cellStyle.getValue() : null;
    XSSFCellStyle style = null;

    if(cellStyleString != null) {
      style = stylesTable.getStyleAt(Integer.parseInt(cellStyleString));
    } else if(stylesTable.getNumCellStyles() > 0) {
      style = stylesTable.getStyleAt(0);
    }

    if(style != null) {
      cell.setNumericFormatIndex(style.getDataFormat());
      String formatString = style.getDataFormatString();

      if(formatString != null) {
        cell.setNumericFormat(formatString);
      } else {
        cell.setNumericFormat(BuiltinFormats.getBuiltinFormat(cell.getNumericFormatIndex()));
      }
    } else {
      cell.setNumericFormatIndex(null);
      cell.setNumericFormat(null);
    }
  }

  /**
   * Tries to format the contents of the last contents appropriately based on
   * the type of cell and the discovered numeric format.
   *
   * @return
   */
  String formattedContents() {
    switch(currentCell.getType() == null ? "" : currentCell.getType()) {
      case "s":           //string stored in shared table
        int idx = Integer.parseInt(lastContents);
        return new XSSFRichTextString(sst.getEntryAt(idx)).toString();
      case "inlineStr":   //inline string (not in sst)
        return new XSSFRichTextString(lastContents).toString();
      case "str":         //forumla type
        return '"' + lastContents + '"';
      case "e":           //error type
        return "ERROR:  " + lastContents;
      case "n":           //numeric type
        if(currentCell.getNumericFormat() != null && lastContents.length() > 0) {
          return dataFormatter.formatRawCellContents(
              Double.parseDouble(lastContents),
              currentCell.getNumericFormatIndex(),
              currentCell.getNumericFormat());
        } else {
          return lastContents;
        }
      default:
        return lastContents;
    }
  }

  /**
   * Returns the contents of the cell, with no formatting applied
   * @return
   */
  String unformattedContents(){
    switch(currentCell.getType() == null ? "" : currentCell.getType()) {
      case "s":           //string stored in shared table
        int idx = Integer.parseInt(lastContents);
        return new XSSFRichTextString(sst.getEntryAt(idx)).toString();
      case "inlineStr":   //inline string (not in sst)
        return new XSSFRichTextString(lastContents).toString();
      default:
        return lastContents;
    }
  }

  /**
   * Returns a new streaming iterator to loop through rows. This iterator is not
   * guaranteed to have all rows in memory, and any particular iteration may
   * trigger a load from disk to read in new data.
   *
   * @return the streaming iterator
   */
  @Override
  public Iterator<Row> iterator() {
    return new StreamingIterator();
  }

  /**
   * Closes the streaming resource, attempting to clean up any temporary files created.
   *
   * @throws com.monitorjbl.xlsx.exceptions.CloseException if there is an issue closing the stream
   */
  @Override
  public void close() {
    try {
      parser.close();
    } catch(XMLStreamException e) {
      throw new CloseException(e);
    }

    if(tmp != null) {
      log.debug("Deleting tmp file [" + tmp.getAbsolutePath() + "]");
      tmp.delete();
    }
  }

  static File writeInputStreamToFile(InputStream is, int bufferSize) throws IOException {
    File f = Files.createTempFile("tmp-", ".xlsx").toFile();
    try(FileOutputStream fos = new FileOutputStream(f)) {
      int read;
      byte[] bytes = new byte[bufferSize];
      while((read = is.read(bytes)) != -1) {
        fos.write(bytes, 0, read);
      }
      is.close();
      fos.close();
      return f;
    }
  }

  public static Builder builder() {
    return new Builder();
  }

  public static class Builder {
    int rowCacheSize = 10;
    int bufferSize = 1024;
    int sheetIndex = 0;
    String sheetName;
    String password;

    /**
     * The number of rows to keep in memory at any given point.
     * <p>
     * Defaults to 10
     * </p>
     *
     * @param rowCacheSize number of rows
     * @return reference to current {@code Builder}
     */
    public Builder rowCacheSize(int rowCacheSize) {
      this.rowCacheSize = rowCacheSize;
      return this;
    }

    /**
     * The number of bytes to read into memory from the input
     * resource.
     * <p>
     * Defaults to 1024
     * </p>
     *
     * @param bufferSize buffer size in bytes
     * @return reference to current {@code Builder}
     */
    public Builder bufferSize(int bufferSize) {
      this.bufferSize = bufferSize;
      return this;
    }

    /**
     * Which sheet to open. There can only be one sheet open
     * for a single instance of {@code StreamingReader}. If
     * more sheets need to be read, a new instance must be
     * created.
     * <p>
     * Defaults to 0
     * </p>
     *
     * @param sheetIndex index of sheet
     * @return reference to current {@code Builder}
     */
    public Builder sheetIndex(int sheetIndex) {
      this.sheetIndex = sheetIndex;
      return this;
    }

    /**
     * Which sheet to open. There can only be one sheet open
     * for a single instance of {@code StreamingReader}. If
     * more sheets need to be read, a new instance must be
     * created.
     *
     * @param sheetName name of sheet
     * @return reference to current {@code Builder}
     */
    public Builder sheetName(String sheetName) {
      this.sheetName = sheetName;
      return this;
    }

    /**
     * For password protected files specify password to open file.
     * If the password is incorrect a {@code ReadException} is thrown on
     * {@code read}.
     * 
     * <p>NULL indicates that no password should be used, this is the
     * default value.</p>
     *
     * @param password to use when opening file
     * @return reference to current {@code Builder}
     */
    public Builder password(String password) {
      this.password = password;
      return this;
    }
    
    /**
     * Reads a given {@code InputStream} and returns a new
     * instance of {@code StreamingReader}. Due to Apache POI
     * limitations, a temporary file must be written in order
     * to create a streaming iterator. This process will use
     * the same buffer size as specified in {@link #bufferSize(int)}.
     *
     * @param is input stream to read in
     * @return built streaming reader instance
     * @throws com.monitorjbl.xlsx.exceptions.ReadException if there is an issue reading the stream
     */
    public StreamingReader read(InputStream is) {
      File f = null;
      try {
        f = writeInputStreamToFile(is, bufferSize);
        log.debug("Created temp file [" + f.getAbsolutePath() + "]");

        StreamingReader r = read(f);
        r.tmp = f;
        return r;
      } catch(IOException e) {
        throw new ReadException("Unable to read input stream", e);
      } catch(RuntimeException e) {
        f.delete();
        throw e;
      }
    }

    /**
     * Reads a given {@code File} and returns a new instance
     * of {@code StreamingReader}.
     *
     * @param f file to read in
     * @return built streaming reader instance
     * @throws com.monitorjbl.xlsx.exceptions.OpenException if there is an issue opening the file
     * @throws com.monitorjbl.xlsx.exceptions.ReadException if there is an issue reading the file
     */
    public StreamingReader read(File f) {
      try {
        OPCPackage pkg;

        if (password != null) {
          // Based on: https://poi.apache.org/encryption.html
          POIFSFileSystem poifs = new POIFSFileSystem(f);
          EncryptionInfo info = new EncryptionInfo(poifs);
          Decryptor d = Decryptor.getInstance(info);
          d.verifyPassword(password);
          pkg = OPCPackage.open(d.getDataStream(poifs));
        } else {
          pkg = OPCPackage.open(f);
        }
        
        XSSFReader reader = new XSSFReader(pkg);
        SharedStringsTable sst = reader.getSharedStringsTable();
        StylesTable styles = reader.getStylesTable();

        InputStream sheet = findSheet(reader);
        if(sheet == null) {
          throw new MissingSheetException("Unable to find sheet at index [" + sheetIndex + "]");
        }

        XMLEventReader parser = XMLInputFactory.newInstance().createXMLEventReader(sheet);
        return new StreamingReader(sst, styles, parser, rowCacheSize);
      } catch(IOException e) {
        throw new OpenException("Failed to open file", e);
      } catch(OpenXML4JException | XMLStreamException e) {
        throw new ReadException("Unable to read workbook", e);
      } catch(GeneralSecurityException e) {
        throw new ReadException("Unable to read workbook - Decryption failed", e);
      }
    }

    InputStream findSheet(XSSFReader reader) throws IOException, InvalidFormatException {
      int index = sheetIndex;
      if(sheetName != null) {
        index = -1;
        //This file is separate from the worksheet data, and should be fairly small
        NodeList nl = searchForNodeList(document(reader.getWorkbookData()), "/workbook/sheets/sheet");
        for(int i = 0; i < nl.getLength(); i++) {
          if(Objects.equals(nl.item(i).getAttributes().getNamedItem("name").getTextContent(), sheetName)) {
            index = i;
          }
        }
        if(index < 0) {
          return null;
        }
      }
      Iterator<InputStream> iter = reader.getSheetsData();
      InputStream sheet = null;

      int i = 0;
      while(iter.hasNext()) {
        InputStream is = iter.next();
        if(i++ == index) {
          sheet = is;
          log.debug("Found sheet at index [" + sheetIndex + "]");
          break;
        }
      }
      return sheet;
    }
  }

  class StreamingIterator implements Iterator<Row> {
    public StreamingIterator() {
      if(rowCacheIterator == null){
        hasNext();
      }
    }

    @Override
    public boolean hasNext() {
      return (rowCacheIterator != null && rowCacheIterator.hasNext()) || getRow();
    }

    @Override
    public Row next() {
      return rowCacheIterator.next();
    }

    @Override
    public void remove() {
      throw new RuntimeException("NotSupported");
    }
  }

}
