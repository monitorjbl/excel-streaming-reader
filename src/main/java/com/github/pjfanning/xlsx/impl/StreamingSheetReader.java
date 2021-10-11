package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.StreamingReader;
import com.github.pjfanning.xlsx.XmlUtils;
import com.github.pjfanning.xlsx.exceptions.CloseException;
import com.github.pjfanning.xlsx.exceptions.NotSupportedException;
import com.github.pjfanning.xlsx.exceptions.ParseException;
import com.github.pjfanning.xlsx.impl.ooxml.HyperlinkData;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.*;
import java.time.LocalDateTime;
import java.util.*;

public class StreamingSheetReader implements Iterable<Row> {
  private static final Logger log = LoggerFactory.getLogger(StreamingSheetReader.class);

  private final StreamingWorkbookReader streamingWorkbookReader;
  private final PackagePart packagePart;
  private final SharedStringsTable sst;
  private final StylesTable stylesTable;
  private final Comments commentsTable;
  private final XMLEventReader parser;
  private final DataFormatter dataFormatter = new DataFormatter();
  private final Set<Integer> hiddenColumns = new HashSet<>();
  private final List<CellRangeAddress> mergedCells = new ArrayList<>();
  private final List<HyperlinkData> hyperlinks = new ArrayList<>();
  private List<XlsxHyperlink> xlsxHyperlinks;

  private int firstRowNum = 0;
  private int lastRowNum;
  private int currentRowNum;
  private int firstColNum = 0;
  private int currentColNum;
  private int rowCacheSize;
  private List<Row> rowCache = new ArrayList<>();
  private Iterator<Row> rowCacheIterator;

  private StringBuilder contentBuilder = new StringBuilder(64);
  private StringBuilder formulaBuilder = new StringBuilder(64);
  private StreamingSheet sheet;
  private StreamingRow currentRow;
  private StreamingCell currentCell;
  private CellAddress activeCell;
  private boolean use1904Dates;
  private boolean insideCharElement = false;
  private boolean insideFormulaElement = false;
  private boolean insideIS = false;

  StreamingSheetReader(StreamingWorkbookReader streamingWorkbookReader,
                       PackagePart packagePart,
                       SharedStringsTable sst, StylesTable stylesTable, Comments commentsTable,
                       XMLEventReader parser, final boolean use1904Dates, int rowCacheSize) {
    this.streamingWorkbookReader = streamingWorkbookReader;
    this.packagePart = packagePart;
    this.sst = sst;
    this.stylesTable = stylesTable;
    this.commentsTable = commentsTable;
    this.parser = parser;
    this.use1904Dates = use1904Dates;
    this.rowCacheSize = rowCacheSize;
  }

  void setSheet(StreamingSheet sheet) {
    this.sheet = sheet;
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
      throw new ParseException("Error reading XML stream", e);
    }
  }

  private String[] splitCellRef(String ref) {
    int splitPos = -1;

    // start at pos 1, since the first char is expected to always be a letter
    for(int i=1;i<ref.length();i++) {
      char c = ref.charAt(i);

      if (c >= '0' && c <= '9') {
        splitPos = i;
        break;
      }
    }

    return new String[] {
            ref.substring(0, splitPos),
            ref.substring(splitPos)
    };
  }

  /**
   * Handles a SAX event.
   *
   * @param event
   * @throws SAXException
   */
  private void handleEvent(XMLEvent event) throws SAXException {
    if(event.getEventType() == XMLStreamConstants.CHARACTERS) {
      if (insideCharElement) {
        contentBuilder.append(event.asCharacters().getData());
      }
      if (insideFormulaElement) {
        Characters c = event.asCharacters();
        formulaBuilder.append(event.asCharacters().getData());
      }
    } else if(event.getEventType() == XMLStreamConstants.START_ELEMENT
        && isSpreadsheetTag(event.asStartElement().getName())) {
      StartElement startElement = event.asStartElement();
      String tagLocalName = startElement.getName().getLocalPart();

      if("row".equals(tagLocalName)) {
        Attribute rowNumAttr = startElement.getAttributeByName(new QName("r"));
        int rowIndex = currentRowNum;
        if(rowNumAttr != null) {
          rowIndex = Integer.parseInt(rowNumAttr.getValue()) - 1;
          currentRowNum = rowIndex;
        }
        Attribute isHiddenAttr = startElement.getAttributeByName(new QName("hidden"));
        boolean isHidden = isHiddenAttr != null && XmlUtils.evaluateBoolean(isHiddenAttr.getValue());
        currentRow = new StreamingRow(sheet, rowIndex, isHidden);
        currentColNum = firstColNum;
      } else if("col".equals(tagLocalName)) {
        Attribute isHiddenAttr = startElement.getAttributeByName(new QName("hidden"));
        boolean isHidden = isHiddenAttr != null && XmlUtils.evaluateBoolean(isHiddenAttr.getValue());
        if(isHidden) {
          Attribute minAttr = startElement.getAttributeByName(new QName("min"));
          Attribute maxAttr = startElement.getAttributeByName(new QName("max"));
          int min = Integer.parseInt(minAttr.getValue()) - 1;
          int max = Integer.parseInt(maxAttr.getValue()) - 1;
          for(int columnIndex = min; columnIndex <= max; columnIndex++)
            hiddenColumns.add(columnIndex);
        }
      } else if("c".equals(tagLocalName)) {
        Attribute ref = startElement.getAttributeByName(new QName("r"));

        if (ref != null) {
          String[] coord = splitCellRef(ref.getValue());
          currentColNum = CellReference.convertColStringToIndex(coord[0]);
          if (currentRow.getRowNum() == currentRowNum) {
            currentCell = new StreamingCell(sheet, currentColNum, currentRow, use1904Dates);
          } else {
            currentCell = new StreamingCell(sheet, currentColNum, Integer.parseInt(coord[1]) - 1, use1904Dates);
          }
        } else if (currentRow != null) {
          currentCell = new StreamingCell(sheet, currentColNum, currentRow, use1904Dates);
        } else {
          currentCell = new StreamingCell(sheet, currentColNum, currentRowNum, use1904Dates);
        }
        setFormatString(startElement, currentCell);

        Attribute type = startElement.getAttributeByName(new QName("t"));
        if (type != null) {
          currentCell.setType(type.getValue());
        } else {
          currentCell.setType("n");
        }

        if (stylesTable != null) {
          Attribute style = startElement.getAttributeByName(new QName("s"));
          if (style != null) {
            String indexStr = style.getValue();
            try {
              int index = Integer.parseInt(indexStr);
              currentCell.setCellStyle(stylesTable.getStyleAt(index));
            } catch (NumberFormatException nfe) {
              log.warn("Ignoring invalid style index {}", indexStr);
            }
          } else {
            currentCell.setCellStyle(stylesTable.getStyleAt(0));
          }
        }
      } else if("v".equals(tagLocalName) || "t".equals(tagLocalName)) {
        insideCharElement = true;
      } else if("is".equals(tagLocalName)) {
        insideIS = true;
      } else if("dimension".equals(tagLocalName)) {
        Attribute refAttr = startElement.getAttributeByName(new QName("ref"));
        String ref = refAttr != null ? refAttr.getValue() : null;
        if(ref != null) {
          // ref is formatted as A1 or A1:F25. Take the last numbers of this string and use it as lastRowNum
          for(int i = ref.length() - 1; i >= 0; i--) {
            if(!Character.isDigit(ref.charAt(i))) {
              try {
                lastRowNum = Integer.parseInt(ref.substring(i + 1)) - 1;
              } catch(NumberFormatException ignore) { }
              break;
            }
          }
          int colonPos = ref.indexOf(':');
          if (colonPos > 0) {
            String firstPart = ref.substring(0, colonPos);
            try {
              CellReference cellReference = new CellReference(firstPart);
              firstRowNum = cellReference.getRow();
            } catch(Exception e) {
              log.warn("Failed to parse cell reference {}", firstPart);
            }
          }
          for(int i = 0; i < ref.length(); i++) {
            if(!Character.isAlphabetic(ref.charAt(i))) {
              firstColNum = CellReference.convertColStringToIndex(ref.substring(0, i));
              break;
            }
          }
        }
      } else if("f".equals(tagLocalName)) {
        insideFormulaElement = true;
        if (currentCell != null) {
          currentCell.setFormulaType(true);
        }
      } else if("mergeCell".equals(tagLocalName)) {
        Attribute ref = startElement.getAttributeByName(QName.valueOf("ref"));
        if(ref != null) {
          mergedCells.add(CellRangeAddress.valueOf(ref.getValue()));
        }
      } else if("selection".equals(tagLocalName)) {
      Attribute activeCell = startElement.getAttributeByName(QName.valueOf("activeCell"));
      if(activeCell != null) {
        String activeCellRef = getAttributeValue(activeCell);
        try {
          this.activeCell = new CellAddress(activeCellRef);
        } catch (Throwable t) {
          log.warn("unable to parse active cell reference {}", activeCellRef);
        }
      }
    } else if("hyperlink".equals(tagLocalName)) {
        String id = null;
        Iterator<Attribute> attributeIterator = startElement.getAttributes();
        while (attributeIterator.hasNext()) {
          Attribute att = attributeIterator.next();
          QName qn = att.getName();
          if ("id".equals(qn.getLocalPart()) && qn.getNamespaceURI().endsWith("relationships")) {
            id = att.getValue();
          }
        }
        Attribute ref = startElement.getAttributeByName(QName.valueOf("ref"));
        Attribute location = startElement.getAttributeByName(QName.valueOf("location"));
        Attribute display = startElement.getAttributeByName(QName.valueOf("display"));
        Attribute tooltip = startElement.getAttributeByName(QName.valueOf("tooltip"));
        hyperlinks.add(new HyperlinkData(id, getAttributeValue(ref), getAttributeValue(location),
                getAttributeValue(display), getAttributeValue(tooltip)));
      }

      if (!insideIS) {
        contentBuilder.setLength(0);
      }
      formulaBuilder.setLength(0);
    } else if(event.getEventType() == XMLStreamConstants.END_ELEMENT
        && isSpreadsheetTag(event.asEndElement().getName())) {
      EndElement endElement = event.asEndElement();
      String tagLocalName = endElement.getName().getLocalPart();

      if("v".equals(tagLocalName) || "t".equals(tagLocalName)) {
        insideCharElement = false;
        currentCell.setRawContents(unformattedContents());
        currentCell.setContentSupplier(formattedContents());
      } else if("row".equals(tagLocalName) && currentRow != null) {
        rowCache.add(currentRow);
        currentRowNum++;
      } else if("c".equals(tagLocalName)) {
        currentRow.getCellMap().put(currentCell.getColumnIndex(), currentCell);
        currentCell = null;
        currentColNum++;
      } else if("is".equals(tagLocalName)) {
        insideIS = false;
      } else if("f".equals(tagLocalName)) {
        insideFormulaElement = true;
        if (currentCell != null) {
          currentCell.setFormula(formulaBuilder.toString());
        }
      }

    }
  }

  /**
   * Returns true if a tag is part of the main namespace for SpreadsheetML:
   * <ul>
   * <li>http://schemas.openxmlformats.org/spreadsheetml/2006/main
   * <li>http://purl.oclc.org/ooxml/spreadsheetml/main
   * </ul>
   * As opposed to http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing, etc.
   */
  private boolean isSpreadsheetTag(QName name) {
    return (name.getNamespaceURI() != null
        && name.getNamespaceURI().endsWith("/main"));
  }

  /**
   * Get the hidden state for a given column
   *
   * @param columnIndex - the column to set (0-based)
   * @return hidden - <code>false</code> if the column is visible
   */
  boolean isColumnHidden(int columnIndex) {
    if(rowCacheIterator == null) {
      getRow();
    }
    return hiddenColumns.contains(columnIndex);
  }

  /**
   * Gets the last row on the sheet
   */
  int getFirstRowNum() {
    if(rowCacheIterator == null) {
      getRow();
    }
    return firstRowNum;
  }

  /**
   * Gets the last row on the sheet
   */
  int getLastRowNum() {
    if(rowCacheIterator == null) {
      getRow();
    }
    return lastRowNum;
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

    if (stylesTable != null) {
      if(cellStyleString != null) {
        style = stylesTable.getStyleAt(Integer.parseInt(cellStyleString));
      } else if(stylesTable.getNumCellStyles() > 0) {
        style = stylesTable.getStyleAt(0);
      }
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

  CellAddress getActiveCell() {
    return activeCell;
  }

  /**
   * Tries to format the contents of the last contents appropriately based on
   * the type of cell and the discovered numeric format.
   */
  Supplier formattedContents() {
    return getFormatterForType(currentCell.getType());
  }

  /**
   * Tries to format the contents of the last contents appropriately based on
   * the provided type and the discovered numeric format.
   */
  private Supplier getFormatterForType(String type) {
    final String lastContents = contentBuilder.toString();
    switch(type) {
      case "s":           //string stored in shared table
        if (!lastContents.isEmpty()) {
            int idx = Integer.parseInt(lastContents);
            return new StringSupplier(sst.getItemAt(idx).toString());
        }
        return new StringSupplier(lastContents);
      case "inlineStr":   //inline string (not in sst)
      case "str":
        return new StringSupplier(new XSSFRichTextString(lastContents).toString());
      case "e":           //error type
        return new StringSupplier("ERROR:  " + lastContents);
      case "n":           //numeric type
        if(currentCell.getNumericFormat() != null && lastContents.length() > 0) {
          // the formatRawCellContents operation incurs a significant overhead on large sheets,
          // and we want to defer the execution of this method until the value is actually needed.
          // it is not needed in all cases..
          final String currentLastContents = lastContents;
          final int currentNumericFormatIndex = currentCell.getNumericFormatIndex();
          final String currentNumericFormat = currentCell.getNumericFormat();

          return new Supplier() {
            String cachedContent;

            @Override
            public Object getContent() {
              if (cachedContent == null) {
                cachedContent = dataFormatter.formatRawCellContents(
                        Double.parseDouble(currentLastContents),
                        currentNumericFormatIndex,
                        currentNumericFormat);
              }

              return cachedContent;
            }
          };
        } else {
          return new StringSupplier(lastContents);
        }
      case "d":           //date type (Strict OOXML format)
        if(currentCell.getNumericFormat() != null && lastContents.length() > 0) {
          // the formatRawCellContents operation incurs a significant overhead on large sheets,
          // and we want to defer the execution of this method until the value is actually needed.
          // it is not needed in all cases..
          final String currentLastContents = lastContents;
          final int currentNumericFormatIndex = currentCell.getNumericFormatIndex();
          final String currentNumericFormat = currentCell.getNumericFormat();

          return new Supplier() {
            String cachedContent;

            @Override
            public Object getContent() {
              if (cachedContent == null) {
                try {
                  Double dv;
                  try {
                    LocalDateTime dt = DateTimeUtil.parseDateTime(currentLastContents);
                    dv = DateUtil.getExcelDate(dt, use1904Dates);
                  } catch (Exception e) {
                    dv = DateTimeUtil.convertTime(currentLastContents);
                  }
                  cachedContent = dataFormatter.formatRawCellContents(
                          dv,
                          currentNumericFormatIndex,
                          currentNumericFormat);
                } catch (Exception e) {
                  log.warn("cannot format strict format date/time " + currentLastContents);
                  cachedContent = currentLastContents;
                }
              }
              return cachedContent;
            }
          };
        } else {
          return new StringSupplier(lastContents);
        }
      default:
        return new StringSupplier(lastContents);
    }
  }

  /**
   * Returns the contents of the cell, with no formatting applied
   */
  String unformattedContents() {
    final String lastContents = contentBuilder.toString();
    switch(currentCell.getType()) {
      case "s":           //string stored in shared table
        if (!lastContents.isEmpty()) {
            int idx = Integer.parseInt(lastContents);
            if (sst == null) throw new NullPointerException("sst is null");
            return sst.getItemAt(idx).toString();
        }
        return lastContents;
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
    return new StreamingRowIterator();
  }

  /**
   * @return the comments associated with this sheet (only if feature is enabled on the Builder)
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadComments(boolean)} is not set to true
   */
  Comments getCellComments() {
    if (!streamingWorkbookReader.getBuilder().readComments()) {
      throw new IllegalStateException("getCellComments() only works if StreamingWorking.Builder setReadComments is set to true");
    }
    return this.commentsTable;
  }

  List<CellRangeAddress> getMergedCells() { return this.mergedCells; }

  XSSFDrawing getDrawingPatriarch() {
    if (!streamingWorkbookReader.getBuilder().readShapes()) {
      throw new IllegalStateException("getDrawingPatriarch() only works if StreamingWorking.Builder setReadShapes is set to true");
    }
    if (sheet != null) {
      List<XSSFShape> shapes = streamingWorkbookReader.getShapes(sheet.getSheetName());
      if (shapes != null) {
        Iterator<XSSFShape> shapesIter = shapes.iterator();
        while (shapesIter.hasNext()) {
          return shapesIter.next().getDrawing();
        }
      }
    }
    return null;
  }

  public void close() {
    try {
      parser.close();
    } catch(XMLStreamException e) {
      throw new CloseException(e);
    }
  }

  StreamingReader.Builder getBuilder() {
    return streamingWorkbookReader.getBuilder();
  }

  /**
   * @return the hyperlinks associated with this sheet (only if feature is enabled on the Builder)
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadHyperlinks(boolean)} is not set to true
   */
  List<XlsxHyperlink> getHyperlinks() {
    if (!getBuilder().readHyperlinks()) {
      throw new IllegalStateException("getHyperlinks() only works if StreamingWorking.Builder setReadHyperlinks is set to true");
    }
    initHyperlinks();
    return xlsxHyperlinks;
  }

  private String getAttributeValue(Attribute att) {
    return att == null ? null : att.getValue();
  }

  private void initHyperlinks() {
    if (xlsxHyperlinks == null || xlsxHyperlinks.isEmpty()) {
      ArrayList<XlsxHyperlink> links = new ArrayList<>();

      try {
        PackageRelationshipCollection hyperRels =
                packagePart.getRelationshipsByType(XSSFRelation.SHEET_HYPERLINKS.getRelation());

        // Turn each one into a XSSFHyperlink
        for(HyperlinkData hyperlink : hyperlinks) {
          PackageRelationship hyperRel = null;
          if(hyperlink.getId() != null) {
            hyperRel = hyperRels.getRelationshipByID(hyperlink.getId());
          }

          links.add( new XlsxHyperlink(hyperlink, hyperRel) );
        }
      } catch (InvalidFormatException e){
        throw new POIXMLException(e);
      }
      xlsxHyperlinks = links;
    }
  }

  class StreamingRowIterator implements Iterator<Row> {
    public StreamingRowIterator() {
      if(rowCacheIterator == null) {
        hasNext();
      }
    }

    @Override
    public boolean hasNext() {
      return (rowCacheIterator != null && rowCacheIterator.hasNext()) || getRow();
    }

    @Override
    public Row next() {
      try {
        return rowCacheIterator.next();
      } catch(NoSuchElementException nsee) {
        //see https://github.com/monitorjbl/excel-streaming-reader/issues/176
        if (hasNext()) {
          return rowCacheIterator.next();
        }
        throw nsee;
      }
    }

    @Override
    public void remove() {
      throw new NotSupportedException();
    }
  }
}
