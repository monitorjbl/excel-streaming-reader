package com.monitorjbl.xlsx.impl;

import com.monitorjbl.xlsx.exceptions.CloseException;
import com.monitorjbl.xlsx.exceptions.ParseException;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.Characters;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;


import static com.monitorjbl.xlsx.impl.CellTypeConstants.STRING;
import static com.monitorjbl.xlsx.impl.CellTypeConstants.INLINE_STR;
import static com.monitorjbl.xlsx.impl.CellTypeConstants.NUMERIC;
import static com.monitorjbl.xlsx.impl.CellTypeConstants.ERROR;
import static com.monitorjbl.xlsx.impl.CellTypeConstants.STR;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_C;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_F;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_R;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_S;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_T;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_V;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_COL;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_DIMENSION;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_HIDDEN;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_MAX;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_MIN;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_ROW;
import static com.monitorjbl.xlsx.impl.LocalPartConstants.LOCAL_PART_REF;

public class StreamingSheetReader implements Iterable<Row> {
  private static final Logger log = LoggerFactory.getLogger(StreamingSheetReader.class);

  private final SharedStrings sst;
  private final StylesTable stylesTable;
  private final XMLEventReader parser;
  private final DataFormatter dataFormatter = new DataFormatter();
  private final Set<Integer> hiddenColumns = new HashSet<>();
  private final int rowCacheSize;
  private final List<Row> rowCache;
  private final boolean use1904Dates;
  private int lastRowNum;
  private int currentRowNum;
  private int firstColNum = 0;
  private int currentColNum;
  private Iterator<Row> rowCacheIterator;
  private String lastContents;
  private Sheet sheet;
  private StreamingRow currentRow;
  private StreamingCell currentCell;

  public StreamingSheetReader(SharedStrings sst, StylesTable stylesTable, XMLEventReader parser,
                              final boolean use1904Dates, int rowCacheSize) {
    this.sst = sst;
    this.stylesTable = stylesTable;
    this.parser = parser;
    this.use1904Dates = use1904Dates;
    this.rowCacheSize = rowCacheSize;
    rowCache = new ArrayList<>(rowCacheSize);
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
    } catch(XMLStreamException e) {
      throw new ParseException("Error reading XML stream", e);
    }
  }

  private String[] splitCellRef(String ref) {
    int splitPos = -1;

    // start at pos 1, since the first char is expected to always be a letter
    for(int i = 1; i < ref.length(); i++) {
      char c = ref.charAt(i);

      if(c >= '0' && c <= '9') {
        splitPos = i;
        break;
      }
    }

    return new String[]{
        ref.substring(0, splitPos),
        ref.substring(splitPos)
    };
  }

  /**
   * Handles a SAX event.
   *
   * @param event
   */
  private void handleEvent(XMLEvent event) {
    if(event.getEventType() == XMLStreamConstants.CHARACTERS) {
      Characters c = event.asCharacters();
      lastContents += c.getData();
    } else if(event.getEventType() == XMLStreamConstants.START_ELEMENT
        && isSpreadsheetTag(event.asStartElement().getName())) {
      StartElement startElement = event.asStartElement();
      final String tagLocalName = startElement.getName().getLocalPart();

      if (LOCAL_PART_ROW.isConstantEquals(tagLocalName)) {
        Attribute rowNumAttr = startElement.getAttributeByName(LOCAL_PART_R.getQname());
        int rowIndex = currentRowNum;
        if (rowNumAttr != null) {
          rowIndex = Integer.parseInt(rowNumAttr.getValue()) - 1;
          currentRowNum = rowIndex;
        }
        Attribute isHiddenAttr = startElement.getAttributeByName(LOCAL_PART_HIDDEN.getQname());
        boolean isHidden = isHiddenAttr != null && ("1".equals(isHiddenAttr.getValue()) || "true".equals(isHiddenAttr.getValue()));
        currentRow = new StreamingRow(sheet, rowIndex, isHidden);
        currentColNum = firstColNum;
      } else if (LOCAL_PART_COL.isConstantEquals(tagLocalName)) {
        Attribute isHiddenAttr = startElement.getAttributeByName(LOCAL_PART_HIDDEN.getQname());
        boolean isHidden = isHiddenAttr != null && ("1".equals(isHiddenAttr.getValue()) || "true".equals(isHiddenAttr.getValue()));
        if (isHidden) {
          Attribute minAttr = startElement.getAttributeByName(LOCAL_PART_MIN.getQname());
          Attribute maxAttr = startElement.getAttributeByName(LOCAL_PART_MAX.getQname());
          int min = Integer.parseInt(minAttr.getValue()) - 1;
          int max = Integer.parseInt(maxAttr.getValue()) - 1;
          for (int columnIndex = min; columnIndex <= max; columnIndex++)
            hiddenColumns.add(columnIndex);
        }
      } else if (LOCAL_PART_C.isConstantEquals(tagLocalName)) {
        Attribute ref = startElement.getAttributeByName(LOCAL_PART_R.getQname());

        if (ref != null) {
          String[] coord = splitCellRef(ref.getValue());
          currentColNum = CellReference.convertColStringToIndex(coord[0]);
          currentCell = new StreamingCell(sheet, currentColNum, Integer.parseInt(coord[1]) - 1, use1904Dates);
        } else {
          currentCell = new StreamingCell(sheet, currentColNum, currentRowNum, use1904Dates);
        }
        setFormatString(startElement, currentCell);

        Attribute type = startElement.getAttributeByName(LOCAL_PART_T.getQname());
        if(type != null) {
          currentCell.setType(type.getValue());
        } else {
          currentCell.setType("n");
        }

        Attribute style = startElement.getAttributeByName(LOCAL_PART_S.getQname());
        if(style != null) {
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
      } else if (LOCAL_PART_DIMENSION.isConstantEquals(tagLocalName)) {
        Attribute refAttr = startElement.getAttributeByName(LOCAL_PART_REF.getQname());
        String ref = refAttr != null ? refAttr.getValue() : null;
        if (ref != null) {
          // ref is formatted as A1 or A1:F25. Take the last numbers of this string and use it as lastRowNum
          for (int i = ref.length() - 1; i >= 0; i--) {
            if (!Character.isDigit(ref.charAt(i))) {
              try {
                lastRowNum = Integer.parseInt(ref.substring(i + 1)) - 1;
              } catch (NumberFormatException ignore) {
              }
              break;
            }
          }
          for (int i = 0; i < ref.length(); i++) {
            if (!Character.isAlphabetic(ref.charAt(i))) {
              firstColNum = CellReference.convertColStringToIndex(ref.substring(0, i));
              break;
            }
          }
        }
      } else if (LOCAL_PART_F.isConstantEquals(tagLocalName)) {
        if (currentCell != null) {
          currentCell.setFormulaType(true);
        }
      }

      // Clear contents cache
      lastContents = "";
    } else if(event.getEventType() == XMLStreamConstants.END_ELEMENT
        && isSpreadsheetTag(event.asEndElement().getName())) {
      EndElement endElement = event.asEndElement();
      final String tagLocalName = endElement.getName().getLocalPart();

      if (LOCAL_PART_V.isConstantEquals(tagLocalName) || LOCAL_PART_T.isConstantEquals(tagLocalName)) {
        currentCell.setRawContents(unformattedContents());
        currentCell.setContentSupplier(formattedContents());
      } else if (LOCAL_PART_ROW.isConstantEquals(tagLocalName) && currentRow != null) {
        rowCache.add(currentRow);
        currentRowNum++;
      } else if (LOCAL_PART_C.isConstantEquals(tagLocalName)) {
        assert currentRow != null;
        currentRow.getCellMap().put(currentCell.getColumnIndex(), currentCell);
        currentCell = null;
        currentColNum++;
      } else if (LOCAL_PART_F.isConstantEquals(tagLocalName)) {
        if (currentCell != null) {
          currentCell.setFormula(lastContents);
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
   *
   * @param name
   * @return
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
   *
   * @return
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
    Attribute cellStyle = startElement.getAttributeByName(LOCAL_PART_S.getQname());
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
  Supplier formattedContents() {
    return getFormatterForType(currentCell.getType());
  }

  /**
   * Tries to format the contents of the last contents appropriately based on
   * the provided type and the discovered numeric format.
   *
   * @return
   */
  private Supplier getFormatterForType(String type) {
    switch(type) {
      case STRING:           //string stored in shared table
        if (!lastContents.isEmpty()) {
          int idx = Integer.parseInt(lastContents);
          return new StringSupplier(sst.getItemAt(idx).toString());
        }
        return new StringSupplier(lastContents);
      case INLINE_STR:   //inline string (not in sst)
      case STR:
        return new StringSupplier(new XSSFRichTextString(lastContents).toString());
      case ERROR:           //error type
        return new StringSupplier("ERROR:  " + lastContents);
      case NUMERIC:           //numeric type
        if (currentCell.getNumericFormat() != null && lastContents.length() > 0) {
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
              if(cachedContent == null) {
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
      default:
        return new StringSupplier(lastContents);
    }
  }

  /**
   * Returns the contents of the cell, with no formatting applied
   *
   * @return
   */
  String unformattedContents() {
    switch(currentCell.getType()) {
      case STRING:           //string stored in shared table
        if(!lastContents.isEmpty()) {
          int idx = Integer.parseInt(lastContents);
          return sst.getItemAt(idx).toString();
        }
        return lastContents;
      case INLINE_STR:   //inline string (not in sst)
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

  public void close() {
    try {
      parser.close();
    } catch(XMLStreamException e) {
      throw new CloseException(e);
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
      return rowCacheIterator.next();
    }

    @Override
    public void remove() {
      throw new RuntimeException("NotSupported");
    }
  }
}
