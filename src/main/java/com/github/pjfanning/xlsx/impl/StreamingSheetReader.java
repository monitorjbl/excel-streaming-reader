package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.CloseableIterator;
import com.github.pjfanning.xlsx.SharedFormula;
import com.github.pjfanning.xlsx.StreamingReader;
import com.github.pjfanning.xlsx.exceptions.OpenException;
import com.github.pjfanning.xlsx.exceptions.ReadException;
import com.github.pjfanning.xlsx.impl.ooxml.HyperlinkData;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFShape;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.StartElement;
import java.io.IOException;
import java.util.*;

public class StreamingSheetReader implements Iterable<Row> {
  private static final Logger LOG = LoggerFactory.getLogger(StreamingSheetReader.class);

  private static XMLInputFactory xmlInputFactory;

  private final StreamingWorkbookReader streamingWorkbookReader;
  private final PackagePart packagePart;
  private final SharedStrings sst;
  private final StylesTable stylesTable;
  private final Comments commentsTable;
  private final boolean use1904Dates;
  private final int rowCacheSize;
  private final Set<Integer> hiddenColumns = new HashSet<>();
  private final Map<Integer, Float> columnWidths = new HashMap<>();
  private final Set<CellRangeAddress> mergedCells = new LinkedHashSet<>();  // use HashSet to prevent duplicates
  private final List<StreamingRowIterator> iterators = new ArrayList<>();
  private final Set<HyperlinkData> hyperlinks = new LinkedHashSet<>();  // use HashSet to prevent duplicates

  private List<XlsxHyperlink> xlsxHyperlinks;
  private Map<String, SharedFormula> sharedFormulaMap;
  private int firstRowNum;
  private int lastRowNum;
  private float defaultRowHeight;
  private int baseColWidth = 8; //POI XSSFSheet default
  private StreamingSheet sheet;
  private CellAddress activeCell;
  private PaneInformation pane;

  StreamingSheetReader(final StreamingWorkbookReader streamingWorkbookReader,
                       final PackagePart packagePart,
                       final SharedStrings sst, final StylesTable stylesTable, final Comments commentsTable,
                       final boolean use1904Dates, final int rowCacheSize) {
    this.streamingWorkbookReader = streamingWorkbookReader;
    this.packagePart = packagePart;
    this.sst = sst;
    this.stylesTable = stylesTable;
    this.commentsTable = commentsTable;
    this.use1904Dates = use1904Dates;
    this.rowCacheSize = rowCacheSize;
  }

  void setSheet(StreamingSheet sheet) {
    this.sheet = sheet;
  }

  void removeIterator(StreamingRowIterator iterator) {
    iterators.remove(iterator);
  }

  Map<String, SharedFormula> getSharedFormulaMap() {
    if (getBuilder().readSharedFormulas()) {
      if (sharedFormulaMap == null) {
        return Collections.emptyMap();
      }
      return Collections.unmodifiableMap(sharedFormulaMap);
    } else {
      throw new IllegalStateException("The reading of shared formulas has been disabled. Enable using StreamingReader.Builder.");
    }
  }

  void addSharedFormula(String siValue, SharedFormula sharedFormula) {
    if (getBuilder().readSharedFormulas()) {
      if (sharedFormulaMap == null) {
        sharedFormulaMap = new HashMap<>();
      }
      sharedFormulaMap.put(siValue, sharedFormula);
    }
  }

  SharedFormula removeSharedFormula(String siValue) {
    if (sharedFormulaMap != null) {
      return sharedFormulaMap.remove(siValue);
    }
    return null;
  }

  boolean isUse1904Dates() {
    return use1904Dates;
  }

  float getDefaultRowHeight() {
    return defaultRowHeight;
  }

  void setDefaultRowHeight(float defaultRowHeight) {
    this.defaultRowHeight = defaultRowHeight;
  }

  int getBaseColWidth() {
    return baseColWidth;
  }

  void setBaseColWidth(int baseColWidth) {
    this.baseColWidth = baseColWidth;
  }

  /**
   * Get the hidden state for a given column
   *
   * @param columnIndex - the column to set (0-based)
   * @return hidden - <code>false</code> if the column is visible
   */
  boolean isColumnHidden(int columnIndex) {
    if (iterators.isEmpty()) {
      // create a new streaming iterator to parse sheet
      iterator();
    }
    return hiddenColumns.contains(columnIndex);
  }

  float getColumnWidth(int columnIndex) {
    if (iterators.isEmpty()) {
      // create a new streaming iterator to parse sheet
      iterator();
    }
    Float width = columnWidths.get(columnIndex);
    return width == null ? getBaseColWidth() : width;
  }

  /**
   * Gets the first row on the sheet
   */
  int getFirstRowNum() {
    if (iterators.isEmpty()) {
      // create a new streaming iterator to parse sheet
      iterator();
    }
    return firstRowNum;
  }

  void setFirstRowNum(int firstRowNum) {
    this.firstRowNum = firstRowNum;
  }

  /**
   * Gets the last row on the sheet
   */
  int getLastRowNum() {
    if (iterators.isEmpty()) {
      // create a new streaming iterator to parse sheet
      iterator();
    }
    return lastRowNum;
  }

  void setLastRowNum(int lastRowNum) {
    this.lastRowNum = lastRowNum;
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

  void setActiveCell(CellAddress activeCell) {
    this.activeCell = activeCell;
  }

  PaneInformation getPane() {
    if (iterators.isEmpty()) {
      // create a new streaming iterator to parse sheet
      iterator();
    }
    return pane;
  }

  void setPane(PaneInformation pane) {
    this.pane = pane;
  }

  /**
   * Returns a new streaming iterator to loop through rows. This iterator is not
   * guaranteed to have all rows in memory, and any particular iteration may
   * trigger a load from disk to read in new data.
   *
   * This is an iterator of the PHYSICAL rows.
   * Meaning the 3rd element may not be the third row if say for instance the second row is undefined.
   *
   * This behaviour changed in v4.0.0. Earlier versions only created one iterator and repeated
   * calls to this method just returned the same iterator. Creating multiple iterators will slow down
   * your application and should be avoided unless necessary.
   *
   * @return the streaming iterator, an instance of {@link CloseableIterator} -
   * it is recommended that you close the iterator when finished with it if you intend to keep the sheet open.
   */
  @Override
  public CloseableIterator<Row> iterator() {
    try {
      //StreamingRowIterator requires a new XMLEventReader with a new InputStream to be provided to start from the
      //beginning of the Sheet
      XMLEventReader parser = getXmlInputFactory().createXMLEventReader(packagePart.getInputStream());
      StreamingRowIterator iterator = new StreamingRowIterator(this,
              sst, stylesTable, parser, use1904Dates, rowCacheSize, hiddenColumns, columnWidths, mergedCells, hyperlinks,
              sharedFormulaMap, defaultRowHeight, sheet);
      iterators.add(iterator);
      return iterator;
    } catch (IOException e) {
      throw new OpenException("Failed to open stream", e);
    } catch (XMLStreamException e) {
      throw new ReadException("Unable to read sheet", e);
    }
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

  List<CellRangeAddress> getMergedCells() {
    return new ArrayList<>(this.mergedCells);
  }

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
    iterators.forEach(iter -> iter.close(false));
  }

  StreamingReader.Builder getBuilder() {
    return streamingWorkbookReader.getBuilder();
  }

  Workbook getWorkbook() {
    return streamingWorkbookReader.getWorkbook();
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

  private static XMLInputFactory getXmlInputFactory() {
    if (xmlInputFactory == null) {
      try {
        xmlInputFactory = XMLHelper.newXMLInputFactory();
      } catch (Exception e) {
        LOG.error("Issue creating XMLInputFactory", e);
        throw e;
      }
    }
    return xmlInputFactory;
  }
}
