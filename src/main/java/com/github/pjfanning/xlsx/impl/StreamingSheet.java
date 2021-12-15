package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.SharedFormula;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.util.*;

public class StreamingSheet implements Sheet {

  private final String name;
  private final StreamingSheetReader reader;

  public StreamingSheet(String name, StreamingSheetReader reader) {
    this.name = name;
    this.reader = reader;
    reader.setSheet(this);
  }

  StreamingSheetReader getReader() {
    return reader;
  }

  /* Supported */

  /**
   * Workbook is only set under certain usage flows.
   */
  @Override
  public Workbook getWorkbook() {
    return reader.getWorkbook();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Iterator<Row> iterator() {
    return reader.iterator();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Iterator<Row> rowIterator() {
    return reader.iterator();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public String getSheetName() {
    return name;
  }

  /**
   * Get the hidden state for a given column
   *
   * @param columnIndex - the column to set (0-based)
   * @return hidden - <code>false</code> if the column is visible
   */
  @Override
  public boolean isColumnHidden(int columnIndex) {
    return reader.isColumnHidden(columnIndex);
  }

  /**
   * Gets the first row on the sheet. This value is only available on some sheets where the
   * sheet XML has the dimension data set. At present, this method will return 0 if this
   * dimension data is missing (this may change in a future release).
   *
   * @return first row contained in this sheet (0-based)
   */
  @Override
  public int getFirstRowNum() {
    return reader.getFirstRowNum();
  }

  /**
   * Gets the last row on the sheet. This value is only available on some sheets where the
   * sheet XML has the dimension data set. At present, this method will return 0 if this
   * dimension data is missing (this may change in a future release).
   *
   * @return last row contained in this sheet (0-based)
   */
  @Override
  public int getLastRowNum() {
    return reader.getLastRowNum();
  }

  /**
   * Return cell comment at row, column, if one exists. Otherwise, return null.
   *
   * @param cellAddress the location of the cell comment
   * @return the cell comment, if one exists. Otherwise, return null.
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadComments(boolean)} is not set to true
   */
  @Override
  public Comment getCellComment(CellAddress cellAddress) {
    Comments sheetComments = reader.getCellComments();
    if (sheetComments == null) {
      return null;
    }
    XSSFComment xssfComment = sheetComments.findCellComment(cellAddress);
    if (xssfComment != null && reader.getBuilder().adjustLegacyComments()) {
      return new WrappedComment(xssfComment);
    }
    return xssfComment;
  }

  /**
   * Returns all cell comments on this sheet.
   * @return A map of each Comment in the sheet, keyed on the cell address where
   * the comment is located.
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadComments(boolean)} is not set to true
   */
  @Override
  public Map<CellAddress, ? extends Comment> getCellComments() {
    Comments sheetComments = reader.getCellComments();
    if (sheetComments == null) {
      return Collections.emptyMap();
    }
    Map<CellAddress, Comment> map = new HashMap<>();
    for(Iterator<CellAddress> iter = sheetComments.getCellAddresses(); iter.hasNext(); ) {
      CellAddress address = iter.next();
      map.put(address, getCellComment(address));
    }
    return map;
  }

  /**
   * Only works after sheet is fully read (because merged regions data is stored
   * at the end of the sheet XML).
   */
  @Override
  public CellRangeAddress getMergedRegion(int index) {
    List<CellRangeAddress> regions = getMergedRegions();
    if(index > regions.size()) {
      throw new NoSuchElementException("index " + index + " is out of range");
    }
    return regions.get(index);
  }

  /**
   * Only works after sheet is fully read (because merged regions data is stored
   * at the end of the sheet XML).
   */
  @Override
  public List<CellRangeAddress> getMergedRegions() {
    return reader.getMergedCells();
  }

  /**
   * Only works after sheet is fully read (because merged regions data is stored
   * at the end of the sheet XML).
   */
  @Override
  public int getNumMergedRegions() {
    List<CellRangeAddress> mergedCells = reader.getMergedCells();
    return mergedCells == null ? 0 : mergedCells.size();
  }

  /**
   * Return the sheet's existing drawing, or null if there isn't yet one.
   *
   * @return a SpreadsheetML drawing
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadShapes(boolean)} is not set to true
   */
  @Override
  public Drawing getDrawingPatriarch() {
    return reader.getDrawingPatriarch();
  }

  /**
   * Get a Hyperlink in this sheet anchored at row, column (only if feature is enabled on the Builder).
   *
   * @param row The row where the hyperlink is anchored
   * @param column The column where the hyperlink is anchored
   * @return hyperlink if there is a hyperlink anchored at row, column; otherwise returns null
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadHyperlinks(boolean)} is not set to true
   */
  @Override
  public Hyperlink getHyperlink(int row, int column) {
    return getHyperlink(new CellAddress(row, column));
  }

  /**
   * Get hyperlink associated with cell (only if feature is enabled on the Builder).
   * This should only be called after all the rows are read because the hyperlink data is
   * at the end of the sheet.
   *
   * @param cellAddress
   * @return the hyperlink associated with this cell (only if feature is enabled on the Builder) - null if not found
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadHyperlinks(boolean)} is not set to true
   */
  @Override
  public Hyperlink getHyperlink(CellAddress cellAddress) {
    for (Hyperlink hyperlink : getHyperlinkList()) {
      if (cellAddress.getRow() >= hyperlink.getFirstRow() && cellAddress.getRow() <= hyperlink.getLastRow()
        && cellAddress.getColumn() >= hyperlink.getFirstColumn() && cellAddress.getColumn() <= hyperlink.getLastColumn()) {
        return hyperlink;
      }
    }
    return null;
  }

  /**
   * Get hyperlinks associated with sheet (only if feature is enabled on the Builder).
   * This should only be called after all the rows are read because the hyperlink data is
   * at the end of the sheet.
   *
   * @return the hyperlinks associated with this sheet (only if feature is enabled on the Builder) - cast to {@link XlsxHyperlink} to access cell reference
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadHyperlinks(boolean)} is not set to true
   */
  @Override
  public List<? extends Hyperlink> getHyperlinkList() {
    return reader.getHyperlinks();
  }

  @Override
  public CellAddress getActiveCell() {
    return reader.getActiveCell();
  }

  /**
   * @return immutable copy of the shared formula map for this sheet
   */
  public Map<String, SharedFormula> getSharedFormulaMap() {
    return reader.getSharedFormulaMap();
  }

  /**
   * @param siValue the ID for the shared formula (appears in Excel sheet XML as an <code>si</code> attribute
   * @param sharedFormula maps the base cell and formula for the shared formula
   */
  public void addSharedFormula(String siValue, SharedFormula sharedFormula) {
    reader.addSharedFormula(siValue, sharedFormula);
  }

  /**
   * @param siValue the ID for the shared formula (appears in Excel sheet XML as an <code>si</code> attribute
   * @return the shared formula that was removed (can be null if no existing shared formula is found)
   */
  public SharedFormula removeSharedFormula(String siValue) {
    return reader.removeSharedFormula(siValue);
  }

  /* Unsupported */

  /**
   * Update operations are not supported
   */
  @Override
  public Row createRow(int rownum) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void removeRow(Row row) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported - use {@link #iterator()} or {@link #rowIterator()} instead
   */
  @Override
  public Row getRow(int rownum) {
    throw new UnsupportedOperationException("use iterator() or rowIterator() instead");
  }

  /**
   * Not supported - use {@link #iterator()} or {@link #rowIterator()} to iterate over rows
   * and count the rows
   */
  @Override
  public int getPhysicalNumberOfRows() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setColumnHidden(int columnIndex, boolean hidden) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setRightToLeft(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public boolean isRightToLeft() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setColumnWidth(int columnIndex, int width) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public int getColumnWidth(int columnIndex) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public float getColumnWidthInPixels(int columnIndex) {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setDefaultColumnWidth(int width) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public int getDefaultColumnWidth() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public short getDefaultRowHeight() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public float getDefaultRowHeightInPoints() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setDefaultRowHeight(short height) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setDefaultRowHeightInPoints(float height) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public CellStyle getColumnStyle(int column) {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public int addMergedRegion(CellRangeAddress region) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public int addMergedRegionUnsafe(CellRangeAddress cellRangeAddress) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void validateMergedRegions() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setVerticallyCenter(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setHorizontallyCenter(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public boolean getHorizontallyCenter() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean getVerticallyCenter() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void removeMergedRegion(int index) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void removeMergedRegions(Collection<Integer> collection) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setForceFormulaRecalculation(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public boolean getForceFormulaRecalculation() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setAutobreaks(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setDisplayGuts(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setDisplayZeros(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public boolean isDisplayZeros() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setFitToPage(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setRowSumsBelow(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setRowSumsRight(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public boolean getAutobreaks() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean getDisplayGuts() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean getFitToPage() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean getRowSumsBelow() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean getRowSumsRight() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean isPrintGridlines() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setPrintGridlines(boolean show) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public boolean isPrintRowAndColumnHeadings() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setPrintRowAndColumnHeadings(boolean b) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public PrintSetup getPrintSetup() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public Header getHeader() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public Footer getFooter() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setSelected(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public double getMargin(short margin) {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setMargin(short margin, double size) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public boolean getProtect() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void protectSheet(String password) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean getScenarioProtect() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setZoom(int i) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public short getTopRow() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public short getLeftCol() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void showInPane(int toprow, int leftcol) {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void shiftRows(int startRow, int endRow, int n) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void shiftRows(int startRow, int endRow, int n, boolean copyRowHeight, boolean resetOriginalRowHeight) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void shiftColumns(int startColumn, int endColumn, final int n) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void createFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void createFreezePane(int colSplit, int rowSplit) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void createSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, int activePane) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public PaneInformation getPaneInformation() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setDisplayGridlines(boolean show) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public boolean isDisplayGridlines() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setDisplayFormulas(boolean show) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public boolean isDisplayFormulas() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setDisplayRowColHeadings(boolean show) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public boolean isDisplayRowColHeadings() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setRowBreak(int row) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public boolean isRowBroken(int row) {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void removeRowBreak(int row) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public int[] getRowBreaks() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public int[] getColumnBreaks() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setColumnBreak(int column) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public boolean isColumnBroken(int column) {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void removeColumnBreak(int column) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setColumnGroupCollapsed(int columnNumber, boolean collapsed) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void groupColumn(int fromColumn, int toColumn) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void ungroupColumn(int fromColumn, int toColumn) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void groupRow(int fromRow, int toRow) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void ungroupRow(int fromRow, int toRow) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setRowGroupCollapsed(int row, boolean collapse) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setDefaultColumnStyle(int column, CellStyle style) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void autoSizeColumn(int column) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void autoSizeColumn(int column, boolean useMergedCells) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public Drawing createDrawingPatriarch() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean isSelected() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public CellRange<? extends Cell> setArrayFormula(String formula, CellRangeAddress range) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public CellRange<? extends Cell> removeArrayFormula(Cell cell) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public DataValidationHelper getDataValidationHelper() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public List<? extends DataValidation> getDataValidations() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void addValidationData(DataValidation dataValidation) {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public AutoFilter setAutoFilter(CellRangeAddress range) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public SheetConditionalFormatting getSheetConditionalFormatting() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public CellRangeAddress getRepeatingRows() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public CellRangeAddress getRepeatingColumns() {
    throw new UnsupportedOperationException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setRepeatingRows(CellRangeAddress rowRangeRef) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setRepeatingColumns(CellRangeAddress columnRangeRef) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public int getColumnOutlineLevel(int columnIndex) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setActiveCell(CellAddress cellAddress) {
    throw new UnsupportedOperationException("update operations are not supported");
  }
}
