package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;

import java.util.*;

public class StreamingSheet implements Sheet {

  private final String name;
  private final StreamingSheetReader reader;
  private final StreamingWorkbook workbook;

  public StreamingSheet(StreamingWorkbook workbook, String name, StreamingSheetReader reader) {
    this.workbook = workbook;
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
    return workbook;
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
   * Gets the first row on the sheet
   *
   * @return first row contained in this sheet (0-based)
   */
  @Override
  public int getFirstRowNum() {
    return reader.getFirstRowNum();
  }

  /**
   * Gets the last row on the sheet
   *
   * @return last row contained in this sheet (0-based)
   */
  @Override
  public int getLastRowNum() {
    return reader.getLastRowNum();
  }

  @Override
  public Comment getCellComment(CellAddress cellAddress) {
    CommentsTable sheetComments = reader.getCellComments();
    if (sheetComments == null) {
      return null;
    }

    final int row = cellAddress.getRow();
    final int column = cellAddress.getColumn();

    CellAddress ref = new CellAddress(row, column);
    CTComment ctComment = sheetComments.getCTComment(ref);
    if(ctComment == null) {
      return null;
    }

    return new XSSFComment(sheetComments, ctComment, null);
  }

  @Override
  public Map<CellAddress, ? extends Comment> getCellComments() {
    CommentsTable sheetComments = reader.getCellComments();
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
   * Not supported
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
   * Not supported
   */
  @Override
  public void setVerticallyCenter(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
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
   * Not supported
   */
  @Override
  public void removeMergedRegion(int index) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void removeMergedRegions(Collection<Integer> collection) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public int getNumMergedRegions() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public CellRangeAddress getMergedRegion(int index) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public List<CellRangeAddress> getMergedRegions() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
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
   * Not supported
   */
  @Override
  public void setAutobreaks(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void setDisplayGuts(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
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
   * Not supported
   */
  @Override
  public void setFitToPage(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void setRowSumsBelow(boolean value) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
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
   * Not supported
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
   * Not supported
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
   * Not supported
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
   * Not supported
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
   * Not supported
   */
  @Override
  public void shiftRows(int startRow, int endRow, int n) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void shiftRows(int startRow, int endRow, int n, boolean copyRowHeight, boolean resetOriginalRowHeight) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
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
   * Not supported
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
   * Not supported
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
   * Not supported
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
   * Not supported
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
   * Not supported
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
   * Not supported
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
   * Not supported
   */
  @Override
  public void removeColumnBreak(int column) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void setColumnGroupCollapsed(int columnNumber, boolean collapsed) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void groupColumn(int fromColumn, int toColumn) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void ungroupColumn(int fromColumn, int toColumn) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void groupRow(int fromRow, int toRow) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void ungroupRow(int fromRow, int toRow) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void setRowGroupCollapsed(int row, boolean collapse) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void setDefaultColumnStyle(int column, CellStyle style) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void autoSizeColumn(int column) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public void autoSizeColumn(int column, boolean useMergedCells) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public Drawing getDrawingPatriarch() {
    throw new UnsupportedOperationException();
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
   * Not supported
   */
  @Override
  public CellRange<? extends Cell> setArrayFormula(String formula, CellRangeAddress range) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
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
   * Not supported
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
   * Not supported
   */
  @Override
  public void setRepeatingRows(CellRangeAddress rowRangeRef) {
    throw new UnsupportedOperationException("update operations are not supported");
  }

  /**
   * Not supported
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
  public Hyperlink getHyperlink(int i, int i1) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public Hyperlink getHyperlink(CellAddress cellAddress) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public List<? extends Hyperlink> getHyperlinkList() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public CellAddress getActiveCell() {
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
