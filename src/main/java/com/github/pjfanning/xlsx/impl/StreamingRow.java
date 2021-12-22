package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.exceptions.NotSupportedException;
import org.apache.poi.ss.usermodel.*;

import java.util.*;

public class StreamingRow implements Row {
  private final Sheet sheet;
  private final int rowIndex;
  private boolean isHidden;
  private TreeMap<Integer, Cell> cellMap = new TreeMap<>();
  private StreamingSheetReader streamingSheetReader;

  public StreamingRow(Sheet sheet, int rowIndex, boolean isHidden) {
    this.sheet = sheet;
    this.rowIndex = rowIndex;
    this.isHidden = isHidden;
  }

  void setStreamingSheetReader(StreamingSheetReader streamingSheetReader) {
    this.streamingSheetReader = streamingSheetReader;
  }

  public Map<Integer, Cell> getCellMap() {
    return cellMap;
  }

 /* Supported */

  /**
   * Get row number this row represents
   *
   * @return the row number (0 based)
   */
  @Override
  public int getRowNum() {
    return rowIndex;
  }

  /**
   * @return Cell iterator of the physically defined cells for this row.
   */
  @Override
  public Iterator<Cell> cellIterator() {
    return cellMap.values().iterator();
  }

  /**
   * @return Cell iterator of the physically defined cells for this row.
   */
  @Override
  public Iterator<Cell> iterator() {
    return cellMap.values().iterator();
  }

  @Override
  public Spliterator<Cell> spliterator() {
    return Spliterators.spliterator(cellMap.values(), Spliterator.ORDERED);
  }

  @Override
  public Sheet getSheet() {
    return sheet;
  }

  /**
   * Get the cell representing a given column (logical cell) 0-based.  If you
   * ask for a cell that is not defined, you get a null.
   *
   * @param cellnum 0 based column number
   * @return Cell representing that column or null if undefined.
   */
  @Override
  public Cell getCell(int cellnum) {
    return cellMap.get(cellnum);
  }

  /**
   * Gets the index of the last cell contained in this row <b>PLUS ONE</b>.
   *
   * @return short representing the last logical cell in the row <b>PLUS ONE</b>,
   * or -1 if the row does not contain any cells.
   */
  @Override
  public short getLastCellNum() {
    return (short) (cellMap.size() == 0 ? -1 : cellMap.lastEntry().getValue().getColumnIndex() + 1);
  }

  /**
   * Get whether or not to display this row with 0 height
   *
   * @return - zHeight height is zero or not.
   */
  @Override
  public boolean getZeroHeight() {
    return isHidden;
  }

  /**
   * Gets the number of defined cells (NOT number of cells in the actual row!).
   * That is to say if only columns 0,4,5 have values then there would be 3.
   *
   * @return int representing the number of defined cells in the row.
   */
  @Override
  public int getPhysicalNumberOfCells() {
    return cellMap.size();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public short getFirstCellNum() {
    if(cellMap.size() == 0) {
      return -1;
    }
    return cellMap.firstKey().shortValue();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Cell getCell(int cellnum, MissingCellPolicy policy) {
    StreamingCell cell = (StreamingCell) cellMap.get(cellnum);
    if(policy == MissingCellPolicy.CREATE_NULL_AS_BLANK) {
      if(cell == null) {
        boolean use1904Dates = streamingSheetReader == null ? false : streamingSheetReader.isUse1904Dates();
        return new StreamingCell(sheet, cellnum, this, use1904Dates);
      }
    } else if(policy == MissingCellPolicy.RETURN_BLANK_AS_NULL) {
      if(cell == null || cell.getCellType() == CellType.BLANK) { return null; }
    }
    return cell;
  }

  /* Not supported */

  /**
   * Not supported
   */
  @Override
  public Cell createCell(int column) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public Cell createCell(int i, CellType cellType) {
    throw new NotSupportedException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void removeCell(Cell cell) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setRowNum(int rowNum) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setHeight(short height) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setZeroHeight(boolean zHeight) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setHeightInPoints(float height) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public short getHeight() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public float getHeightInPoints() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean isFormatted() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public CellStyle getRowStyle() {
    throw new NotSupportedException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setRowStyle(CellStyle style) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Not supported
   */
  @Override
  public int getOutlineLevel() {
    throw new NotSupportedException();
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void shiftCellsRight(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void shiftCellsLeft(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {
    throw new NotSupportedException("update operations are not supported");
  }

}
