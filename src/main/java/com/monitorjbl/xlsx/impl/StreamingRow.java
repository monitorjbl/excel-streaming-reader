package com.monitorjbl.xlsx.impl;

import com.monitorjbl.xlsx.exceptions.NotSupportedException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;

public class StreamingRow implements Row {
  private int rowIndex;
  private Map<Integer, Cell> cellMap = new TreeMap<>();

  public StreamingRow(int rowIndex) {
    this.rowIndex = rowIndex;
  }

  public Map<Integer, Cell> getCellMap() {
    return cellMap;
  }

  public void setCellMap(Map<Integer, Cell> cellMap) {
    this.cellMap = cellMap;
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
  public Cell createCell(int column, int type) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void removeCell(Cell cell) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setRowNum(int rowNum) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public Cell getCell(int cellnum, MissingCellPolicy policy) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public short getFirstCellNum() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public short getLastCellNum() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public int getPhysicalNumberOfCells() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setHeight(short height) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setZeroHeight(boolean zHeight) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean getZeroHeight() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setHeightInPoints(float height) {
    throw new NotSupportedException();
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
   * Not supported
   */
  @Override
  public void setRowStyle(CellStyle style) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public Sheet getSheet() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public int getOutlineLevel() {
    throw new NotSupportedException();
  }

}