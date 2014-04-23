package com.thundermoose.xlsx.impl;

import com.thundermoose.xlsx.exceptions.NotSupportedException;
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

  @Override
  public int getRowNum() {
    return rowIndex;
  }

  @Override
  public Iterator<Cell> cellIterator() {
    return cellMap.values().iterator();
  }

  @Override
  public Iterator<Cell> iterator() {
    return cellMap.values().iterator();
  }

  @Override
  public Cell getCell(int cellnum) {
    return cellMap.get(cellnum);
  }

  /* Not supported */

  @Override
  public Cell createCell(int column) {
    throw new NotSupportedException();
  }

  @Override
  public Cell createCell(int column, int type) {
    throw new NotSupportedException();
  }

  @Override
  public void removeCell(Cell cell) {
    throw new NotSupportedException();
  }

  @Override
  public void setRowNum(int rowNum) {
    throw new NotSupportedException();
  }

  @Override
  public Cell getCell(int cellnum, MissingCellPolicy policy) {
    throw new NotSupportedException();
  }

  @Override
  public short getFirstCellNum() {
    throw new NotSupportedException();
  }

  @Override
  public short getLastCellNum() {
    throw new NotSupportedException();
  }

  @Override
  public int getPhysicalNumberOfCells() {
    throw new NotSupportedException();
  }

  @Override
  public void setHeight(short height) {
    throw new NotSupportedException();
  }

  @Override
  public void setZeroHeight(boolean zHeight) {
    throw new NotSupportedException();
  }

  @Override
  public boolean getZeroHeight() {
    throw new NotSupportedException();
  }

  @Override
  public void setHeightInPoints(float height) {
    throw new NotSupportedException();
  }

  @Override
  public short getHeight() {
    throw new NotSupportedException();
  }

  @Override
  public float getHeightInPoints() {
    throw new NotSupportedException();
  }

  @Override
  public boolean isFormatted() {
    throw new NotSupportedException();
  }

  @Override
  public CellStyle getRowStyle() {
    throw new NotSupportedException();
  }

  @Override
  public void setRowStyle(CellStyle style) {
    throw new NotSupportedException();
  }

  @Override
  public Sheet getSheet() {
    throw new NotSupportedException();
  }

}