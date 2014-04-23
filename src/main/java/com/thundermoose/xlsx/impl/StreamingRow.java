package com.thundermoose.xlsx.impl;

import com.thundermoose.xlsx.exceptions.NotSupportedException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

public class StreamingRow implements Row {
  private int rowIndex;
  private List<Cell> cellList = new LinkedList<Cell>();

  public StreamingRow(int rowIndex) {
    this.rowIndex = rowIndex;
  }

  public List<Cell> getCellList() {
    return cellList;
  }

  public void setCellList(List<Cell> cellList) {
    this.cellList = cellList;
  }

  /* Supported */

  @Override
  public int getRowNum() {
    return rowIndex;
  }

  @Override
  public Iterator<Cell> cellIterator() {
    return cellList.iterator();
  }

  @Override
  public Iterator<Cell> iterator() {
    return cellList.iterator();
  }

  @Override
  public Cell getCell(int cellnum) {
    return cellList.size() > cellnum ? cellList.get(cellnum) : null;
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