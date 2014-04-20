package com.thundermoose.xlsx.impl;

import com.thundermoose.xlsx.exceptions.NotSupportedException;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Calendar;
import java.util.Date;

public class StreamingCell implements Cell {
  public static final String NUMERIC_REGEX = "-?\\d+(\\.\\d+)?";

  private int columnIndex;
  private int rowIndex;

  private Object contents;
  private Row row;

  public StreamingCell(int columnIndex, int rowIndex) {
    this.columnIndex = columnIndex;
    this.rowIndex = rowIndex;
  }

  public Object getContents() {
    return contents;
  }

  public void setContents(Object contents) {
    this.contents = contents;
  }

  public void setRow(Row row) {
    this.row = row;
  }

  static boolean isNumeric(String str) {
    return str.matches(NUMERIC_REGEX);
  }

  /* Supported */

  @Override
  public int getColumnIndex() {
    return columnIndex;
  }

  @Override
  public int getRowIndex() {
    return rowIndex;
  }

  @Override
  public Row getRow() {
    return row;
  }

  @Override
  public int getCellType() {
    if (contents == null) {
      return Cell.CELL_TYPE_BLANK;
    } else if (isNumeric(contents.toString())) {
      return Cell.CELL_TYPE_NUMERIC;
    } else {
      return Cell.CELL_TYPE_STRING;
    }
  }

  @Override
  public String getStringCellValue() {
    return (String) contents;
  }

  @Override
  public double getNumericCellValue() {
    return Double.parseDouble((String) contents);
  }

  @Override
  public Date getDateCellValue() {
    return HSSFDateUtil.getJavaDate(getNumericCellValue());
  }

  /* Not supported */

  @Override
  public void setCellType(int cellType) {
    throw new NotSupportedException();
  }

  @Override
  public Sheet getSheet() {
    throw new NotSupportedException();
  }

  @Override
  public int getCachedFormulaResultType() {
    throw new NotSupportedException();
  }

  @Override
  public void setCellValue(double value) {
    throw new NotSupportedException();
  }

  @Override
  public void setCellValue(Date value) {
    throw new NotSupportedException();
  }

  @Override
  public void setCellValue(Calendar value) {
    throw new NotSupportedException();
  }

  @Override
  public void setCellValue(RichTextString value) {
    throw new NotSupportedException();
  }

  @Override
  public void setCellValue(String value) {
    throw new NotSupportedException();
  }

  @Override
  public void setCellFormula(String formula) throws FormulaParseException {
    throw new NotSupportedException();
  }

  @Override
  public String getCellFormula() {
    throw new NotSupportedException();
  }

  @Override
  public RichTextString getRichStringCellValue() {
    throw new NotSupportedException();
  }

  @Override
  public void setCellValue(boolean value) {
    throw new NotSupportedException();
  }

  @Override
  public void setCellErrorValue(byte value) {
    throw new NotSupportedException();
  }

  @Override
  public boolean getBooleanCellValue() {
    return false;
  }

  @Override
  public byte getErrorCellValue() {
    throw new NotSupportedException();
  }

  @Override
  public void setCellStyle(CellStyle style) {
    throw new NotSupportedException();
  }

  @Override
  public CellStyle getCellStyle() {
    throw new NotSupportedException();
  }

  @Override
  public void setAsActiveCell() {
    throw new NotSupportedException();
  }

  @Override
  public void setCellComment(Comment comment) {
    throw new NotSupportedException();
  }

  @Override
  public Comment getCellComment() {
    throw new NotSupportedException();
  }

  @Override
  public void removeCellComment() {
    throw new NotSupportedException();
  }

  @Override
  public Hyperlink getHyperlink() {
    throw new NotSupportedException();
  }

  @Override
  public void setHyperlink(Hyperlink link) {
    throw new NotSupportedException();
  }

  @Override
  public CellRangeAddress getArrayFormulaRange() {
    throw new NotSupportedException();
  }

  @Override
  public boolean isPartOfArrayFormulaGroup() {
    throw new NotSupportedException();
  }
}