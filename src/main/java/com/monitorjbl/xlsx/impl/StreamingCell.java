package com.monitorjbl.xlsx.impl;

import com.monitorjbl.xlsx.exceptions.NotSupportedException;
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

  /**
   * Returns column index of this cell
   *
   * @return zero-based column index of a column in a sheet.
   */
  @Override
  public int getColumnIndex() {
    return columnIndex;
  }

  /**
   * Returns row index of a row in the sheet that contains this cell
   *
   * @return zero-based row index of a row in the sheet that contains this cell
   */
  @Override
  public int getRowIndex() {
    return rowIndex;
  }

  /**
   * Returns the Row this cell belongs to. Note that keeping references to cell
   * rows around after the iterator window has passed <b>will</b> preserve them.
   *
   * @return the Row that owns this cell
   */
  @Override
  public Row getRow() {
    return row;
  }

  /**
   * Return the cell type. Note that only the numeric, string, and blank types are
   * currently supported.
   *
   * @return the cell type
   * @see Cell#CELL_TYPE_BLANK
   * @see Cell#CELL_TYPE_NUMERIC
   * @see Cell#CELL_TYPE_STRING
   */
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

  /**
   * Get the value of the cell as a string. For numeric cells we throw an exception.
   * For blank cells we return an empty string.
   *
   * @return the value of the cell as a string
   */
  @Override
  public String getStringCellValue() {
    return (String) contents;
  }

  /**
   * Get the value of the cell as a number. For strings we throw an exception. For
   * blank cells we return a 0.
   *
   * @return the value of the cell as a number
   * @throws NumberFormatException if the cell value isn't a parsable <code>double</code>.
   */
  @Override
  public double getNumericCellValue() {
    return Double.parseDouble((String) contents);
  }

  /**
   * Get the value of the cell as a date. For strings we throw an exception. For
   * blank cells we return a null.
   *
   * @return the value of the cell as a date
   * @throws IllegalStateException if the cell type returned by {@link #getCellType()} is CELL_TYPE_STRING
   * @throws NumberFormatException if the cell value isn't a parsable <code>double</code>.
   */
  @Override
  public Date getDateCellValue() {
    return HSSFDateUtil.getJavaDate(getNumericCellValue());
  }

  /* Not supported */

  /**
   * Not supported
   */
  @Override
  public void setCellType(int cellType) {
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
  public int getCachedFormulaResultType() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setCellValue(double value) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setCellValue(Date value) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setCellValue(Calendar value) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setCellValue(RichTextString value) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setCellValue(String value) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setCellFormula(String formula) throws FormulaParseException {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public String getCellFormula() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public RichTextString getRichStringCellValue() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setCellValue(boolean value) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setCellErrorValue(byte value) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean getBooleanCellValue() {
    return false;
  }

  /**
   * Not supported
   */
  @Override
  public byte getErrorCellValue() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setCellStyle(CellStyle style) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public CellStyle getCellStyle() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setAsActiveCell() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setCellComment(Comment comment) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public Comment getCellComment() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void removeCellComment() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public Hyperlink getHyperlink() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public void setHyperlink(Hyperlink link) {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public CellRangeAddress getArrayFormulaRange() {
    throw new NotSupportedException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean isPartOfArrayFormulaGroup() {
    throw new NotSupportedException();
  }
}