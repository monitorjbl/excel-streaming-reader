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
  private int columnIndex;
  private int rowIndex;

  private Object contents;
  private String type;
  private Row row;

  /**
   * The cell type of the data value is indicated by an attribute on the
   * cell.  The value is usually in a "v" element within the cell.
   */
  public static enum XSSF_DATA_TYPE {
     BOOLEAN,
     ERROR,
     FORMULA,
     INLINE_STRING,
     NUMBER,
     SST_INDEX
  }

  /**
   * Used to format the actual Excel spreadsheet cell value, instead of the raw
   * one.
   */
  private XSSF_DATA_TYPE xssfDataType = null;

  /**
   * The style index, used to look up a built-in style in a
   * <CODE>StylesTable</CODE> to format numeric cell values.
   */
  private Short numericFormatIndex = null;

  /**
   * The format <CODE>String</CODE> used in conjunction with the
   * <CODE>numericFormatIndex</CODE> to format a numeric cell value.
   *
   * @see numericFormatIndex
   */
  private String numericFormatString = null;

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

  public String getType() {
    return type;
  }

  public void setType(String type) {
    this.type = type;
  }

  public void setRow(Row row) {
    this.row = row;
  }

  /**
   * Return the Excel XSSF spreadsheet cell type, used to determine the actual
   * cell value instead of the raw one.
   *
   * @return the Excel XSSF spreadsheet cell type, used to determine the actual
   *         cell value instead of the raw one.
   */
  public XSSF_DATA_TYPE getXssfDataType() {
     return xssfDataType;
  }

  /**
   * Set the Excel XSSF spreadsheet cell type, used to determine the actual
   * cell value instead of the raw one.
   *
   * @param xssfDataType the Excel XSSF spreadsheet cell type, used to determine
   *        the actual cell value instead of the raw one.
   */
  public void setXssfDataType(XSSF_DATA_TYPE xssfDataType) {
     this.xssfDataType = xssfDataType;
  }

  /**
   * Return the style index, used to look up a built-in style in a
   * <CODE>StylesTable</CODE> to format numeric cell values.
   *
   * @return the style index, used to look up a built-in style in a
   *         <CODE>StylesTable</CODE> to format numeric cell values.
   */
  public Short getNumericFormatIndex() {
     return numericFormatIndex;
  }

  /**
   * Set the style index, used to look up a built-in style in a
   * <CODE>StylesTable</CODE> to format numeric cell values.
   *
   * @param numericFormatIndex the style index, used to look up a built-in style
   *        in a <CODE>StylesTable</CODE> to format numeric cell values.
   */
  public void setNumericFormatIndex(Short numericFormatIndex) {
     this.numericFormatIndex = numericFormatIndex;
  }

  /**
   * Return the format <CODE>String</CODE> used in conjunction with the
   * <CODE>numericFormatIndex</CODE> to format a numeric cell value.
   *
   * @return the format <CODE>String</CODE> used in conjunction with the
   *         <CODE>numericFormatIndex</CODE> to format a numeric cell value.
   */
  public String getNumericFormatString() {
     return numericFormatString;
  }

  /**
   * Set the format <CODE>String</CODE> used in conjunction with the
   * <CODE>numericFormatIndex</CODE> to format a numeric cell value.
   *
   * @param numericFormatString the format <CODE>String</CODE> used in
   *        conjunction with the <CODE>numericFormatIndex</CODE> to format a
   *        numeric cell value.
   */
  public void setNumericFormatString(String numericFormatString) {
     this.numericFormatString = numericFormatString;
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
   * @throws UnsupportedOperationException Thrown if the type is not one supported by the streamer.
   *                                       It may be possible to still read the value as a supported type
   *                                       via {@code getStringCellValue()}, {@code getNumericCellValue},
   *                                       or {@code getDateCellValue()}
   * @see Cell#CELL_TYPE_BLANK
   * @see Cell#CELL_TYPE_NUMERIC
   * @see Cell#CELL_TYPE_STRING
   */
  @Override
  public int getCellType() {
    if (contents == null || type == null) {
      return Cell.CELL_TYPE_BLANK;
    } else if ("n".equals(type)) {
      return Cell.CELL_TYPE_NUMERIC;
    } else if ("s".equals(type)) {
      return Cell.CELL_TYPE_STRING;
    } else {
      throw new UnsupportedOperationException("Unsupported cell type '" + type + "'");
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
