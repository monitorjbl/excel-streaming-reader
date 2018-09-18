package com.monitorjbl.xlsx.impl;

import com.monitorjbl.xlsx.exceptions.NotSupportedException;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.util.Calendar;
import java.util.Date;

public class StreamingCell implements Cell {

  private static final Supplier NULL_SUPPLIER = new Supplier() {
    @Override
    public Object getContent() {
      return null;
    }
  };

  private static final String FALSE_AS_STRING = "0";
  private static final String TRUE_AS_STRING  = "1";

  private int columnIndex;
  private int rowIndex;
  private final boolean use1904Dates;

  private Supplier contentsSupplier = NULL_SUPPLIER;
  private Object rawContents;
  private String formula;
  private String numericFormat;
  private Short numericFormatIndex;
  private String type;
  private String cachedFormulaResultType;
  private Row row;
  private CellStyle cellStyle;

  public StreamingCell(int columnIndex, int rowIndex, boolean use1904Dates) {
    this.columnIndex = columnIndex;
    this.rowIndex = rowIndex;
    this.use1904Dates = use1904Dates;
  }

  String getRawCachedFormulaResultType() {
    return cachedFormulaResultType;
  }

  boolean supportsSupplierOverride() {
    return "n".equals(cachedFormulaResultType);
  }

  public void setContentSupplier(Supplier contentsSupplier) {
    this.contentsSupplier = contentsSupplier;
  }

  public void setRawContents(Object rawContents) {
    this.rawContents = rawContents;
  }

  public String getNumericFormat() {
    return numericFormat;
  }

  public void setNumericFormat(String numericFormat) {
    this.numericFormat = numericFormat;
  }

  public Short getNumericFormatIndex() {
    return numericFormatIndex;
  }

  public void setNumericFormatIndex(Short numericFormatIndex) {
    this.numericFormatIndex = numericFormatIndex;
  }

  public void setFormula(String formula) {
    this.formula = formula;
  }

  public String getType() {
    return type;
  }

  public void setType(String type) {
    if("str".equals(type)) {
      // this is a formula cell, cache the value's type
      cachedFormulaResultType = this.type;
    }
    this.type = type;
  }

  public void setRow(Row row) {
    this.row = row;
  }

  @Override
  public void setCellStyle(CellStyle cellStyle) {
    this.cellStyle = cellStyle;
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
   * Return the cell type.
   *
   * @return the cell type
   */
  @Override
  public CellType getCellTypeEnum() {
    return getCellType();
  }

  /**
   * Return the cell type.
   *
   * @return the cell type
   */
  @Override
  public CellType getCellType() {
    if(contentsSupplier.getContent() == null || type == null) {
      return CellType.BLANK;
    } else if("n".equals(type)) {
      return CellType.NUMERIC;
    } else if("s".equals(type) || "inlineStr".equals(type)) {
      return CellType.STRING;
    } else if("str".equals(type)) {
      return CellType.FORMULA;
    } else if("b".equals(type)) {
      return CellType.BOOLEAN;
    } else if("e".equals(type)) {
      return CellType.ERROR;
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
    Object c = contentsSupplier.getContent();

    return c == null ? "" : (String) c;
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
    return rawContents == null ? 0.0 : Double.parseDouble((String) rawContents);
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
    if(getCellType() == CellType.STRING){
      throw new IllegalStateException("Cell type cannot be CELL_TYPE_STRING");
    }
    return rawContents == null ? null : HSSFDateUtil.getJavaDate(getNumericCellValue(), use1904Dates);
  }

  /**
   * Get the value of the cell as a boolean. For strings we throw an exception. For
   * blank cells we return a false.
   *
   * @return the value of the cell as a date
   */
  @Override
  public boolean getBooleanCellValue() {
    CellType cellType = getCellType();
    switch(cellType) {
      case BLANK:
        return false;
      case BOOLEAN:
        return rawContents != null && TRUE_AS_STRING.equals(rawContents);
      case FORMULA:
        throw new NotSupportedException();
      default:
        throw typeMismatch(CellType.BOOLEAN, cellType, false);
    }
  }

  private static RuntimeException typeMismatch(CellType expectedTypeCode, CellType actualTypeCode, boolean isFormulaCell) {
    String msg = "Cannot get a "
            + getCellTypeName(expectedTypeCode) + " value from a "
            + getCellTypeName(actualTypeCode) + " " + (isFormulaCell ? "formula " : "") + "cell";
    return new IllegalStateException(msg);
  }

  /**
   * Used to help format error messages
   */
  private static String getCellTypeName(CellType cellTypeCode) {
    switch (cellTypeCode) {
      case BLANK:   return "blank";
      case STRING:  return "text";
      case BOOLEAN: return "boolean";
      case ERROR:   return "error";
      case NUMERIC: return "numeric";
      case FORMULA: return "formula";
    }
    return "#unknown cell type (" + cellTypeCode + ")#";
  }

  /**
   * @return the style of the cell
   */
  @Override
  public CellStyle getCellStyle() {
    return this.cellStyle;
  }

  /**
   * Return a formula for the cell, for example, <code>SUM(C4:E4)</code>
   *
   * @return a formula for the cell
   * @throws IllegalStateException if the cell type returned by {@link #getCellType()} is not CELL_TYPE_FORMULA
   */
  @Override
  public String getCellFormula() {
    if (type == null || !"str".equals(type))
      throw new IllegalStateException("This cell does not have a formula");
    return formula;
  }

  /**
   * Only valid for formula cells
   * @return one of ({@link #CELL_TYPE_NUMERIC}, {@link #CELL_TYPE_STRING},
   *     {@link #CELL_TYPE_BOOLEAN}, {@link #CELL_TYPE_ERROR}) depending
   * on the cached value of the formula
   */
  @Override
  public CellType getCachedFormulaResultType() {
    return getCachedFormulaResultTypeEnum();
  }

  /**
   * Not supported
   */
  @Override
  public CellType getCachedFormulaResultTypeEnum() {
    if (type != null && "str".equals(type)) {
      if(contentsSupplier.getContent() == null || cachedFormulaResultType == null) {
        return CellType.BLANK;
      } else if("n".equals(cachedFormulaResultType)) {
        return CellType.NUMERIC;
      } else if("s".equals(cachedFormulaResultType) || "inlineStr".equals(cachedFormulaResultType)) {
        return CellType.STRING;
      } else if("str".equals(cachedFormulaResultType)) {
        return CellType.FORMULA;
      } else if("b".equals(cachedFormulaResultType)) {
        return CellType.BOOLEAN;
      } else if("e".equals(cachedFormulaResultType)) {
        return CellType.ERROR;
      } else {
        throw new UnsupportedOperationException("Unsupported cell type '" + cachedFormulaResultType + "'");
      }
    }
    else  {
      throw new IllegalStateException("Only formula cells have cached results");
    }
  }

  /* Not supported */

  /**
   * Not supported
   */
  @Override
  public void setCellType(CellType cellType) {
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
   * Get the value of the cell as a XSSFRichTextString
   * <p>
   * For numeric cells we throw an exception. For blank cells we return an empty string.
   * For formula cells we return the pre-calculated value if a string, otherwise an exception
   * </p>
   * @return the value of the cell as a XSSFRichTextString
   */
  @Override
  public XSSFRichTextString getRichStringCellValue() {
    CellType cellType = getCellTypeEnum();
    XSSFRichTextString rt;
    switch (cellType) {
      case BLANK:
        rt = new XSSFRichTextString("");
        break;
      case STRING:
        rt = new XSSFRichTextString(getStringCellValue());
        break;
      default:
        throw new NotSupportedException();
    }
    return rt;
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
  public byte getErrorCellValue() {
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
  public CellAddress getAddress() {
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
  public void removeHyperlink() {
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
