package com.monitorjbl.xlsx.impl;

import java.util.Calendar;
import java.util.Date;

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
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;

import com.monitorjbl.xlsx.exceptions.NotSupportedException;

public class StreamingCell implements Cell {

  private static final Supplier NULL_SUPPLIER = new Supplier() {
    @Override
    public Object getContent() {
      return null;
    }
  };
  private static final Supplier NOT_SUPPORTED_SUPPLIER = new Supplier() {
    @Override
    public Object getContent() {
      throw new NotSupportedException();
    }
  };

  private static final String FALSE_AS_STRING = "0";
  private static final String TRUE_AS_STRING  = "1";

  private int columnIndex;
  private int rowIndex;
  private final boolean use1904Dates;

  private Supplier commentsTableSupplier = NOT_SUPPORTED_SUPPLIER;
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

  public void setCommentsTableSupplier(Supplier commentsTableSupplier) {
    this.commentsTableSupplier = commentsTableSupplier;
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
   * Will return {@link CellType} in version 4.0 of POI.
   * For forwards compatibility, do not hard-code cell type literals in your code.
   *
   * @return the cell type
   */
  @Override
  public int getCellType() {
    return getCellTypeEnum().getCode();
  }

  /**
   * Return the cell type.
   *
   * @return the cell type
   * Will be renamed to <code>getCellType()</code> when we make the CellType enum transition in POI 4.0. See bug 59791.
   */
  @Override
  public CellType getCellTypeEnum() {
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
    if(getCellType() == CELL_TYPE_STRING){
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
    int cellType = getCellType();
    switch(cellType) {
      case CELL_TYPE_BLANK:
        return false;
      case CELL_TYPE_BOOLEAN:
        return rawContents != null && TRUE_AS_STRING.equals(rawContents);
      case CELL_TYPE_FORMULA:
        throw new NotSupportedException();
      default:
        throw typeMismatch(CELL_TYPE_BOOLEAN, cellType, false);
    }
  }

  private static RuntimeException typeMismatch(int expectedTypeCode, int actualTypeCode, boolean isFormulaCell) {
    String msg = "Cannot get a "
            + getCellTypeName(expectedTypeCode) + " value from a "
            + getCellTypeName(actualTypeCode) + " " + (isFormulaCell ? "formula " : "") + "cell";
    return new IllegalStateException(msg);
  }

  /**
   * Used to help format error messages
   */
  private static String getCellTypeName(int cellTypeCode) {
    switch (cellTypeCode) {
      case CELL_TYPE_BLANK:   return "blank";
      case CELL_TYPE_STRING:  return "text";
      case CELL_TYPE_BOOLEAN: return "boolean";
      case CELL_TYPE_ERROR:   return "error";
      case CELL_TYPE_NUMERIC: return "numeric";
      case CELL_TYPE_FORMULA: return "formula";
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
  public int getCachedFormulaResultType() {
    return getCachedFormulaResultTypeEnum().getCode();
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

  /**
   * Retrieve cell comment if set and comment parsing is enabled otherwise null will be returned
   *
   * @return the comment set to the current cell or null if not set or comments reading is not enabled
   */
  @Override
  public Comment getCellComment() {
    CellAddress ref = new CellAddress(rowIndex, columnIndex);
    CommentsTable commentsTable = (CommentsTable) this.commentsTableSupplier.getContent();
    CTComment ctComment = commentsTable.getCTComment(ref);
    if (ctComment == null) {
      return null;
    } else {
      return new XSSFComment(commentsTable, ctComment, null);
    }
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public CellAddress getAddress() {
    return new CellAddress(rowIndex, columnIndex);
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
  public void setCellComment(Comment comment) {
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
