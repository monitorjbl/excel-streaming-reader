package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.XmlUtils;
import com.github.pjfanning.xlsx.exceptions.NotSupportedException;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

public class StreamingCell implements Cell {

  private static final Supplier NULL_SUPPLIER = () -> null;

  private final Sheet sheet;
  private final int columnIndex;
  private int rowIndex;
  private Row row;
  private final boolean use1904Dates;

  private Supplier contentsSupplier = NULL_SUPPLIER;
  private String rawContents;
  private String formula;
  private String numericFormat;
  private Short numericFormatIndex;
  private String type;
  private CellStyle cellStyle;
  private boolean formulaType;

  public StreamingCell(Sheet sheet, int columnIndex, int rowIndex, boolean use1904Dates) {
    this.sheet = sheet;
    this.columnIndex = columnIndex;
    this.rowIndex = rowIndex;
    this.use1904Dates = use1904Dates;
  }

  public StreamingCell(Sheet sheet, int columnIndex, Row row, boolean use1904Dates) {
    this.sheet = sheet;
    this.columnIndex = columnIndex;
    this.row = row;
    this.use1904Dates = use1904Dates;
  }

  public void setContentSupplier(Supplier contentsSupplier) {
    this.contentsSupplier = contentsSupplier;
  }

  public void setRawContents(String rawContents) {
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
    this.type = type;
  }

  public boolean isFormulaType() {
    return formulaType;
  }

  public void setFormulaType(boolean formulaType) {
    this.formulaType = formulaType;
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
    return (row == null) ? rowIndex : row.getRowNum();
  }

  /**
   * Row is not guaranteed to be set. Will return null when row is not set.
   */
  @Override
  public Row getRow() {
    return row;
  }

  @Override
  public Sheet getSheet() {
    return sheet;
  }

  /**
   * Return the cell type.
   *
   * @return the cell type
   */
  @Override
  public CellType getCellType() {
    if(formulaType) {
      return CellType.FORMULA;
    } else if(contentsSupplier.getContent() == null || type == null) {
      return CellType.BLANK;
    } else if("n".equals(type)) {
      return CellType.NUMERIC;
    } else if("d".equals(type)) {
      return CellType.NUMERIC;
    } else if("s".equals(type) || "inlineStr".equals(type) || "str".equals(type)) {
      return CellType.STRING;
    } else if("b".equals(type)) {
      return CellType.BOOLEAN;
    } else if("e".equals(type)) {
      return CellType.ERROR;
    } else {
      throw new UnsupportedOperationException("Unsupported cell type '" + type + "'");
    }
  }

  /**
   * Get the value of the cell as a string.
   * For blank cells we return an empty string.
   *
   * @return the value of the cell as a string
   */
  @Override
  public String getStringCellValue() {
    Object c = contentsSupplier.getContent();

    return c == null ? "" : c.toString();
  }

  /**
   * Get the value of the cell as a number. For strings we throw an exception. For
   * blank cells we return a 0.
   *
   * @return the value of the cell as a number
   * @throws NumberFormatException if the cell value isn't a parseable <code>double</code>.
   */
  @Override
  public double getNumericCellValue() {
    if ("d".equals(type)) {
      try {
        LocalDateTime dt = DateTimeUtil.parseDateTime(rawContents);
        return DateUtil.getExcelDate(dt, use1904Dates);
      } catch (Exception e) {
        try {
          return DateTimeUtil.convertTime(rawContents);
        } catch (Exception e2) {
          throw new IllegalStateException("cannot parse strict format date/time " + rawContents);
        }
      }
    }
    return rawContents == null ? 0.0 : Double.parseDouble(rawContents);
  }

  /**
   * Get the value of the cell as a date. For strings we throw an exception. For
   * blank cells we return a null.
   *
   * @return the value of the cell as a date
   * @throws IllegalStateException if the cell type returned by {@link #getCellType()} is CELL_TYPE_STRING
   * @throws NumberFormatException if the cell value isn't a parseable <code>double</code>.
   */
  @Override
  public Date getDateCellValue() {
    if(getCellType() == CellType.STRING) {
      throw new IllegalStateException("Cell type cannot be CELL_TYPE_STRING");
    }
    return rawContents == null ? null : DateUtil.getJavaDate(getNumericCellValue(), use1904Dates);
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

    if (cellType == CellType.FORMULA) {
      cellType = getCachedFormulaResultType();
    }
    switch(cellType) {
      case BLANK:
        return false;
      case BOOLEAN:
        return rawContents != null && XmlUtils.evaluateBoolean(rawContents);
      default:
        throw typeMismatch(CellType.BOOLEAN, cellType, isFormulaType());
    }
  }

  private static IllegalStateException typeMismatch(CellType expectedType, CellType actualType, boolean isFormulaCell) {
    String msg = "Cannot get a "
            + getCellTypeName(expectedType) + " value from a "
            + getCellTypeName(actualType) + " " + (isFormulaCell ? "formula " : "") + "cell";
    return new IllegalStateException(msg);
  }

  /**
   * Used to help format error messages
   */
  private static String getCellTypeName(CellType cellType) {
    switch (cellType) {
      case BLANK:   return "blank";
      case STRING:  return "text";
      case BOOLEAN: return "boolean";
      case ERROR:   return "error";
      case NUMERIC: return "numeric";
      case FORMULA: return "formula";
    }
    return "#unknown cell type (" + cellType + ")#";
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
    if (!formulaType)
      throw new IllegalStateException("This cell does not have a formula");
    return formula;
  }

  /**
   * Get the value of the cell as a XSSFRichTextString
   * <p>
   * For numeric cells we throw an exception. For blank cells we return an empty string.
   * For formula cells we return the pre-calculated value if a string, otherwise an exception
   * </p>
   * @return the value of the cell as a XSSFRichTextString
   * @throws NotSupportedException if the cell type is unsupported
   */
  @Override
  public XSSFRichTextString getRichStringCellValue() {
    CellType cellType = getCellType();
    if (cellType == CellType.FORMULA) {
      cellType = getCachedFormulaResultType();
    }
    XSSFRichTextString rt;
    switch (cellType) {
      case BLANK:
        rt = new XSSFRichTextString("");
        break;
      case STRING:
        rt = new XSSFRichTextString(getStringCellValue());
        break;
      default:
        throw new NotSupportedException("getRichStringCellValue does not support cell type " + cellType.toString());
    }
    return rt;
  }

  /**
   * Only valid for formula cells
   * @return one of ({@link CellType#NUMERIC}, {@link CellType#STRING},
   *     {@link CellType#BOOLEAN}, {@link CellType#ERROR}) depending
   * on the cached value of the formula
   * @throws IllegalStateException if cell is not formula type
   * @throws NotSupportedException if cell formula type is unknown
   */
  @Override
  public CellType getCachedFormulaResultType() {
    if (formulaType) {
      if(contentsSupplier.getContent() == null || type == null) {
        return CellType.BLANK;
      } else if("n".equals(type)) {
        return CellType.NUMERIC;
      } else if("d".equals(type)) {
        return CellType.NUMERIC;
      } else if("s".equals(type) || "inlineStr".equals(type) || "str".equals(type)) {
        return CellType.STRING;
      } else if("b".equals(type)) {
        return CellType.BOOLEAN;
      } else if("e".equals(type)) {
        return CellType.ERROR;
      } else {
        throw new NotSupportedException("Unsupported cell type '" + type + "'");
      }
    } else  {
      throw new IllegalStateException("Only formula cells have cached results");
    }
  }

  @Override
  public LocalDateTime getLocalDateTimeCellValue() {
    if(getCellType() == CellType.STRING) {
      throw new IllegalStateException("Cell type cannot be CELL_TYPE_STRING");
    }
    return rawContents == null ? null : DateUtil.getLocalDateTime(getNumericCellValue(), use1904Dates);
  }

  @Override
  public byte getErrorCellValue() {
    CellType cellType = getCellType();
    if(cellType != CellType.ERROR) {
      throw typeMismatch(CellType.ERROR, cellType, false);
    }
    String code = rawContents;
    if (code == null) {
      return 0;
    }
    try {
      return FormulaError.forString(code).getCode();
    } catch (final IllegalArgumentException e) {
      throw new IllegalStateException("Unexpected error code", e);
    }
  }

  @Override
  public CellAddress getAddress() {
    return new CellAddress(this);
  }

  /**
   * Returns cell comment associated with this cell
   *
   * @return the cell comment associated with this cell or {@code null}
   * @throws IllegalStateException if StreamingWorkbook.Builder setReadComments is not set to true
   */
  @Override
  public Comment getCellComment() {
    return (sheet == null) ? null : sheet.getCellComment(getAddress());
  }

  /**
   * Returns hyperlink associated with this cell. This is not recommended as this data is stored at the end
   * of the sheet. Use the hyperlink methods on sheet instance instead or keep this cell instance in memory
   * until after all the rows have been read.
   *
   * @return the hyperlink associated with this cell or {@code null}
   * @throws IllegalStateException if StreamingWorkbook.Builder setReadHyperlinks is not set to true
   */
  @Override
  public Hyperlink getHyperlink() {
    return (sheet == null) ? null : sheet.getHyperlink(getAddress());
  }

  /* Not supported */

  /**
   * Update operations are not supported
   */
  @Override
  public void setCellType(CellType cellType) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setBlank() { throw new NotSupportedException("update operations are not supported"); }

  /**
   * Update operations are not supported
   */
  @Override
  public void setCellValue(double value) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setCellValue(Date value) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setCellValue(Calendar value) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setCellValue(LocalDate value) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setCellValue(LocalDateTime value) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setCellValue(RichTextString value) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setCellValue(String value) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setCellFormula(String formula) throws FormulaParseException {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void removeFormula() throws IllegalStateException {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setCellValue(boolean value) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setCellErrorValue(byte value) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setAsActiveCell() {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setCellComment(Comment comment) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void removeCellComment() {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void setHyperlink(Hyperlink link) {
    throw new NotSupportedException("update operations are not supported");
  }

  /**
   * Update operations are not supported
   */
  @Override
  public void removeHyperlink() {
    throw new NotSupportedException("update operations are not supported");
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
