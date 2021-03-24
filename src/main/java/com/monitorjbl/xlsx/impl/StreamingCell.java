package com.monitorjbl.xlsx.impl;

import com.monitorjbl.xlsx.exceptions.NotSupportedException;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.util.Calendar;
import java.util.Date;

public class StreamingCell implements Cell {

    private static final Supplier NULL_SUPPLIER = () -> null;
    private static final String FALSE_AS_STRING = "0";
    private static final String TRUE_AS_STRING = "1";

    private final Sheet sheet;
    private int columnIndex;
    private int rowIndex;
    private final boolean use1904Dates;

    private Supplier contentsSupplier = NULL_SUPPLIER;
    private Object rawContents;
    private String formula;
    private String numericFormat;
    private Short numericFormatIndex;
    private String type;
    private CellStyle cellStyle;
    private Row row;
    private boolean formulaType;

    public StreamingCell(Sheet sheet, int columnIndex, int rowIndex, boolean use1904Dates) {
        this.sheet = sheet;
        this.columnIndex = columnIndex;
        this.rowIndex = rowIndex;
        this.use1904Dates = use1904Dates;
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
     * Sets the Row this cell belongs to. Note that keeping references to cell
     * rows around after the iterator window has passed <b>will</b> preserve them.
     * <p>
     * The row is not automatically set.
     *
     * @param row The row
     */
    public void setRow(Row row) {
        this.row = row;
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
        } else if("s".equals(type) || "inlineStr".equals(type) || "str".equals(type)) {
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
        if(getCellType() == CellType.STRING) {
            throw new IllegalStateException("Cell type cannot be CELL_TYPE_STRING");
        }
        return rawContents == null ? null : DateUtil.getJavaDate(getNumericCellValue(), use1904Dates);
    }

    @Override
    public LocalDateTime getLocalDateTimeCellValue() {
        return LocalDateTime.ofInstant(Instant.ofEpochMilli(getDateCellValue().getTime()), ZoneOffset.systemDefault());
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

    /**
     * Get the value of the cell as a XSSFRichTextString
     * <p>
     * For numeric cells we throw an exception. For blank cells we return an empty string.
     * For formula cells we return the pre-calculated value if a string, otherwise an exception
     * </p>
     *
     * @return the value of the cell as a XSSFRichTextString
     */
    @Override
    public XSSFRichTextString getRichStringCellValue() {
        CellType cellType = getCellType();
        XSSFRichTextString rt;
        switch(cellType) {
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

    @Override
    public Sheet getSheet() {
        return sheet;
    }

    private static RuntimeException typeMismatch(CellType expectedType, CellType actualType, boolean isFormulaCell) {
        String msg = "Cannot get a "
            + getCellTypeName(expectedType) + " value from a "
            + getCellTypeName(actualType) + " " + (isFormulaCell ? "formula " : "") + "cell";
        return new IllegalStateException(msg);
    }

    /**
     * Used to help format error messages
     */
    private static String getCellTypeName(CellType cellType) {
        switch(cellType) {
            case BLANK:
                return "blank";
            case STRING:
                return "text";
            case BOOLEAN:
                return "boolean";
            case ERROR:
                return "error";
            case NUMERIC:
                return "numeric";
            case FORMULA:
                return "formula";
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
        if(!formulaType)
            throw new IllegalStateException("This cell does not have a formula");
        return formula;
    }

    /**
     * Only valid for formula cells
     *
     * @return one of ({@link CellType#NUMERIC}, {@link CellType#STRING},
     * {@link CellType#BOOLEAN}, {@link CellType#ERROR}) depending
     * on the cached value of the formula
     */
    @Override
    public CellType getCachedFormulaResultType() {
        if(formulaType) {
            if(contentsSupplier.getContent() == null || type == null) {
                return CellType.BLANK;
            } else if("n".equals(type)) {
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
        } else {
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
    public void setCellValue(LocalDateTime value) {
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

    /**
     * Not supported
     */
    @Override
    public void setBlank() {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void removeFormula() throws IllegalStateException {
        throw new NotSupportedException();
    }
}