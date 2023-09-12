package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.*;
import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.ptg.*;
import org.apache.poi.ss.formula.udf.AggregatingUDFFinder;
import org.apache.poi.ss.formula.udf.IndexedUDFFinder;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Internal;
import org.apache.poi.util.NotImplemented;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFTable;

import static com.github.pjfanning.xlsx.impl.NumberUtil.parseInt;

/**
 * Copied from POI BaseXSSFEvaluationWorkbook but a lot of stuff is removed because it is not easy
 * or impossible to support in excel-streaming-reader
 */
@Internal
abstract class BaseEvaluationWorkbook implements FormulaRenderingWorkbook, EvaluationWorkbook, FormulaParsingWorkbook {
  /**
   * The locator of user-defined functions.
   * By default, includes functions from the Excel Analysis Toolpack
   */
  private final IndexedUDFFinder _udfFinder = new IndexedUDFFinder(AggregatingUDFFinder.DEFAULT);

  protected final Workbook _uBook;

  protected BaseEvaluationWorkbook(Workbook book) {
    _uBook = book;
  }

  /* (non-JavaDoc), inherit JavaDoc from EvaluationWorkbook
   * @since POI 3.15 beta 3
   */
  @Override
  public void clearAllCachedResultValues() {

  }

  private int convertFromExternalSheetIndex(int externSheetIndex) {
    return externSheetIndex;
  }

  /**
   * XSSF doesn't use external sheet indexes, so when asked treat
   * it just as a local index
   */
  @Override
  public int convertFromExternSheetIndex(int externSheetIndex) {
    return externSheetIndex;
  }

  /**
   * @return  the external sheet index of the sheet with the given internal
   * index. Used by some of the more obscure formula and named range things.
   * Fairly easy on XSSF (we think...) since the internal and external
   * indices are the same
   */
  private int convertToExternalSheetIndex(int sheetIndex) {
    return sheetIndex;
  }

  @Override
  public int getExternalSheetIndex(String sheetName) {
    int sheetIndex = _uBook.getSheetIndex(sheetName);
    return convertToExternalSheetIndex(sheetIndex);
  }

  private int resolveBookIndex(String bookName) {
    // Strip the [] wrapper, if still present
    if (bookName.startsWith("[") && bookName.endsWith("]")) {
      bookName = bookName.substring(1, bookName.length()-2);
    }

    // Is it already in numeric form?
    try {
      return parseInt(bookName);
    } catch (NumberFormatException e) {}

    // Not properly referenced
    throw new RuntimeException("Book not linked for filename " + bookName);
  }

  @Override
  public EvaluationName getName(String name, int sheetIndex) {
    //EvaluationNames are not supported in excel-streaming-reader
    return null;
  }

  @Override
  public String getSheetName(int sheetIndex) {
    return _uBook.getSheetName(sheetIndex);
  }

  @Override
  @NotImplemented
  public ExternalName getExternalName(int externSheetIndex, int externNameIndex) {
    throw new IllegalStateException("ExternalNames are not supported in excel-streaming-reader");
  }

  @Override
  @NotImplemented
  public ExternalName getExternalName(String nameName, String sheetName, int externalWorkbookNumber) {
    throw new IllegalStateException("ExternalNames are not supported in excel-streaming-reader");
  }

  /**
   * Return an external name (named range, function, user-defined function) Pxg
   */
  @Override
  public NameXPxg getNameXPtg(String name, SheetIdentifier sheet) {
    // First, try to find it as a User Defined Function
    IndexedUDFFinder udfFinder = (IndexedUDFFinder)getUDFFinder();
    FreeRefFunction func = udfFinder.findFunction(name);
    if (func != null) {
      return new NameXPxg(null, name);
    }

    // Otherwise, try it as a named range
    if (sheet == null) {
      if (!_uBook.getNames(name).isEmpty()) {
        return new NameXPxg(null, name);
      }
      return null;
    }
    if (sheet.getSheetIdentifier() == null) {
      // Workbook + Named Range only
      int bookIndex = resolveBookIndex(sheet.getBookName());
      return new NameXPxg(bookIndex, null, name);
    }

    // Use the sheetname and process
    String sheetName = sheet.getSheetIdentifier().getName();

    if (sheet.getBookName() != null) {
      int bookIndex = resolveBookIndex(sheet.getBookName());
      return new NameXPxg(bookIndex, sheetName, name);
    } else {
      return new NameXPxg(sheetName, name);
    }
  }

  @Override
  public Ptg get3DReferencePtg(CellReference cell, SheetIdentifier sheet) {
    if (sheet.getBookName() != null) {
      int bookIndex = resolveBookIndex(sheet.getBookName());
      return new Ref3DPxg(bookIndex, sheet, cell);
    } else {
      return new Ref3DPxg(sheet, cell);
    }
  }

  @Override
  public Ptg get3DReferencePtg(AreaReference area, SheetIdentifier sheet) {
    if (sheet.getBookName() != null) {
      int bookIndex = resolveBookIndex(sheet.getBookName());
      return new Area3DPxg(bookIndex, sheet, area);
    } else {
      return new Area3DPxg(sheet, area);
    }
  }

  @Override
  @NotImplemented
  public String resolveNameXText(NameXPtg n) {
    throw new IllegalStateException("resolveNameXText is not supported in excel-streaming-reader");
  }

  @Override
  @NotImplemented
  public ExternalSheet getExternalSheet(int externSheetIndex) {
    throw new IllegalStateException("HSSF-style external references are not supported for XSSF");
  }

  @Override
  @NotImplemented
  public ExternalSheet getExternalSheet(String firstSheetName, String lastSheetName, int externalWorkbookNumber) {
    throw new IllegalStateException("ExternalSheets are not supported in excel-streaming-reader");
  }

  @Override
  @NotImplemented
  public int getExternalSheetIndex(String workbookName, String sheetName) {
    throw new IllegalStateException("ExternalSheets are not supported in excel-streaming-reader");
  }

  @Override
  public int getSheetIndex(String sheetName) {
    return _uBook.getSheetIndex(sheetName);
  }

  @Override
  public String getSheetFirstNameByExternSheet(int externSheetIndex) {
    int sheetIndex = convertFromExternalSheetIndex(externSheetIndex);
    return _uBook.getSheetName(sheetIndex);
  }

  @Override
  public String getSheetLastNameByExternSheet(int externSheetIndex) {
    // XSSF does multi-sheet references differently, so this is the same as the first
    return getSheetFirstNameByExternSheet(externSheetIndex);
  }

  @Override
  @NotImplemented
  public String getNameText(NamePtg namePtg) {
    throw new IllegalStateException("getNameText is not supported in excel-streaming-reader");
  }

  @Override
  @NotImplemented
  public EvaluationName getName(NamePtg namePtg) {
    throw new IllegalStateException("EvaluationNames are not supported in excel-streaming-reader");
  }

  @Override
  @NotImplemented
  public XSSFName createName() {
    throw new IllegalStateException("XSSFNames are not supported in excel-streaming-reader");
  }

  @Override
  @NotImplemented
  public XSSFTable getTable(String name) {
    if (name == null) return null;
    throw new IllegalStateException("XSSFTables are not supported in excel-streaming-reader");
  }

  @Override
  public UDFFinder getUDFFinder() {
    return _udfFinder;
  }

  @Override
  public SpreadsheetVersion getSpreadsheetVersion(){
    return SpreadsheetVersion.EXCEL2007;
  }
}
