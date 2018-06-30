package com.monitorjbl.xlsx.impl;

import com.monitorjbl.xlsx.exceptions.MissingSheetException;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;

public class StreamingWorkbook implements Workbook, AutoCloseable {
  private final StreamingWorkbookReader reader;

  public StreamingWorkbook(StreamingWorkbookReader reader) {
    this.reader = reader;
  }

  int findSheetByName(String name) {
    for(int i = 0; i < reader.getSheetProperties().size(); i++) {
      if(reader.getSheetProperties().get(i).get("name").equals(name)) {
        return i;
      }
    }
    return -1;
  }

  /* Supported */

  /**
   * {@inheritDoc}
   */
  @Override
  public Iterator<Sheet> iterator() {
    return reader.iterator();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Iterator<Sheet> sheetIterator() {
    return iterator();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public String getSheetName(int sheet) {
    return reader.getSheetProperties().get(sheet).get("name");
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public int getSheetIndex(String name) {
    return findSheetByName(name);
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public int getSheetIndex(Sheet sheet) {
    if(sheet instanceof StreamingSheet) {
      return findSheetByName(sheet.getSheetName());
    } else {
      throw new UnsupportedOperationException("Cannot use non-StreamingSheet sheets");
    }
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public int getNumberOfSheets() {
    return reader.getSheets().size();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Sheet getSheetAt(int index) {
    return reader.getSheets().get(index);
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Sheet getSheet(String name) {
    int index = getSheetIndex(name);
    if(index == -1) {
      throw new MissingSheetException("Sheet '" + name + "' does not exist");
    }
    return reader.getSheets().get(index);
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public boolean isSheetHidden(int sheetIx) {
    return "hidden".equals(reader.getSheetProperties().get(sheetIx).get("state"));
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public boolean isSheetVeryHidden(int sheetIx) {
    return "veryHidden".equals(reader.getSheetProperties().get(sheetIx).get("state"));
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public void close() throws IOException {
    reader.close();
  }

  /* Not supported */

  /**
   * Not supported
   */
  @Override
  public int getActiveSheetIndex() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setActiveSheet(int sheetIndex) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public int getFirstVisibleTab() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setFirstVisibleTab(int sheetIndex) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setSheetOrder(String sheetname, int pos) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setSelectedTab(int index) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setSheetName(int sheet, String name) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public Sheet createSheet() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public Sheet createSheet(String sheetname) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public Sheet cloneSheet(int sheetNum) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void removeSheetAt(int index) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public Font createFont() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public Font findFont(boolean b, short i, short i1, String s, boolean b1, boolean b2, short i2, byte b3) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public short getNumberOfFonts() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public int getNumberOfFontsAsInt() { throw new UnsupportedOperationException(); }

  /**
   * Not supported
   */
  @Override
  public Font getFontAt(short idx) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public Font getFontAt(int i) { throw new UnsupportedOperationException(); }

  /**
   * Not supported
   */
  @Override
  public CellStyle createCellStyle() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public int getNumCellStyles() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public CellStyle getCellStyleAt(int i) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void write(OutputStream stream) throws IOException {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public int getNumberOfNames() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public Name getName(String name) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public List<? extends Name> getNames(String s) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public List<? extends Name> getAllNames() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public Name getNameAt(int nameIndex) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public Name createName() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public int getNameIndex(String name) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void removeName(int index) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void removeName(String name) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void removeName(Name name) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public int linkExternalWorkbook(String name, Workbook workbook) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setPrintArea(int sheetIndex, String reference) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public String getPrintArea(int sheetIndex) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void removePrintArea(int sheetIndex) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public MissingCellPolicy getMissingCellPolicy() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setMissingCellPolicy(MissingCellPolicy missingCellPolicy) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public DataFormat createDataFormat() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public int addPicture(byte[] pictureData, int format) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public List<? extends PictureData> getAllPictures() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public CreationHelper getCreationHelper() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean isHidden() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setHidden(boolean hiddenFlag) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setSheetHidden(int sheetIx, boolean hidden) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public SheetVisibility getSheetVisibility(int i) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setSheetVisibility(int i, SheetVisibility sheetVisibility) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void addToolPack(UDFFinder toopack) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public void setForceFormulaRecalculation(boolean value) {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public boolean getForceFormulaRecalculation() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public SpreadsheetVersion getSpreadsheetVersion() {
    throw new UnsupportedOperationException();
  }

  /**
   * Not supported
   */
  @Override
  public int addOlePackage(byte[] bytes, String s, String s1, String s2) throws IOException {
    throw new UnsupportedOperationException();
  }
}
