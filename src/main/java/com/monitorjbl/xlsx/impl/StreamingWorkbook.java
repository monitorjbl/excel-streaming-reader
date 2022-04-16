package com.monitorjbl.xlsx.impl;

import com.monitorjbl.xlsx.exceptions.MissingSheetException;
import com.monitorjbl.xlsx.notsupportedoperations.WorkbookAdapter;
import org.apache.poi.ss.formula.EvaluationWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.IOException;
import java.util.Iterator;

public class StreamingWorkbook implements WorkbookAdapter, AutoCloseable {
  private final StreamingWorkbookReader reader;

  public StreamingWorkbook(StreamingWorkbookReader reader) {
    this.reader = reader;
  }

  int findSheetByName(String name) {
    for (int i = 0; i < reader.getSheetProperties().size(); i++) {
      if (reader.getSheetProperties().get(i).get("name").equals(name)) {
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

  @Override
  public EvaluationWorkbook createEvaluationWorkbook() {
    return null;
  }
}
