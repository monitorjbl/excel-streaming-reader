package com.monitorjbl.xlsx.impl;

import com.monitorjbl.xlsx.notsupportedoperations.SheetAdapter;
import org.apache.poi.ss.usermodel.Row;

import java.util.Iterator;

public class StreamingSheet implements SheetAdapter {

	private final String name;
	private final StreamingSheetReader reader;

	public StreamingSheet(String name, StreamingSheetReader reader) {
		this.name = name;
		this.reader = reader;
		reader.setSheet(this);
	}

	StreamingSheetReader getReader() {
    return reader;
  }

  /* Supported */

  /**
   * {@inheritDoc}
   */
  @Override
  public Iterator<Row> iterator() {
    return reader.iterator();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Iterator<Row> rowIterator() {
    return reader.iterator();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public String getSheetName() {
    return name;
  }

  /**
   * Get the hidden state for a given column
   *
   * @param columnIndex - the column to set (0-based)
   * @return hidden - <code>false</code> if the column is visible
   */
  @Override
  public boolean isColumnHidden(int columnIndex) {
    return reader.isColumnHidden(columnIndex);
  }

  /**
   * Gets the last row on the sheet
   *
   * @return last row contained n this sheet (0-based)
   */
  @Override
  public int getLastRowNum() {
    return reader.getLastRowNum();
  }

}
