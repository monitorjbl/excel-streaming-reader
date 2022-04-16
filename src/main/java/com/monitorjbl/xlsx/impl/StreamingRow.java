package com.monitorjbl.xlsx.impl;

import com.monitorjbl.xlsx.notsupportedoperations.RowAdapter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;

public class StreamingRow implements RowAdapter {
	private final Sheet sheet;
	private final int rowIndex;
	private final boolean isHidden;
	private TreeMap<Integer, Cell> cellMap = new TreeMap<>();

	public StreamingRow(Sheet sheet, int rowIndex, boolean isHidden) {
		this.sheet = sheet;
		this.rowIndex = rowIndex;
		this.isHidden = isHidden;
	}

	@Override
	public Sheet getSheet() {
    return sheet;
  }

  public Map<Integer, Cell> getCellMap() {
    return cellMap;
  }

  public void setCellMap(TreeMap<Integer, Cell> cellMap) {
    this.cellMap = cellMap;
  }

 /* Supported */

  /**
   * Get row number this row represents
   *
   * @return the row number (0 based)
   */
  @Override
  public int getRowNum() {
    return rowIndex;
  }

  /**
   * @return Cell iterator of the physically defined cells for this row.
   */
  @Override
  public Iterator<Cell> cellIterator() {
    return cellMap.values().iterator();
  }

  /**
   * @return Cell iterator of the physically defined cells for this row.
   */
  @Override
  public Iterator<Cell> iterator() {
    return cellMap.values().iterator();
  }

  /**
   * Get the cell representing a given column (logical cell) 0-based.  If you
   * ask for a cell that is not defined, you get a null.
   *
   * @param cellnum 0 based column number
   * @return Cell representing that column or null if undefined.
   */
  @Override
  public Cell getCell(int cellnum) {
    return cellMap.get(cellnum);
  }

  /**
   * Gets the index of the last cell contained in this row <b>PLUS ONE</b>.
   *
   * @return short representing the last logical cell in the row <b>PLUS ONE</b>,
   * or -1 if the row does not contain any cells.
   */
  @Override
  public short getLastCellNum() {
    return (short) (cellMap.size() == 0 ? -1 : cellMap.lastEntry().getValue().getColumnIndex() + 1);
  }

  /**
   * Get whether or not to display this row with 0 height
   *
   * @return - zHeight height is zero or not.
   */
  @Override
  public boolean getZeroHeight() {
    return isHidden;
  }

  /**
   * Gets the number of defined cells (NOT number of cells in the actual row!).
   * That is to say if only columns 0,4,5 have values then there would be 3.
   *
   * @return int representing the number of defined cells in the row.
   */
  @Override
  public int getPhysicalNumberOfCells() {
    return cellMap.size();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public short getFirstCellNum() {
    if(cellMap.size() == 0) {
      return -1;
    }
    return cellMap.firstKey().shortValue();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Cell getCell(int cellnum, MissingCellPolicy policy) {
    StreamingCell cell = (StreamingCell) cellMap.get(cellnum);
    if(policy == MissingCellPolicy.CREATE_NULL_AS_BLANK) {
      if(cell == null) { return new StreamingCell(sheet, cellnum, rowIndex, false); }
    } else if(policy == MissingCellPolicy.RETURN_BLANK_AS_NULL) {
      if(cell == null || cell.getCellType() == CellType.BLANK) { return null; }
    }
    return cell;
  }

}
