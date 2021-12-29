package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.EvaluationSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.Internal;

/**
 * wrapper for a sheet under evaluation (only supports current row)
 */
@Internal
final class CurrentRowEvaluationSheet implements EvaluationSheet {
  private final Sheet _xs;
  private final Row _row;

  CurrentRowEvaluationSheet(Sheet sheet, Row row) {
    _xs = sheet;
    _row = row;
  }

  Sheet getSheet() {
    return _xs;
  }

  /* (non-Javadoc)
   * @see org.apache.poi.ss.formula.EvaluationSheet#getlastRowNum()
   * @since POI 4.0.0
   */
  @Override
  public int getLastRowNum() {
    return _xs.getLastRowNum();
  }

  /* (non-Javadoc)
   * @see org.apache.poi.ss.formula.EvaluationSheet#isRowHidden(int)
   * @since POI 4.1.0
   */
  @Override
  public boolean isRowHidden(int rowIndex) {
    if (_row == null) return false;
    return _row.getZeroHeight();
  }

  @Override
  public EvaluationCell getCell(int rowIndex, int columnIndex) {
    if (_row == null) {
      return null;
    }
    Cell cell = _row.getCell(columnIndex);
    if (cell == null) {
      return null;
    }
    return new OoxmlEvaluationCell(cell, this);
  }

  /* (non-JavaDoc), inherit JavaDoc from EvaluationSheet
   * @since POI 3.15 beta 3
   */
  @Override
  public void clearAllCachedResultValues() {
    //this class does not cache results
  }
}
