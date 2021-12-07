package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.EvaluationSheet;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.Internal;

/**
 * wrapper around the workbook
 */
@Internal
public final class CurrentRowEvaluationWorkbook extends BaseEvaluationWorkbook {

  private final Row _row;

  CurrentRowEvaluationWorkbook(Workbook wb, Row row) {
    super(wb);
    _row = row;
  }

  @Override
  public int getSheetIndex(EvaluationSheet evalSheet) {
    Sheet sheet = ((CurrentRowEvaluationSheet)evalSheet).getSheet();
    return _uBook.getSheetIndex(sheet);
  }

  @Override
  public EvaluationSheet getSheet(int sheetIndex) {
    return new CurrentRowEvaluationSheet(_uBook.getSheetAt(sheetIndex), _row);
  }

  @Override
  public Ptg[] getFormulaTokens(EvaluationCell evalCell) {
    Cell cell = ((OoxmlEvaluationCell)evalCell).getCell();
    return FormulaParser.parse(cell.getCellFormula(), this, FormulaType.CELL, _uBook.getSheetIndex(cell.getSheet()));
  }
}

