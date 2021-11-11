package com.github.pjfanning.xlsx;

import org.apache.poi.ss.util.CellAddress;

public class SharedFormula {

  private final CellAddress cellAddress;
  private final String formula;

  public SharedFormula(CellAddress cellAddress, String formula) {
    this.cellAddress = cellAddress;
    this.formula = formula;
  }

  public CellAddress getCellAddress() {
    return cellAddress;
  }

  public String getFormula() {
    return formula;
  }
}
