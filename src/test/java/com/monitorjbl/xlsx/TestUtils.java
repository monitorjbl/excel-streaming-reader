package com.monitorjbl.xlsx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import static org.junit.jupiter.api.Assertions.assertEquals;

final class TestUtils {

  static Workbook openWorkbook(String fileName) throws IOException {
    try(InputStream stream = TestUtils.class.getResourceAsStream("/" + fileName)) {
      return StreamingReader.builder()
          .open(stream);
    }
  }

  static void expectSameStringContent(Cell cell1, Cell cell2) {
    assertEquals(cell1.getStringCellValue(), cell2.getStringCellValue(),
        "Cell " + ref(cell1) + " has should equal cell " + ref(cell2) + " string value.");
  }

  static void expectStringContent(Cell cell, String value) {
    assertEquals(value, cell.getStringCellValue(), "Cell " + ref(cell) + " has wrong string content.");
  }

  static void expectCachedType(Cell cell, CellType cellType) {
    assertEquals(cellType, cell.getCachedFormulaResultType(), "Cell " + ref(cell) + " has wrong cached type." + cellType);
  }

  static void expectType(Cell cell, CellType cellType) {
    assertEquals(cellType, cell.getCellType(), "Cell " + ref(cell) + " has wrong type.");
  }

  static void expectFormula(Cell cell, String formula) {
    assertEquals(formula, cell.getCellFormula(), "Cell " + ref(cell) + " has wrong formula.");
  }

  private static String ref(Cell cell) {
    return new CellReference(cell).formatAsString();
  }

  static Cell getCellFromNextRow(Iterator<Row> rowIterator, int index) {
    return nextRow(rowIterator)
        .getCell(index);
  }

  static Row nextRow(Iterator<Row> rowIterator) {
    return rowIterator.next();
  }

}
