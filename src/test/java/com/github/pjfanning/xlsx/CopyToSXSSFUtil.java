package com.github.pjfanning.xlsx;

import com.github.pjfanning.xlsx.impl.XlsxHyperlink;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.InputStream;

public class CopyToSXSSFUtil {
  public static SXSSFWorkbook copyToSXSSF(final InputStream inputStream) throws Exception {
    SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(5);
    try (Workbook wbInput = StreamingReader.builder().setReadHyperlinks(true).open(inputStream)) {
      //note that StreamingReader.builder().setReadHyperlinks(true) and that cellCopyPolicy.setCopyHyperlink(false)
      //the hyperlinks appear at end of sheet, so we need to iterate them separately at the end
      final CellCopyPolicy cellCopyPolicy = new CellCopyPolicy();
      cellCopyPolicy.setCopyHyperlink(false);
      final CellCopyContext cellCopyContext = new CellCopyContext();
      for (Sheet sheetInput : wbInput) {
        SXSSFSheet sheetOutput = sxssfWorkbook.createSheet(sheetInput.getSheetName());
        for (Row rowInput : sheetInput) {
          SXSSFRow rowOutput = sheetOutput.createRow(rowInput.getRowNum());
          for (Cell cellInput : rowInput) {
            SXSSFCell cellOutput = rowOutput.createCell(cellInput.getColumnIndex());
            CellUtil.copyCell(cellInput, cellOutput, cellCopyPolicy, cellCopyContext);
          }
        }
        //POI 5.2.3 adds a SXSSFSheet.addHyperlink so there will no need tp get the XSSFSheet
        XSSFSheet xssfSheet = sxssfWorkbook.getXSSFWorkbook().getSheet(sheetInput.getSheetName());
        for (Hyperlink hyperlink : sheetInput.getHyperlinkList()) {
          xssfSheet.addHyperlink(((XlsxHyperlink)hyperlink).createXSSFHyperlink());
        }
      }
    } catch (Exception e) {
      sxssfWorkbook.dispose();
      throw e;
    }
    return sxssfWorkbook;
  }
}
