package com.monitorjbl.xlsx.impl;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.io.IOException;

import static com.monitorjbl.xlsx.XmlUtils.document;
import static com.monitorjbl.xlsx.XmlUtils.searchForNodeList;

public class WorkbookUtil {
    public static boolean use1904Dates(XSSFReader reader) throws IOException, InvalidFormatException {
      NodeList workbookPr = searchForNodeList(document(reader.getWorkbookData()), "/ss:workbook/ss:workbookPr");
      if (workbookPr.getLength() == 1) {
        final Node date1904 = workbookPr.item(0).getAttributes().getNamedItem("date1904");
        if (date1904 != null) {
          String value = date1904.getTextContent();
          return "1".equals(value) || "true".equals(value);
        }
      }
      return false;
    }
}
