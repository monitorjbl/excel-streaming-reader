package com.monitorjbl.xlsx;

import com.monitorjbl.xlsx.exceptions.ParseException;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;
import java.io.IOException;
import java.io.InputStream;

public class XmlUtils {
  public static Document document(InputStream is) {
    try {
      return DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(is);
    } catch(SAXException | IOException | ParserConfigurationException e) {
      throw new ParseException(e);
    }
  }

  public static NodeList searchForNodeList(Document document, String xpath) {
    try {
      return (NodeList) XPathFactory.newInstance().newXPath().compile(xpath)
          .evaluate(document, XPathConstants.NODESET);
    } catch(XPathExpressionException e) {
      throw new ParseException(e);
    }
  }

}
