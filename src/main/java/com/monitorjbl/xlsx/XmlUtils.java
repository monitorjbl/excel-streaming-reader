package com.monitorjbl.xlsx;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Method;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.monitorjbl.xlsx.exceptions.ParseException;

public class XmlUtils {
  private static final Logger log = LoggerFactory.getLogger(XmlUtils.class);

  public static Document document(InputStream is) {
    try {
      DocumentBuilderFactory documentBuilderFactory = DocumentBuilderFactory.newInstance();
      documentBuilderFactory.setValidating(false);
      trySetXercesSecurityManager(documentBuilderFactory);
      trySetSAXFeature(documentBuilderFactory, "http://javax.xml.XMLConstants/feature/secure-processing", true);
      return documentBuilderFactory.newDocumentBuilder().parse(is);
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

  private static void trySetSAXFeature(DocumentBuilderFactory dbf, String feature, boolean enabled) {
    try {
      dbf.setFeature(feature, enabled);
    } catch (Exception e) {
      log.warn("SAX Feature unsupported: {}", feature);
    } catch (AbstractMethodError e) {
      log.warn("Cannot set SAX feature {} because outdated XML parser in classpath", feature);
    }

  }

  private static void trySetXercesSecurityManager(DocumentBuilderFactory dbf) {
    String[] classNames = new String[]{"com.sun.org.apache.xerces.internal.util.SecurityManager", "org.apache.xerces.util.SecurityManager"};

    for (String securityManagerClassName : classNames) {
      try {
        Object mgr = Class.forName(securityManagerClassName).newInstance();
        Method setLimit = mgr.getClass().getMethod("setEntityExpansionLimit", Integer.TYPE);
        setLimit.invoke(mgr, Integer.valueOf(4096));
        dbf.setAttribute("http://apache.org/xml/properties/security-manager", mgr);
        return;
      } catch (Throwable t) {
        // allow to iterate over classNames
      }
    }
    log.warn("SAX Security Manager could not be setup");
  }
}
