package com.monitorjbl.xlsx;

import com.monitorjbl.xlsx.exceptions.ParseException;
import org.apache.poi.ooxml.util.DocumentHelper;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.XMLConstants;
import javax.xml.namespace.NamespaceContext;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class XmlUtils {
  public static Document document(InputStream is) {
    try {
      return DocumentHelper.readDocument(is);
    } catch(SAXException | IOException e) {
      throw new ParseException(e);
    }
  }

  public static NodeList searchForNodeList(Document document, String xpath) {
    try {
      XPath xp = XPathFactory.newInstance().newXPath();
      NamespaceContextImpl nc = new NamespaceContextImpl();
      nc.addNamespace("ss", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
      xp.setNamespaceContext(nc);
      return (NodeList)xp.compile(xpath)
              .evaluate(document, XPathConstants.NODESET);
    } catch(XPathExpressionException e) {
      throw new ParseException(e);
    }
  }

  private static class NamespaceContextImpl implements NamespaceContext {
    private Map<String, String> urisByPrefix = new HashMap<String, String>();

    private Map<String, Set> prefixesByURI = new HashMap<String, Set>();

    public NamespaceContextImpl() {
      addNamespace(XMLConstants.XML_NS_PREFIX, XMLConstants.XML_NS_URI);
      addNamespace(XMLConstants.XMLNS_ATTRIBUTE, XMLConstants.XMLNS_ATTRIBUTE_NS_URI);
    }

    public void addNamespace(String prefix, String namespaceURI) {
      urisByPrefix.put(prefix, namespaceURI);
      if (prefixesByURI.containsKey(namespaceURI)) {
        (prefixesByURI.get(namespaceURI)).add(prefix);
      } else {
        Set<String> set = new HashSet<String>();
        set.add(prefix);
        prefixesByURI.put(namespaceURI, set);
      }
    }

    public String getNamespaceURI(String prefix) {
      if (prefix == null)
        throw new IllegalArgumentException("prefix cannot be null");
      if (urisByPrefix.containsKey(prefix))
        return (String) urisByPrefix.get(prefix);
      else
        return XMLConstants.NULL_NS_URI;
    }

    public String getPrefix(String namespaceURI) {
      return (String) getPrefixes(namespaceURI).next();
    }

    public Iterator getPrefixes(String namespaceURI) {
      if (namespaceURI == null)
        throw new IllegalArgumentException("namespaceURI cannot be null");
      if (prefixesByURI.containsKey(namespaceURI)) {
        return ((Set) prefixesByURI.get(namespaceURI)).iterator();
      } else {
        return Collections.EMPTY_SET.iterator();
      }
    }
  }
}
