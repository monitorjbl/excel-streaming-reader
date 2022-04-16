package com.monitorjbl.xlsx.utils;

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

public final class XmlUtils {
  private XmlUtils() {
    throw new RuntimeException("It is not good practice to instantiate utility classes.");
  }

  public static Document document(InputStream is) {
    try {
      return DocumentHelper.readDocument(is);
    } catch (SAXException | IOException e) {
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
    private final Map<String, String> urisByPrefix;
    private final Map<String, Set<String>> prefixesByURI;

    public NamespaceContextImpl() {
      prefixesByURI = new HashMap<>();
      urisByPrefix = new HashMap<>();
      addNamespace(XMLConstants.XML_NS_PREFIX, XMLConstants.XML_NS_URI);
      addNamespace(XMLConstants.XMLNS_ATTRIBUTE, XMLConstants.XMLNS_ATTRIBUTE_NS_URI);
    }

    public void addNamespace(String prefix, String namespaceURI) {
      urisByPrefix.put(prefix, namespaceURI);
      if (prefixesByURI.containsKey(namespaceURI)) {
        (prefixesByURI.get(namespaceURI)).add(prefix);
      } else {
        Set<String> set = new HashSet<>();
        set.add(prefix);
        prefixesByURI.put(namespaceURI, set);
      }
    }

    public String getNamespaceURI(String prefix) {
      if (prefix == null)
        throw new IllegalArgumentException("prefix cannot be null");
      return urisByPrefix.getOrDefault(prefix, XMLConstants.NULL_NS_URI);
    }

    public String getPrefix(String namespaceURI) {
      return getPrefixes(namespaceURI).next();
    }

    public Iterator<String> getPrefixes(String namespaceURI) {
      if (namespaceURI == null)
        throw new IllegalArgumentException("namespaceURI cannot be null");
      return prefixesByURI.getOrDefault(namespaceURI, Collections.EMPTY_SET).iterator();
    }
  }
}
