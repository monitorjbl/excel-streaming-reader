package com.github.pjfanning.xlsx.impl.ooxml;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.util.Beta;
import org.apache.poi.util.XMLHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.xml.namespace.QName;
import javax.xml.stream.*;
import javax.xml.stream.events.*;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

@Beta
public class OoXmlStrictConverter implements AutoCloseable {

    private static final Logger LOGGER = LoggerFactory.getLogger(OoXmlStrictConverter.class);
    private static final QName CONFORMANCE = new QName("conformance");
    private static final Properties mappings;
    private static XMLEventFactory XEF;
    private static XMLInputFactory XIF;
    private static XMLOutputFactory XOF;

    static {
        mappings = OoXmlStrictConverterUtils.readMappings();
    }

    private final XMLEventWriter xew;
    private final XMLEventReader xer;
    private int depth = 0;
    private boolean inDateCell;
    private boolean inDateValue;

    public OoXmlStrictConverter(InputStream is, OutputStream os) throws XMLStreamException {
        this.xer = getXmlInputFactory().createXMLEventReader(is);
        this.xew = getXmlOutputFactory().createXMLEventWriter(os);
    }

    public boolean convertNextElement() throws XMLStreamException {
        if (!xer.hasNext()) {
            return false;
        }

        XMLEvent xe = xer.nextEvent();
        if(xe.isStartElement()) {
            xew.add(convertDateStartElement(convertStartElement(xe.asStartElement(), depth==0)));
            depth++;
        } else if(xe.isEndElement()) {
            xew.add(updateDateFlagsOnEndElement(convertEndElement(xe.asEndElement())));
            depth--;
        } else {
            if (inDateValue) {
                xew.add(convertDateValueToNumeric(xe));
            } else {
                // Add as is
                xew.add(xe);
            }
        }

        xew.flush();

        return true;
    }

    private XMLEvent convertDateValueToNumeric(XMLEvent xe) {
        if (!xe.isCharacters()) {
            return xe;
        }

        Date date = DateUtil.parseYYYYMMDDDate(xe.asCharacters().getData());

        double excelDate = DateUtil.getExcelDate(date);

        return getXmlEventFactory().createCharacters(Double.toString(excelDate));
    }

    private EndElement updateDateFlagsOnEndElement(EndElement endElement) {
        if (inDateValue) {
            if ("v".equals(endElement.getName().getLocalPart())) {
                inDateValue = false;
            }
            return endElement;
        }

        if (inDateCell) {
            if (isCell(endElement.getName())) {
                inDateCell = false;
            }
            return endElement;
        }

        return endElement;
    }

    private StartElement convertDateStartElement(StartElement startElement) {

        if (inDateCell) {
            if ("v".equals(startElement.getName().getLocalPart())) {
                this.inDateValue = true;
            }
            return startElement;
        }

        if (!isDateCell(startElement)) {
            return startElement;
        }

        this.inDateCell = true;

        // Change to numeric cell.
        return getXmlEventFactory().createStartElement(startElement.getName(),
                changeTypeAttributeToNumeric(startElement.getAttributes()),
                startElement.getNamespaces());

    }

    private Iterator<? extends Attribute> changeTypeAttributeToNumeric(
            Iterator<Attribute> attributes) {
        List<Attribute> result = new ArrayList<>();

        while (attributes.hasNext()) {
            Attribute attribute = attributes.next();
            if (!"t".equals(attribute.getName().getLocalPart())) {
                result.add(attribute);
                continue;
            }

            result.add(getXmlEventFactory().createAttribute(attribute.getName(), "n"));
        }

        return Collections.unmodifiableList(result).iterator();
    }

    private boolean isDateCell(StartElement startElement) {
        if (!isCell(startElement.getName())) {
            return false;
        }

        Attribute typeAttribute = startElement.getAttributeByName(QName.valueOf("t"));
        if (typeAttribute == null) {
            return false;
        }

        return "d".equals(typeAttribute.getValue());
    }

    private boolean isCell(QName elementName) {
        return "c".equals(elementName.getLocalPart());
    }


    @Override
    public void close() throws XMLStreamException {
        xer.close();
        xew.close();
    }

    private static StartElement convertStartElement(StartElement startElement, boolean root) {
        return getXmlEventFactory().createStartElement(updateQName(startElement.getName()),
                processAttributes(startElement.getAttributes(), startElement.getName().getNamespaceURI(), root),
                processNamespaces(startElement.getNamespaces()));
    }

    private static EndElement convertEndElement(EndElement endElement) {
        return getXmlEventFactory().createEndElement(updateQName(endElement.getName()),
                processNamespaces(endElement.getNamespaces()));

    }

    private static QName updateQName(QName qn) {
        String namespaceUri = qn.getNamespaceURI();
        if(OoXmlStrictConverterUtils.isNotBlank(namespaceUri)) {
            String mappedUri = mappings.getProperty(namespaceUri);
            if(mappedUri != null) {
                qn = OoXmlStrictConverterUtils.isBlank(qn.getPrefix()) ? new QName(mappedUri, qn.getLocalPart())
                        : new QName(mappedUri, qn.getLocalPart(), qn.getPrefix());
            }
        }
        return qn;
    }

    private static Iterator<Attribute> processAttributes(final Iterator<Attribute> iter,
            final String elementNamespaceUri, final boolean rootElement) {
        ArrayList<Attribute> list = new ArrayList<>();
        while(iter.hasNext()) {
            Attribute att = iter.next();
            QName qn = updateQName(att.getName());
            if(rootElement && mappings.containsKey(elementNamespaceUri) && att.getName().equals(CONFORMANCE)) {
                //drop attribute
            } else {
                String newValue = att.getValue();
                for(String key : mappings.stringPropertyNames()) {
                    if(att.getValue().startsWith(key)) {
                        newValue = att.getValue().replace(key, mappings.getProperty(key));
                        break;
                    }
                }
                list.add(getXmlEventFactory().createAttribute(qn, newValue));
            }
        }
        return Collections.unmodifiableList(list).iterator();
    }

    private static Iterator<Namespace> processNamespaces(final Iterator<Namespace> iter) {
        ArrayList<Namespace> list = new ArrayList<>();
        while(iter.hasNext()) {
            Namespace ns = iter.next();
            if(!ns.isDefaultNamespaceDeclaration() && !mappings.containsKey(ns.getNamespaceURI())) {
                list.add(ns);
            }
        }
        return Collections.unmodifiableList(list).iterator();
    }

    private static XMLInputFactory getXmlInputFactory() {
        if (XIF == null) {
            try {
                XIF = XMLHelper.newXMLInputFactory();
            } catch (Throwable t) {
                LOGGER.error("Issue creating XMLInputFactory", t);
                throw t;
            }
        }
        return XIF;
    }

    private static XMLOutputFactory getXmlOutputFactory() {
        if (XOF == null) {
            try {
                XOF = XMLHelper.newXMLOutputFactory();
            } catch (Throwable t) {
                LOGGER.error("Issue creating XMLOutputFactory", t);
                throw t;
            }
        }
        return XOF;
    }

    private static XMLEventFactory getXmlEventFactory() {
        if (XEF == null) {
            try {
                XEF = XMLHelper.newXMLEventFactory();
            } catch (Throwable t) {
                LOGGER.error("Issue creating XMLEventFactory", t);
                throw t;
            }
        }
        return XEF;
    }
}
