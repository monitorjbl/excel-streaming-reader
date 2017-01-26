package com.monitorjbl.xlsx.sst;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Characters;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

public class BufferedStringsTable extends SharedStringsTable implements AutoCloseable {
  private final FileBackedList<CTRstImpl> list;

  private BufferedStringsTable(PackagePart part, PackageRelationship rel, File file, int cacheSize) throws IOException {
    this.list = new FileBackedList<>(CTRstImpl.class, file, cacheSize);
    readFrom(part.getInputStream());
  }

  @Override
  public void readFrom(InputStream is) throws IOException {
    try {
      XMLInputFactory xmlInputFactory = XMLInputFactory.newInstance();
      XMLEventReader xmlEventReader = xmlInputFactory.createXMLEventReader(is);

      while(xmlEventReader.hasNext()) {
        XMLEvent xmlEvent = xmlEventReader.nextEvent();

        if(xmlEvent.isStartElement() && xmlEvent.asStartElement().getName().getLocalPart().equals("si")) {
          list.add(parseCTRst(xmlEventReader));
        }
      }
    } catch(
        XMLStreamException e)

    {
      throw new IOException(e);
    }

  }

  private CTRstImpl parseCTRst(XMLEventReader xmlEventReader) throws XMLStreamException {
    StartElement ele = xmlEventReader.nextEvent().asStartElement();

    switch(ele.getName().getLocalPart()) {
      case "t":
        Characters chars = xmlEventReader.nextEvent().asCharacters();
        return new CTRstImpl(chars.getData());
      case "phoneticPr":
      case "rPh;":
      case "r":
        return null;
    }
    throw new IllegalArgumentException("");
  }

  public CTRst getEntryAt(int idx) {
    return list.getAt(idx);
  }

  public static SharedStringsTable getSharedStringsTable(File tmp, int cacheSize, OPCPackage pkg) throws IOException, InvalidFormatException {
    ArrayList<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.SHARED_STRINGS.getContentType());
    return parts.size() == 0 ? null : new BufferedStringsTable(parts.get(0), null, tmp, cacheSize);
  }

  @Override
  public void close() {
    list.close();
  }
}
