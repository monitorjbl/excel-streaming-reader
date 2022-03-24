/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package com.github.pjfanning.xlsx.impl.ooxml;

import com.github.pjfanning.poi.xssf.streaming.MapBackedCommentsTable;
import com.github.pjfanning.poi.xssf.streaming.MapBackedSharedStringsTable;
import com.github.pjfanning.poi.xssf.streaming.TempFileCommentsTable;
import com.github.pjfanning.poi.xssf.streaming.TempFileSharedStringsTable;
import com.github.pjfanning.xlsx.CommentsImplementationType;
import com.github.pjfanning.xlsx.StreamingReader;
import com.github.pjfanning.xlsx.exceptions.OpenException;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.util.Internal;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.*;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.xmlbeans.XmlException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLStreamException;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@Internal
public class OoxmlReader extends XSSFReader {

  private static final Logger LOGGER = LoggerFactory.getLogger(OoxmlReader.class);
  static final String PURL_COMMENTS_RELATIONSHIP_URL = "http://purl.oclc.org/ooxml/officeDocument/relationships/comments";
  static final String PURL_DRAWING_RELATIONSHIP_URL = "http://purl.oclc.org/ooxml/officeDocument/relationships/drawing";

  private static final Set<String> OVERRIDE_WORKSHEET_RELS =
          Collections.unmodifiableSet(new HashSet<>(
                  Arrays.asList(XSSFRelation.WORKSHEET.getRelation(),
                          "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet",
                          XSSFRelation.CHARTSHEET.getRelation(),
                          XSSFRelation.MACRO_SHEET_BIN.getRelation())
          ));

  private final OoxmlSheetReader ooxmlSheetReader;

  /**
   * Creates a new XSSFReader, for the given package
   */
  @Internal
  public OoxmlReader(StreamingReader.Builder builder,
                     OPCPackage pkg, boolean strictOoxmlChecksNeeded) throws IOException, OpenXML4JException {
    super(pkg, true);

    PackageRelationship coreDocRelationship = this.pkg.getRelationshipsByType(
            PackageRelationshipTypes.CORE_DOCUMENT).getRelationship(0);

    // strict OOXML likely not fully supported, see #57699
    // this code is similar to POIXMLDocumentPart.getPartFromOPCPackage(), but I could not combine it
    // easily due to different return values
    if (coreDocRelationship == null) {
      coreDocRelationship = this.pkg.getRelationshipsByType(
              PackageRelationshipTypes.STRICT_CORE_DOCUMENT).getRelationship(0);

      if (coreDocRelationship == null) {
        throw new POIXMLException("OOXML file structure broken/invalid - no core document found!");
      }
    }

    // Get the part that holds the workbook
    workbookPart = this.pkg.getPart(coreDocRelationship);
    ooxmlSheetReader = new OoxmlSheetReader(builder, workbookPart, strictOoxmlChecksNeeded);
  }


  /**
   * Opens up the Shared Strings Table, parses it, and
   * returns a handy object for working with
   * shared strings.
   */
  @Override
  public SharedStringsTable getSharedStringsTable() throws IOException {
    ArrayList<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.SHARED_STRINGS.getContentType());
    return parts.isEmpty() ? null : new SharedStringsTable(parts.get(0));
  }

  /**
   * Opens up the Shared Strings Table, parses it, and
   * returns a handy object for working with
   * shared strings.
   */
  public SharedStrings getSharedStrings(StreamingReader.Builder builder) throws IOException, SAXException {
    ArrayList<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.SHARED_STRINGS.getContentType());
    if (parts.isEmpty()) {
      return null;
    } else {
      switch (builder.getSharedStringsImplementationType()) {
        case POI_DEFAULT:
          return new SharedStringsTable(parts.get(0));
        case CUSTOM_MAP_BACKED:
          return new MapBackedSharedStringsTable(parts.get(0).getPackage());
        case TEMP_FILE_BACKED:
          return new TempFileSharedStringsTable(parts.get(0).getPackage(), builder.encryptSstTempFile());
        default:
          return new ReadOnlySharedStringsTable(parts.get(0));
      }
    }
  }

  /**
   * Opens up the Styles Table, parses it, and
   * returns a handy object for working with cell styles
   */
  @Override
  public StylesTable getStylesTable() throws IOException, InvalidFormatException {
    ArrayList<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.STYLES.getContentType());
    if (parts.isEmpty()) return null;

    // Create the Styles Table, and associate the Themes if present
    StylesTable styles = new StylesTable(parts.get(0));
    parts = pkg.getPartsByContentType(XSSFRelation.THEME.getContentType());
    if (!parts.isEmpty()) {
      styles.setTheme(new ThemesTable(parts.get(0)));
    }
    return styles;
  }

  /**
   * Returns an Iterator which will let you get at all the
   * different Sheets in turn.
   * Each sheet's InputStream is only opened when fetched
   * from the Iterator. It's up to you to close the
   * InputStreams when done with each one.
   */
  public Iterator<SheetData> sheetIterator() {
    return ooxmlSheetReader.iterator();
  }

  public SheetData getSheetData(final String name) {
    return ooxmlSheetReader.getSheetData(name);
  }

  public SheetData getSheetDataAt(final int idx) {
    return ooxmlSheetReader.getSheetData(idx);
  }

  public int getNumberOfSheets() {
    return ooxmlSheetReader.size();
  }

  static class OoxmlSheetReader {

    public final StreamingReader.Builder builder;

    /**
     * Maps relId and the corresponding PackagePart
     */
    private final Map<String, PackagePart> sheetMap;

    /**
     * List over CTSheet objects, returns sheets in {@code logical} order.
     * We can't rely on the Ooxml4J's relationship iterator because it returns objects in physical order,
     * i.e. as they are stored in the underlying package
     */
    private final ArrayList<XSSFSheetRef> sheetRefList;

    private final boolean strictOoxmlChecksNeeded;

    private final Map<XSSFSheetRef, SheetData> sheetDataMap = new HashMap<>();

    /**
     * Construct a new SheetIterator
     *
     * @param wb package part holding workbook.xml
     */
    OoxmlSheetReader(final StreamingReader.Builder builder,
                     final PackagePart wb, final boolean strictOoxmlChecksNeeded) throws IOException {
      this.builder = builder;
      this.strictOoxmlChecksNeeded = strictOoxmlChecksNeeded;
      /*
       * The order of sheets is defined by the order of CTSheet elements in workbook.xml
       */
      try {
        //step 1. Map sheet's relationship Id and the corresponding PackagePart
        sheetMap = new HashMap<>();
        OPCPackage pkg = wb.getPackage();
        Set<String> worksheetRels = getSheetRelationships();
        for (PackageRelationship rel : wb.getRelationships()) {
          String relType = rel.getRelationshipType();
          if (worksheetRels.contains(relType)) {
            PackagePartName relName = PackagingURIHelper.createPartName(rel.getTargetURI());
            sheetMap.put(rel.getId(), pkg.getPart(relName));
          }
        }
        //step 2. Read array of CTSheet elements, wrap it in a LinkedList
        sheetRefList = createSheetListFromWB(wb);
      } catch (InvalidFormatException e) {
        throw new POIXMLException(e);
      }
    }

    int size() {
      return sheetRefList.size();
    }

    OoxmlSheetIterator iterator() {
      return new OoxmlSheetIterator(this, sheetRefList);
    }

    SheetData getSheetData(final String name) {
      XSSFSheetRef matchedSheetRef = null;
      for (XSSFSheetRef sheetRef : sheetRefList) {
        if (name.equalsIgnoreCase(sheetRef.getName())) {
          matchedSheetRef = sheetRef;
          break;
        }
      }
      if (matchedSheetRef == null) {
        throw new NoSuchElementException("Failed to find sheet " + name);
      }
      return getSheetData(matchedSheetRef);
    }

    SheetData getSheetData(final int idx) {
      XSSFSheetRef matchedSheetRef = sheetRefList.get(idx);
      if (matchedSheetRef == null) {
        throw new NoSuchElementException("Failed to find sheet with id " + idx);
      }
      return getSheetData(matchedSheetRef);
    }

    SheetData getSheetData(final XSSFSheetRef sheetRef) {
      SheetData sd = sheetDataMap.get(sheetRef);
      if (sd == null) {
        sd = createSheetData(builder, sheetRef, sheetMap, strictOoxmlChecksNeeded);
        sheetDataMap.put(sheetRef, sd);
      }
      return sd;
    }

    private ArrayList<XSSFSheetRef> createSheetListFromWB(PackagePart wb) throws IOException {

      XMLSheetRefReader xmlSheetRefReader = new XMLSheetRefReader();
      XMLReader xmlReader;
      try {
        xmlReader = XMLHelper.newXMLReader();
      } catch (ParserConfigurationException | SAXException e) {
        throw new POIXMLException(e);
      }
      xmlReader.setContentHandler(xmlSheetRefReader);
      try (InputStream stream = wb.getInputStream()) {
        xmlReader.parse(new InputSource(stream));
      } catch (SAXException e) {
        throw new OpenException(e);
      }

      final ArrayList<XSSFSheetRef> validSheets = new ArrayList<>();
      for (XSSFSheetRef xssfSheetRef : xmlSheetRefReader.getSheetRefs()) {
        //if there's no relationship id, silently skip the sheet
        String sheetId = xssfSheetRef.getId();
        if (sheetId != null && sheetId.length() > 0) {
          validSheets.add(xssfSheetRef);
        }
      }
      return validSheets;
    }

    /**
     * Gets string representations of relationships
     * that are sheet-like.  Added to allow subclassing
     * by XSSFBReader.  This is used to decide what
     * relationships to load into the sheetRefs
     *
     * @return all relationships that are sheet-like
     */
    private Set<String> getSheetRelationships() {
      return OVERRIDE_WORKSHEET_RELS;
    }
  }

  /**
   * Iterator over sheet data.
   */
  static class OoxmlSheetIterator implements Iterator<SheetData> {

    /**
     * List over CTSheet objects, returns sheets in {@code logical} order.
     * We can't rely on the Ooxml4J's relationship iterator because it returns objects in physical order,
     * i.e. as they are stored in the underlying package
     */
    private final ArrayList<XSSFSheetRef> sheetRefList;

    private final OoxmlSheetReader ooxmlSheetReader;

    private int sheetRefPosition;
    private SheetData sheetData;

    OoxmlSheetIterator(final OoxmlSheetReader ooxmlSheetReader,
                       final ArrayList<XSSFSheetRef> sheetRefList) {
      this.ooxmlSheetReader = ooxmlSheetReader;
      this.sheetRefList = sheetRefList;
    }

    /**
     * Returns {@code true} if the iteration has more elements.
     *
     * @return {@code true} if the iterator has more elements.
     */
    @Override
    public boolean hasNext() {
      return sheetRefPosition < sheetRefList.size();
    }

    /**
     * Returns input stream of the next sheet in the iteration
     *
     * @return input stream of the next sheet in the iteration
     */
    @Override
    public SheetData next() {
      try {
        XSSFSheetRef xssfSheetRef = sheetRefList.get(sheetRefPosition++);
        sheetData = ooxmlSheetReader.getSheetData(xssfSheetRef);
        return sheetData;
      } catch (IndexOutOfBoundsException e) {
        throw new NoSuchElementException("Sheet iterator has no more elements");
      } catch (Exception e) {
        throw new POIXMLException(e);
      }
    }

    /**
     * We're read only, so remove isn't supported
     */
    @Override
    public void remove() {
      throw new IllegalStateException("Not supported");
    }
  }

  public static class SheetData {
    private final PackagePart sheetPart;
    private final String sheetName;
    private final Comments comments;
    private final List<XSSFShape> shapes;

    SheetData(final PackagePart sheetPart, final String sheetName, final Comments comments, List<XSSFShape> shapes) {
      this.sheetPart = sheetPart;
      this.sheetName = sheetName;
      this.comments = comments;
      this.shapes = shapes;
    }

    public PackagePart getSheetPart() {
      return sheetPart;
    }

    public String getSheetName() {
      return sheetName;
    }

    public Comments getComments() {
      return comments;
    }

    public List<XSSFShape> getShapes() {
      return shapes;
    }
  }

  private static SheetData createSheetData(final StreamingReader.Builder builder,
                                           final XSSFSheetRef xssfSheetRef,
                                           final Map<String, PackagePart> sheetMap,
                                           final boolean strictOoxmlChecksNeeded) {
    final PackagePart sheetPart = sheetMap.get(xssfSheetRef.getId());
    final List<XSSFShape> shapes = builder.readShapes() ? getShapes(sheetPart, strictOoxmlChecksNeeded) : null;
    final Comments comments = builder.readComments() ? getSheetComments(builder, sheetPart, strictOoxmlChecksNeeded) : null;
    return new SheetData(sheetPart, xssfSheetRef.getName(), comments, shapes);
  }

  /**
   * Returns the comments associated with this sheet,
   * or null if there aren't any
   */
  private static Comments getSheetComments(final StreamingReader.Builder builder,
                                           final PackagePart sheetPkg,
                                           final boolean strictOoxmlChecksNeeded) {
    // Do we have a comments relationship? (Only ever one if so)
    try {
      PackageRelationshipCollection commentsList =
              sheetPkg.getRelationshipsByType(XSSFRelation.SHEET_COMMENTS.getRelation());
      if (commentsList.size() == 0 && strictOoxmlChecksNeeded) {
        commentsList =
                sheetPkg.getRelationshipsByType(OoxmlReader.PURL_COMMENTS_RELATIONSHIP_URL);
      }
      if (commentsList.size() > 0) {
        PackageRelationship comments = commentsList.getRelationship(0);
        PackagePartName commentsName = PackagingURIHelper.createPartName(comments.getTargetURI());
        PackagePart commentsPart = sheetPkg.getPackage().getPart(commentsName);
        return parseComments(builder, commentsPart, strictOoxmlChecksNeeded);
      }
    } catch (InvalidFormatException | IOException | XMLStreamException e) {
      LOGGER.warn("issue getting sheet comments", e);
      return null;
    }
    return null;
  }

  private static Comments parseComments(final StreamingReader.Builder builder,
                                        final PackagePart commentsPart,
                                        final boolean strictOoxmlChecksNeeded)
          throws IOException, XMLStreamException {
    if (builder.getCommentsImplementationType() == CommentsImplementationType.TEMP_FILE_BACKED) {
      try (InputStream is = commentsPart.getInputStream()) {
        TempFileCommentsTable ct = new TempFileCommentsTable(
                builder.encryptCommentsTempFile(),
                builder.fullFormatRichText());
        try {
          ct.readFrom(is);
        } catch (IOException | RuntimeException e) {
          ct.close();
          throw e;
        }
        return ct;
      }
    } else if (builder.getCommentsImplementationType() == CommentsImplementationType.CUSTOM_MAP_BACKED) {
      try (InputStream is = commentsPart.getInputStream()) {
        MapBackedCommentsTable ct = new MapBackedCommentsTable(
                builder.fullFormatRichText());
        try {
          ct.readFrom(is);
        } catch (IOException | RuntimeException e) {
          ct.close();
          throw e;
        }
        return ct;
      }
    } else if (strictOoxmlChecksNeeded) {
      return OoxmlStrictHelper.getCommentsTable(builder, commentsPart);
    } else {
      return new CommentsTable(commentsPart);
    }
  }

  /**
   * Returns the shapes associated with this sheet,
   * an empty list or null if there is an exception
   */
  private static List<XSSFShape> getShapes(final PackagePart sheetPkg, final boolean strictOoxmlChecksNeeded) {
    final List<XSSFShape> shapes = new LinkedList<>();
    try {
      PackageRelationshipCollection drawingsList = sheetPkg.getRelationshipsByType(XSSFRelation.DRAWINGS.getRelation());
      if (drawingsList.size() == 0 && strictOoxmlChecksNeeded) {
        drawingsList = sheetPkg.getRelationshipsByType(PURL_DRAWING_RELATIONSHIP_URL);
      }
      for (int i = 0; i < drawingsList.size(); i++) {
        PackageRelationship drawings = drawingsList.getRelationship(i);
        PackagePartName drawingsName = PackagingURIHelper.createPartName(drawings.getTargetURI());
        PackagePart drawingsPart = sheetPkg.getPackage().getPart(drawingsName);
        if (drawingsPart == null) {
          //parts can go missing; Excel ignores them silently -- TIKA-2134
          LOGGER.warn("Missing drawing: {}. Skipping it.", drawingsName);
          continue;
        }
        XSSFDrawing drawing = new XSSFDrawing(drawingsPart);
        shapes.addAll(drawing.getShapes());
      }
    } catch (XmlException | InvalidFormatException | IOException e) {
      LOGGER.warn("issue getting shapes", e);
      return null;
    }
    return shapes;
  }

}