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

import com.github.pjfanning.poi.xssf.streaming.TempFileCommentsTable;
import com.github.pjfanning.xlsx.StreamingReader;
import com.github.pjfanning.xlsx.impl.StreamingWorkbookReader;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.util.Internal;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.*;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.xmlbeans.XmlException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

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

  private final boolean strictOoxmlChecksNeeded;
  private final StreamingWorkbookReader streamingWorkbookReader;

  /**
   * Creates a new XSSFReader, for the given package
   */
  @Internal
  public OoxmlReader(StreamingWorkbookReader streamingWorkbookReader,
                     OPCPackage pkg, boolean strictOoxmlChecksNeeded) throws IOException, OpenXML4JException {
    super(pkg, true);
    this.streamingWorkbookReader = streamingWorkbookReader;
    this.strictOoxmlChecksNeeded = strictOoxmlChecksNeeded;

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
    return parts.isEmpty() ? null :
            builder.useSstReadOnly() ? new ReadOnlySharedStringsTable(parts.get(0)) :
              new SharedStringsTable(parts.get(0));
  }

  /**
   * Opens up the Styles Table, parses it, and
   * returns a handy object for working with cell styles
   */
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
  public OoxmlSheetIterator getSheetsData() throws IOException {
    return new OoxmlSheetIterator(workbookPart);
  }

  /**
   * Iterator over sheet data.
   */
  public class OoxmlSheetIterator extends SheetIterator {

    /**
     * Construct a new SheetIterator
     *
     * @param wb package part holding workbook.xml
     */
    OoxmlSheetIterator(PackagePart wb) throws IOException {
      super(wb);
    }

    /**
     * Gets string representations of relationships
     * that are sheet-like.  Added to allow subclassing
     * by XSSFBReader.  This is used to decide what
     * relationships to load into the sheetRefs
     *
     * @return all relationships that are sheet-like
     */
    protected Set<String> getSheetRelationships() {
      return OVERRIDE_WORKSHEET_RELS;
    }

    /**
     * Returns the comments associated with this sheet,
     * or null if there aren't any
     */
    public Comments getSheetComments(StreamingReader.Builder builder) {
      PackagePart sheetPkg = getSheetPart();

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
          return parseComments(builder, commentsPart);
        }
      } catch (InvalidFormatException|IOException|XMLStreamException e) {
        LOGGER.warn("issue getting sheet comments", e);
        return null;
      }
      return null;
    }

    private Comments parseComments(StreamingReader.Builder builder, PackagePart commentsPart) throws IOException, XMLStreamException, InvalidFormatException {
      if (builder.useCommentsTempFile()) {
        try (InputStream is = commentsPart.getInputStream()) {
          TempFileCommentsTable ct = new TempFileCommentsTable(
                  builder.encryptCommentsTempFile(),
                  builder.fullFormatRichText());
          try {
            ct.readFrom(is);
          } catch (IOException|RuntimeException e) {
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
    public List<XSSFShape> getShapes() {
      PackagePart sheetPkg = getSheetPart();
      List<XSSFShape> shapes = new LinkedList<>();
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
      } catch (XmlException|InvalidFormatException|IOException e) {
        LOGGER.warn("issue getting shapes", e);
        return null;
      }
      return shapes;
    }
  }
}
