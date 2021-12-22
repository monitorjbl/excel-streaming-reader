package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.StreamingReader;
import com.github.pjfanning.xlsx.impl.ooxml.OoxmlStrictHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.model.ThemesTable;
import org.junit.Test;

import java.io.File;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

public class OoxmlStrictHelperTest {
  @Test
  public void testThemes() throws Exception {
    StreamingReader.Builder builder1 = StreamingReader.builder().setAvoidTempFiles(false);
    StreamingReader.Builder builder2 = StreamingReader.builder().setAvoidTempFiles(true);
    for(StreamingReader.Builder builder : new StreamingReader.Builder[]{builder1, builder2}) {
      try (OPCPackage pkg = OPCPackage.open(new File("src/test/resources/sample.strict.xlsx"), PackageAccess.READ)) {
        ThemesTable themes = OoxmlStrictHelper.getThemesTable(builder, pkg);
        assertNotNull(themes.getThemeColor(ThemesTable.ThemeElement.DK1.idx));
      }
    }
  }

  @Test
  public void testStyles() throws Exception {
    StreamingReader.Builder builder1 = StreamingReader.builder().setAvoidTempFiles(false);
    StreamingReader.Builder builder2 = StreamingReader.builder().setAvoidTempFiles(true);
    for(StreamingReader.Builder builder : new StreamingReader.Builder[]{builder1, builder2}) {
      try(OPCPackage pkg = OPCPackage.open(new File("src/test/resources/sample.strict.xlsx"), PackageAccess.READ)) {
        StylesTable styles = OoxmlStrictHelper.getStylesTable(builder, pkg);
        ThemesTable themes = OoxmlStrictHelper.getThemesTable(builder, pkg);
        styles.setTheme(themes);
        assertEquals("has right borders", 1, styles.getBorders().size());
        assertEquals("has right fonts", 11, styles.getFonts().size());
        assertEquals("has right cell styles", 3, styles.getNumCellStyles());
      }
    }
  }

  @Test
  public void testSharedStrings() throws Exception {
    StreamingReader.Builder builder1 = StreamingReader.builder().setAvoidTempFiles(false);
    StreamingReader.Builder builder2 = StreamingReader.builder().setAvoidTempFiles(true);
    StreamingReader.Builder builder3 = StreamingReader.builder()
            .setUseSstReadOnly(true)
            .setAvoidTempFiles(false);
    StreamingReader.Builder builder4 = StreamingReader.builder()
            .setUseSstReadOnly(true)
            .setAvoidTempFiles(true);
    for (StreamingReader.Builder builder : new StreamingReader.Builder[]{builder1, builder2, builder3, builder4}) {
      try (OPCPackage pkg = OPCPackage.open(new File("src/test/resources/sample.strict.xlsx"), PackageAccess.READ)) {
        SharedStrings sst = OoxmlStrictHelper.getSharedStringsTable(builder, pkg);
        assertEquals("has right count", 15, sst.getUniqueCount());
        assertEquals("has right count", 19, sst.getCount());
        assertEquals("ipsum", sst.getItemAt(1).getString());
      }
    }
  }
}
