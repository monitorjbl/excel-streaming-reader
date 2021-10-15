package com.github.pjfanning.xlsx;

import com.github.pjfanning.xlsx.exceptions.*;
import com.github.pjfanning.xlsx.impl.*;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

/**
 * Streaming Excel workbook implementation. Most advanced features of POI are not supported.
 * Use this only if your application can handle iterating through an entire workbook, row by
 * row.
 */
public class StreamingReader implements AutoCloseable {
  private final StreamingWorkbookReader workbook;

  public StreamingReader(StreamingWorkbookReader workbook) {
    this.workbook = workbook;
  }

  /**
   * Closes the streaming resource, attempting to clean up any temporary files created.
   *
   * @throws CloseException if there is an issue closing the stream
   */
  @Override
  public void close() throws IOException {
    workbook.close();
  }

  public static Builder builder() {
    return new Builder();
  }

  public static class Builder {
    private int rowCacheSize = 10;
    private int bufferSize = 1024;
    private boolean avoidTempFiles = false;
    private boolean useSstTempFile = false;
    private boolean encryptSstTempFile = false;
    private boolean useCommentsTempFile = false;
    private boolean encryptCommentsTempFile = false;
    private boolean adjustLegacyComments = false;
    private boolean readComments = false;
    private boolean readCoreProperties = false;
    private boolean readHyperlinks = false;
    private boolean readShapes = false;
    private boolean fullFormatRichText = false;
    private String password;
    private boolean convertFromOoXmlStrict;

    public int getRowCacheSize() {
      return rowCacheSize;
    }

    public int getBufferSize() {
      return bufferSize;
    }

    /**
     * @return The password to use to unlock this workbook
     */
    public String getPassword() {
      return password;
    }

    /**
     * @return Whether to avoid temp files when reading input streams.
     */
    public boolean avoidTempFiles() {
      return avoidTempFiles;
    }

    /**
     * @return Whether to use a temp file for the Shared Strings data. If false, no
     * temp file will be used and the entire table will be loaded into memory.
     */
    public boolean useSstTempFile() {
      return useSstTempFile;
    }

    /**
     * @return Whether to encrypt the temp file for the Shared Strings data. Only applies if <code>useSstTempFile()</code>
     * is true.
     */
    public boolean encryptSstTempFile() {
      return encryptSstTempFile;
    }

    /**
     * @return Whether to adjust comments to remove boiler-plate text (related to threaded comments).
     * See https://github.com/pjfanning/excel-streaming-reader/issues/57
     */
    public boolean adjustLegacyComments() {
      return adjustLegacyComments;
    }

    /**
     * @return Whether to read the comments associated with sheets and cells.
     */
    public boolean readComments() {
      return readComments;
    }

    /**
     * @return Whether to use a temp file for the Comments data. If false, no
     * temp file will be used and the entire table will be loaded into memory.
     */
    public boolean useCommentsTempFile() {
      return useCommentsTempFile;
    }

    /**
     * @return Whether to encrypt the temp file for the Comments data. Only applies if <code>useCommentsTempFile()</code>
     * is true.
     */
    public boolean encryptCommentsTempFile() {
      return encryptCommentsTempFile;
    }

    /**
     * @return Whether to read the core document properties.
     */
    public boolean readCoreProperties() {
      return readCoreProperties;
    }

    /**
     * @return Whether to read the hyperlink data that appear in sheets.
     */
    public boolean readHyperlinks() {
      return readHyperlinks;
    }

    /**
     * @return Whether to read the shapes (associated with pictures that appear in sheets).
     */
    public boolean readShapes() {
      return readShapes;
    }

    /**
     * Whether to parse the full formatting data for rich text shared strings and comments.
     * This only has an effect if temp file SST and/or Comments Table support is enabled. The default is false.
     * When you don't use temp file support, full formatting data is returned for the rich text anyway.
     * @return Whether to parse the full formatting data for rich text shared strings and comments.
     * @see #useSstTempFile()
     * @see #useCommentsTempFile()
     */
    public boolean fullFormatRichText() {
      return fullFormatRichText;
    }

    /**
     * @return Whether to convert the input from Strict OOXML (prevent "Strict OOXML isn't currently supported")
     * @deprecated this is no longer needed, OOXML strict format is handled automatically
     */
    public boolean convertFromOoXmlStrict() {
      return convertFromOoXmlStrict;
    }

    /**
     * The number of rows to keep in memory at any given point.
     * <p>
     * Defaults to 10
     * </p>
     *
     * @param rowCacheSize number of rows
     * @return reference to current {@code Builder}
     */
    public Builder rowCacheSize(int rowCacheSize) {
      this.rowCacheSize = rowCacheSize;
      return this;
    }

    /**
     * The number of bytes to read into memory from the input
     * resource.
     * <p>
     * Defaults to 1024
     * </p>
     *
     * @param bufferSize buffer size in bytes
     * @return reference to current {@code Builder}
     */
    public Builder bufferSize(int bufferSize) {
      this.bufferSize = bufferSize;
      return this;
    }

    /**
     * For password protected files, specify password to open file.
     * If the password is incorrect a {@code ReadException} is thrown on
     * {@code read}.
     * <p>NULL indicates that no password should be used, this is the
     * default value.</p>
     *
     * @param password to use when opening file
     * @return reference to current {@code Builder}
     */
    public Builder password(String password) {
      this.password = password;
      return this;
    }

    /**
     * Convert the file from Strict OOXML to regular XLSX.
     * Strict OOXML is not supported by POI. This is an experimental feature.
     *
     * @param convertFromOoXmlStrict whether to convert from OOXML
     * @return reference to current {@code Builder}
     * @deprecated this is no longer needed, OOXML strict format is handled automatically
     */
    @Deprecated
    public Builder convertFromOoXmlStrict(boolean convertFromOoXmlStrict) {
      this.convertFromOoXmlStrict = convertFromOoXmlStrict;
      return this;
    }

    /**
     * Enables a mode where the code tries to avoid creating temp files. This is independent of
     * {@code #setUseSstTempFile} and {@code #setUseCommentsTempFile}.
     * <p>
     * By default, temp files are used to avoid holding onto too much data in memory.
     *
     * @param avoidTempFiles whether to avoid using temp files when reading from input streams
     * @return reference to current {@code Builder}
     */
    public Builder setAvoidTempFiles(boolean avoidTempFiles) {
      this.avoidTempFiles = avoidTempFiles;
      return this;
    }

    /**
     * Enables use of Shared Strings Table temp file. This option exists to accommodate
     * extremely large workbooks with millions of unique strings. Normally, the SST is entirely
     * loaded into memory, but with large workbooks with high cardinality (i.e., very few
     * duplicate values) the SST may not fit entirely into memory.
     * <p>
     * By default, the entire SST *will* be loaded into memory. <strong>However</strong>,
     * enabling this option at all will have some noticeable performance degradation as you are
     * trading memory for disk space.
     * <p>
     * If you enable this feature, you also want to enable <code>fullFormatRichText</code>.
     *
     * @param useSstTempFile whether to use a temp file to store the Shared Strings Table data
     * @return reference to current {@code Builder}
     * @see #setEncryptSstTempFile(boolean)
     * @see #setFullFormatRichText(boolean) 
     */
    public Builder setUseSstTempFile(boolean useSstTempFile) {
      this.useSstTempFile = useSstTempFile;
      return this;
    }

    /**
     * Enables use of encryption in Shared Strings Table temp file. This only applies if <code>setUseSstTempFile</code>
     * is set to true.
     * <p>
     * By default, the temp file is not encrypted. <strong>However</strong>,
     * enabling this option could slow down the processing of Shared Strings data.
     *
     * @param encryptSstTempFile whether to encrypt the temp file used to store the Shared Strings Table data
     * @return reference to current {@code Builder}
     * @see #setUseSstTempFile(boolean) 
     */
    public Builder setEncryptSstTempFile(boolean encryptSstTempFile) {
      this.encryptSstTempFile = encryptSstTempFile;
      return this;
    }

    /**
     * Enables use of Comments temp file. This option exists to accommodate
     * workbooks with lots of comments. Normally, the Comments are all
     * loaded into memory.
     * <p>
     * By default, all the Comments data *will* be loaded into memory. <strong>However</strong>,
     * enabling this option at all will have some noticeable performance degradation as you are
     * trading memory for disk space.
     * <p>
     * If you enable this feature, you also want to enable <code>fullFormatRichText</code>.
     *
     * @param useCommentsTempFile whether to use a temp file to store the Comments data
     * @return reference to current {@code Builder}
     * @see #setReadComments(boolean)
     * @see #setEncryptCommentsTempFile(boolean)
     * @see #setFullFormatRichText(boolean) 
     */
    public Builder setUseCommentsTempFile(boolean useCommentsTempFile) {
      this.useCommentsTempFile = useCommentsTempFile;
      return this;
    }

    /**
     * Enables use of encryption in the Comments temp file. This only applies if <code>setUseCommentsTempFile</code>
     * is set to true.
     * <p>
     * By default, the temp file is not encrypted. <strong>However</strong>,
     * enabling this option could slow down the processing of Comments data.
     *
     * @param encryptCommentsTempFile whether to encrypt the temp file used to store the Comments data
     * @return reference to current {@code Builder}
     * @see #setReadComments(boolean)
     * @see #setUseCommentsTempFile(boolean)
     */
    public Builder setEncryptCommentsTempFile(boolean encryptCommentsTempFile) {
      this.encryptCommentsTempFile = encryptCommentsTempFile;
      return this;
    }

    /**
     * Enables the reading of the comments.
     *
     * @param readComments whether to read the comments associated with sheets and cells
     * @return reference to current {@code Builder}
     */
    public Builder setReadComments(boolean readComments) {
      this.readComments = readComments;
      return this;
    }

    /**
     * Enables adjustments to comments to remove boiler-plate text (related to threaded comments).
     * See https://github.com/pjfanning/excel-streaming-reader/issues/57.
     *
     * @param adjustLegacyComments whether to adjust legacy comments to remove boiler-plate comments
     * @return reference to current {@code Builder}
     */
    public Builder setAdjustLegacyComments(boolean adjustLegacyComments) {
      this.adjustLegacyComments = adjustLegacyComments;
      return this;
    }

    /**
     * Enables the reading of the core document properties.
     *
     * @param readCoreProperties whether to read the core document properties
     * @return reference to current {@code Builder}
     */
    public Builder setReadCoreProperties(boolean readCoreProperties) {
      this.readCoreProperties = readCoreProperties;
      return this;
    }

    /**
     * Enables the reading of hyperlink data associated wit sheets).
     *
     * @param readHyperlinks whether to read hyperlink data (associated with sheets)
     * @return reference to current {@code Builder}
     */
    public Builder setReadHyperlinks(boolean readHyperlinks) {
      this.readHyperlinks = readHyperlinks;
      return this;
    }

    /**
     * Enables the reading of shape data.
     *
     * @param readShapes whether to read shapes (associated with pictures that appear in sheets)
     * @return reference to current {@code Builder}
     */
    public Builder setReadShapes(boolean readShapes) {
      this.readShapes = readShapes;
      return this;
    }

    /**
     * Whether to parse the full formatting data for rich text shared strings and comments.
     * This only has an effect if temp file SST and/or Comments Table support is enabled. The default is false.
     * When you don't use temp file support, full formatting data is returned for the rich text anyway.
     * @param fullFormatRichText Whether to parse the full formatting data for rich text shared strings and comments.
     * @return reference to current {@code Builder}
     * @see #setUseSstTempFile(boolean)
     * @see #setUseCommentsTempFile(boolean)
     */
    public Builder setFullFormatRichText(boolean fullFormatRichText) {
      this.fullFormatRichText = fullFormatRichText;
      return this;
    }

    /**
     * Reads a given {@code InputStream} and returns a new
     * instance of {@code Workbook}. Due to Apache POI
     * limitations, a temporary file must be written in order
     * to create a streaming iterator. This process will use
     * the same buffer size as specified in {@link #bufferSize(int)}.
     *
     * @param is input stream to read in
     * @return A {@link Workbook} that can be read from
     * @throws ReadException if there is an issue reading the stream
     */
    public Workbook open(InputStream is) {
      StreamingWorkbookReader workbook = new StreamingWorkbookReader(this);
      workbook.init(is);
      return new StreamingWorkbook(workbook);
    }

    /**
     * Reads a given {@code File} and returns a new instance
     * of {@code Workbook}.
     *
     * @param file file to read in
     * @return built streaming reader instance
     * @throws OpenException if there is an issue opening the file
     * @throws ReadException if there is an issue reading the file
     */
    public Workbook open(File file) {
      StreamingWorkbookReader workbook = new StreamingWorkbookReader(this);
      workbook.init(file);
      return new StreamingWorkbook(workbook);
    }
  }
}
