package com.github.pjfanning.xlsx;

import com.github.pjfanning.xlsx.exceptions.*;
import com.github.pjfanning.xlsx.impl.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.Beta;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

/**
 * Streaming Excel workbook implementation. Most advanced features of POI are not supported.
 * Use this only if your application can handle iterating through an entire workbook, row by
 * row.
 */
public class StreamingReader implements AutoCloseable {
  private static final Logger log = LoggerFactory.getLogger(StreamingReader.class);

  private File tmp;
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
    try {
      workbook.close();
    } finally {
      if(tmp != null) {
        if (log.isDebugEnabled()) {
          log.debug("Deleting tmp file [" + tmp.getAbsolutePath() + "]");
        }
        tmp.delete();
      }
    }
  }

  public static Builder builder() {
    return new Builder();
  }

  public static class Builder {
    private int rowCacheSize = 10;
    private int bufferSize = 1024;
    private boolean useSstTempFile = false;
    private boolean encryptSstTempFile = false;
    private boolean readCoreProperties = false;
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
     * @return Whether to read the core document properties.
     */
    public boolean readCoreProperties() {
      return readCoreProperties;
    }

    /**
     * @return Whether to convert the input from Strict OOXML (prevent "Strict OOXML isn't currently supported")
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
     * For password protected files specify password to open file.
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
     */
    @Beta
    public Builder convertFromOoXmlStrict(boolean convertFromOoXmlStrict) {
      this.convertFromOoXmlStrict = convertFromOoXmlStrict;
      return this;
    }

    /**
     * Enables use of Shared Strings Table temp file. This option exists to accommodate
     * extremely large workbooks with millions of unique strings. Normally the SST is entirely
     * loaded into memory, but with large workbooks with high cardinality (i.e., very few
     * duplicate values) the SST may not fit entirely into memory.
     * <p>
     * By default, the entire SST *will* be loaded into memory. <strong>However</strong>,
     * enabling this option at all will have some noticeable performance degradation as you are
     * trading memory for disk space.
     *
     * @param useSstTempFile whether to use a temp file to store the Shared Strings Table data
     * @return reference to current {@code Builder}
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
     */
    public Builder setEncryptSstTempFile(boolean encryptSstTempFile) {
      this.encryptSstTempFile = encryptSstTempFile;
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
