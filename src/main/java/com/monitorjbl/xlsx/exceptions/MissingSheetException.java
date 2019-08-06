package com.monitorjbl.xlsx.exceptions;

import com.monitorjbl.xlsx.StreamingReader;

/**
 * Will be thrown by {@link StreamingReader.Builder#read(java.io.File)}, if a sheet is not found.
 *
 * @deprecated The POI API {@link org.apache.poi.ss.usermodel.Workbook#getSheet} is designed to return {@code null},
 * if a sheet is not found. This class will be removed in a future release.
 */
@Deprecated
public class MissingSheetException extends RuntimeException {

  public MissingSheetException() {
    super();
  }

  public MissingSheetException(String msg) {
    super(msg);
  }

  public MissingSheetException(Exception e) {
    super(e);
  }

  public MissingSheetException(String msg, Exception e) {
    super(msg, e);
  }
}
