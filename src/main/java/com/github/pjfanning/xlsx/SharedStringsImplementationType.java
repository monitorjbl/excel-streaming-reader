package com.github.pjfanning.xlsx;

public enum SharedStringsImplementationType {
  /**
   * The default <code>SharedStringsTable</code> in POI
   */
  POI_DEFAULT,
  /**
   * The read-only <code>SharedStringsTable</code> in POI (more efficient than the default POI implementation).
   * This is the default in <code>excel-streaming-reader</code> since v3.5.0.
   */
  POI_READ_ONLY,
  /**
   * The temp file backed <code>SharedStringsTable</code> in <code>poi-shared-strings</code>.
   * Saves on memory but still has good performance, especially if full-format text is set to false.
   * @see StreamingReader.Builder#setFullFormatRichText(boolean)
   * @see StreamingReader.Builder#setEncryptSstTempFile(boolean)
   */
  TEMP_FILE_BACKED,
  /**
   * The concurrent map backed <code>SharedStringsTable</code> in <code>poi-shared-strings</code>.
   * More performant if full-format text is set to false.
   * @see StreamingReader.Builder#setFullFormatRichText(boolean)
   */
  CUSTOM_MAP_BACKED
}
