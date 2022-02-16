package com.github.pjfanning.xlsx;

public enum CommentsImplementationType {
  /**
   * The default <code>CommentsTable</code> in POI
   */
  POI_DEFAULT,
  /**
   * The temp file backed <code>CommentsTable</code> in <code>poi-shared-strings</code>.
   * Saves on memory but still has good performance, especially if full-format text is set to false.
   * @see StreamingReader.Builder#setFullFormatRichText(boolean)
   * @see StreamingReader.Builder#setEncryptCommentsTempFile(boolean)
   */
  TEMP_FILE_BACKED,
  /**
   * The concurrent map backed <code>CommentsTable</code> in <code>poi-shared-strings</code>.
   * More performant if full-format text is set to false.
   * @see StreamingReader.Builder#setFullFormatRichText(boolean)
   */
  CUSTOM_MAP_BACKED
}
