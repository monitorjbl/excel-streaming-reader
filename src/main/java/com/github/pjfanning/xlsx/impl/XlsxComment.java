package com.github.pjfanning.xlsx.impl;

import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;

public class XlsxComment extends XSSFComment {
  XlsxComment(CommentsTable comments, CTComment comment) {
    super(comments, comment, null);
  }

  @Override
  public XSSFRichTextString getString() {
    XSSFRichTextString rts = super.getString();
    String text = rts.getString();
    if(rts.getString().contains("Your version of Excel allows you to read this threaded comment")) {
      String splitText = "Comment:";
      int pos = text.lastIndexOf(splitText);
      if (pos != -1) {
        return new XSSFRichTextString(ltrim(text.substring(pos + splitText.length())));
      }
    }
    return rts;
  }

  private String ltrim(String s) {
    int i = 0;
    while (i < s.length() && Character.isWhitespace(s.charAt(i))) {
      i++;
    }
    return s.substring(i);
  }
}
