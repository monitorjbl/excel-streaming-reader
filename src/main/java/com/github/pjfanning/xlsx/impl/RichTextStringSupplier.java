package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.usermodel.RichTextString;

class RichTextStringSupplier implements Supplier {
  private final RichTextString val;

  RichTextStringSupplier(RichTextString val) {
    this.val = val;
  }

  @Override
  public Object getContent() {
        return val;
    }
}
