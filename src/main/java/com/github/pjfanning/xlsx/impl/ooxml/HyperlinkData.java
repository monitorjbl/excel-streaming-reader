package com.github.pjfanning.xlsx.impl.ooxml;

public class HyperlinkData {

  private final String id;
  private final String ref;
  private final String location;
  private final String display;
  private final String tooltip;

  public HyperlinkData(String id, String ref, String location, String display, String tooltip) {
    this.id = id;
    this.ref = ref;
    this.location = location;
    this.display = display;
    this.tooltip = tooltip;
  }

  public String getId() {
    return id;
  }

  public String getRef() {
    return ref;
  }

  public String getLocation() {
    return location;
  }

  public String getDisplay() {
    return display;
  }

  public String getTooltip() {
    return tooltip;
  }
}
