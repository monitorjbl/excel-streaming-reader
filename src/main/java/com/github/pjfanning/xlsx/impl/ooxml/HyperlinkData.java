package com.github.pjfanning.xlsx.impl.ooxml;

import java.util.Objects;

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

  @Override
  public boolean equals(Object o) {
    if (this == o) return true;
    if (!(o instanceof HyperlinkData)) return false;
    HyperlinkData that = (HyperlinkData) o;
    return Objects.equals(id, that.id)
        && Objects.equals(ref, that.ref)
        && Objects.equals(location, that.location)
        && Objects.equals(display, that.display)
        && Objects.equals(tooltip, that.tooltip);
  }

  @Override
  public int hashCode() {
    return Objects.hash(
        id,
        ref,
        location,
        display,
        tooltip);
  }
}
