package com.github.pjfanning.xlsx.impl.ooxml;

import org.junit.Test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotEquals;

public class HyperlinkDataTest {
  @Test
  public void testEquals() {
    HyperlinkData hd1 = new HyperlinkData("id1", "ref1", "locn1", "disp1", "tip1");
    HyperlinkData hd2 = new HyperlinkData("id1", "ref1", "locn1", "disp1", "tip1");
    assertEquals(hd1, hd2);
    assertEquals(hd1.hashCode(), hd2.hashCode());
  }

  @Test
  public void testEqualsWithNullTooltip() {
    HyperlinkData hd1 = new HyperlinkData("id1", "ref1", "locn1", "disp1", null);
    HyperlinkData hd2 = new HyperlinkData("id1", "ref1", "locn1", "disp1", null);
    assertEquals(hd1, hd2);
    assertEquals(hd1.hashCode(), hd2.hashCode());
  }

  @Test
  public void testEqualsWithNullDisplay() {
    HyperlinkData hd1 = new HyperlinkData("id1", "ref1", "locn1", null, "tip1");
    HyperlinkData hd2 = new HyperlinkData("id1", "ref1", "locn1", null, "tip1");
    assertEquals(hd1, hd2);
    assertEquals(hd1.hashCode(), hd2.hashCode());
  }

  @Test
  public void testNotEquals() {
    HyperlinkData hd1 = new HyperlinkData("id1", "ref1", "locn1", "disp1", "tip1");
    HyperlinkData hd2 = new HyperlinkData("id1", "ref1", "locn1", "disp1", null);
    HyperlinkData hd3 = new HyperlinkData("id1", "ref1", "locn1", "disp1", "tip2");
    assertNotEquals(hd1, hd2);
    assertNotEquals(hd1, hd3);
    assertNotEquals(hd1.hashCode(), hd2.hashCode());
    assertNotEquals(hd1.hashCode(), hd3.hashCode());
  }

}
