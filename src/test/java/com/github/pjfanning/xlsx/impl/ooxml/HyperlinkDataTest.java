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
  public void testEqualsWithNullLocation() {
    HyperlinkData hd1 = new HyperlinkData("id1", "ref1", null, "disp1", "tip1");
    HyperlinkData hd2 = new HyperlinkData("id1", "ref1", null, "disp1", "tip1");
    assertEquals(hd1, hd2);
    assertEquals(hd1.hashCode(), hd2.hashCode());
  }

  @Test
  public void testEqualsWithNullRef() {
    HyperlinkData hd1 = new HyperlinkData("id1", null, "locn1", "disp1", "tip1");
    HyperlinkData hd2 = new HyperlinkData("id1", null, "locn1", "disp1", "tip1");
    assertEquals(hd1, hd2);
    assertEquals(hd1.hashCode(), hd2.hashCode());
  }

  @Test
  public void testEqualsWithNullId() {
    HyperlinkData hd1 = new HyperlinkData(null, "ref1", "locn1", "disp1", "tip1");
    HyperlinkData hd2 = new HyperlinkData(null, "ref1", "locn1", "disp1", "tip1");
    assertEquals(hd1, hd2);
    assertEquals(hd1.hashCode(), hd2.hashCode());
  }

  @Test
  public void testNotEquals() {
    HyperlinkData hd1 = new HyperlinkData("id1", "ref1", "locn1", "disp1", "tip1");
    HyperlinkData hd2 = new HyperlinkData("id1", "ref1", "locn1", "disp1", null);
    HyperlinkData hd3 = new HyperlinkData("id1", "ref1", "locn1", "disp1", "tip2");
    HyperlinkData hd4 = new HyperlinkData("id1", "ref1", "locn1", "disp2", "tip1");
    HyperlinkData hd5 = new HyperlinkData("id1", "ref1", "locn1", null, "tip1");
    HyperlinkData hd6 = new HyperlinkData("id1", "ref1", "locn2", "disp1", "tip1");
    HyperlinkData hd7 = new HyperlinkData("id1", "ref1", null, "disp1", "tip1");
    HyperlinkData hd8 = new HyperlinkData("id1", "ref2", "locn1", "disp1", "tip1");
    HyperlinkData hd9 = new HyperlinkData("id1", null, "locn1", "disp1", "tip1");
    HyperlinkData hd10 = new HyperlinkData("id2", "ref1", "locn1", "disp1", "tip1");
    HyperlinkData hd11 = new HyperlinkData(null, "ref1", "locn1", "disp1", "tip1");
    assertNotEquals(hd1, hd2);
    assertNotEquals(hd1, hd3);
    assertNotEquals(hd1, hd4);
    assertNotEquals(hd1, hd5);
    assertNotEquals(hd1, hd6);
    assertNotEquals(hd1, hd7);
    assertNotEquals(hd1, hd8);
    assertNotEquals(hd1, hd9);
    assertNotEquals(hd1, hd10);
    assertNotEquals(hd1, hd11);
    assertNotEquals(hd1.hashCode(), hd2.hashCode());
    assertNotEquals(hd1.hashCode(), hd3.hashCode());
    assertNotEquals(hd1.hashCode(), hd4.hashCode());
    assertNotEquals(hd1.hashCode(), hd5.hashCode());
    assertNotEquals(hd1.hashCode(), hd6.hashCode());
    assertNotEquals(hd1.hashCode(), hd7.hashCode());
    assertNotEquals(hd1.hashCode(), hd8.hashCode());
    assertNotEquals(hd1.hashCode(), hd9.hashCode());
    assertNotEquals(hd1.hashCode(), hd10.hashCode());
    assertNotEquals(hd1.hashCode(), hd11.hashCode());
  }

}
