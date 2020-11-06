package com.github.pjfanning.xlsx;

import org.junit.Test;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

public class XmlUtilsTest {
  @Test
  public void testEvaluateBoolean() {
    assertTrue(XmlUtils.evaluateBoolean("1"));
    assertTrue(XmlUtils.evaluateBoolean("true"));
    assertTrue(XmlUtils.evaluateBoolean("TRUE"));
    assertFalse(XmlUtils.evaluateBoolean("0"));
    assertFalse(XmlUtils.evaluateBoolean("false"));
    assertFalse(XmlUtils.evaluateBoolean("FALSE"));
    assertFalse(XmlUtils.evaluateBoolean(""));
    assertFalse(XmlUtils.evaluateBoolean("fake"));
  }
}
