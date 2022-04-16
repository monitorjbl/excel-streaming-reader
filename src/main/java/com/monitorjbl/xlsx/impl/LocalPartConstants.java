package com.monitorjbl.xlsx.impl;


import javax.xml.namespace.QName;

public enum LocalPartConstants {
	LOCAL_PART_C("c"),
	LOCAL_PART_F("f"),
	LOCAL_PART_V("v"),
	LOCAL_PART_T("t"),
	LOCAL_PART_R("r"),
	LOCAL_PART_S("s"),
	LOCAL_PART_ROW("row"),
	LOCAL_PART_COL("col"),
	LOCAL_PART_DIMENSION("dimension"),
	LOCAL_PART_HIDDEN("hidden"),
	LOCAL_PART_MIN("min"),
	LOCAL_PART_MAX("max"),
	LOCAL_PART_REF("ref"),
	LOCAL_PART_PHONETIC_PR("phoneticPr"),
	LOCAL_PART_RPR("rPr"),
	LOCAL_PART_UNKNOWN("unknown");

	private final String constant;
	private final QName qName;

	LocalPartConstants(final String val) {
		this.constant = val;
		this.qName = new QName(val);
	}

	public boolean isConstantEquals(String otherConstant) {
		return constant.equals(otherConstant);
	}


	public QName getQname() {
		return qName;
	}

	public static LocalPartConstants parseValue(String value) {
		for (LocalPartConstants val : LocalPartConstants.values()) {
			if (val.constant.equalsIgnoreCase(value.trim())) {
				return val;
			}
		}
		return LOCAL_PART_UNKNOWN;
	}
}