package com.monitorjbl.xlsx.impl;

import org.apache.poi.ss.usermodel.CellType;

import java.util.stream.Stream;


enum CellTypeHelper {

	_NONE("none", CellType._NONE),
	NUMERIC("numeric", CellType.NUMERIC, CellTypeConstants.NUMERIC),
	STRING("text", CellType.STRING, CellTypeConstants.STRING, CellTypeConstants.STR, CellTypeConstants.INLINE_STR),
	FORMULA("formula", CellType.FORMULA),
	BLANK("blank", CellType.BLANK),
	BOOLEAN("boolean", CellType.BOOLEAN, CellTypeConstants.BOOLEAN),
	ERROR("error", CellType.ERROR, CellTypeConstants.ERROR);


	private final String val;
	private final CellType type;
	private final String[] shortHand;

	CellTypeHelper(String val, CellType type, String... shortHand) {
		this.val = val;
		this.type = type;
		this.shortHand = shortHand;
	}

	static String getValueFromCellType(final CellType cellType) {
		for (CellTypeHelper cth : CellTypeHelper.values()) {
			if (cth.type.equals(cellType)) {
				return cth.val;
			}
		}
		return "#unknown cell type (" + cellType + ")#";
	}

	static CellType getCellTypeFromShortHand(final String shortHand) {
		for (CellTypeHelper cth : CellTypeHelper.values()) {
			if (Stream.of(cth.shortHand).anyMatch(s -> s.equalsIgnoreCase(shortHand))) {
				return cth.type;
			}
		}
		throw new UnsupportedOperationException("Unsupported cell type '" + shortHand + "'");
	}
}
