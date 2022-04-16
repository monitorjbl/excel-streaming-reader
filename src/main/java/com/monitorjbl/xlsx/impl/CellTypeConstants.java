package com.monitorjbl.xlsx.impl;

final class CellTypeConstants {

	final static String NUMERIC = "n";
	final static String STR = "str";
	final static String STRING = "s";
	final static String INLINE_STR = "inlineStr";
	final static String BOOLEAN = "b";
	final static String ERROR = "e";


	private CellTypeConstants() {
		throw new RuntimeException("It is not good practice to instantiate constants classes.");
	}
}
