package com.monitorjbl.xlsx.notsupportedoperations;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.CellReferenceType;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public interface WorkbookAdapter extends Workbook {

	/* Not supported */

	/**
	 * Not supported
	 */
	@Override
	default int getActiveSheetIndex() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setActiveSheet(int sheetIndex) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default int getFirstVisibleTab() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setFirstVisibleTab(int sheetIndex) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setSheetOrder(String sheetname, int pos) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setSelectedTab(int index) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setSheetName(int sheet, String name) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Sheet createSheet() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Sheet createSheet(String sheetname) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Sheet cloneSheet(int sheetNum) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void removeSheetAt(int index) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Font createFont() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Font findFont(boolean b, short i, short i1, String s, boolean b1, boolean b2, short i2, byte b3) {
		throw new UnsupportedOperationException();
	}

	@Override
	default int getNumberOfFonts() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default int getNumberOfFontsAsInt() { throw new UnsupportedOperationException(); }

	/**
	 * Not supported
	 */
	@Override
	default Font getFontAt(int i) { throw new UnsupportedOperationException(); }

	/**
	 * Not supported
	 */
	@Override
	default CellStyle createCellStyle() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default int getNumCellStyles() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default CellStyle getCellStyleAt(int i) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void write(OutputStream stream) throws IOException {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default int getNumberOfNames() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Name getName(String name) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default List<? extends Name> getNames(String s) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default List<? extends Name> getAllNames() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Name createName() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void removeName(Name name) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default int linkExternalWorkbook(String name, Workbook workbook) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setPrintArea(int sheetIndex, String reference) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default String getPrintArea(int sheetIndex) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void removePrintArea(int sheetIndex) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Row.MissingCellPolicy getMissingCellPolicy() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setMissingCellPolicy(Row.MissingCellPolicy missingCellPolicy) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default DataFormat createDataFormat() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default int addPicture(byte[] pictureData, int format) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default List<? extends PictureData> getAllPictures() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default CreationHelper getCreationHelper() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isHidden() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setHidden(boolean hiddenFlag) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setSheetHidden(int sheetIx, boolean hidden) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default SheetVisibility getSheetVisibility(int i) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setSheetVisibility(int i, SheetVisibility sheetVisibility) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void addToolPack(UDFFinder toopack) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setForceFormulaRecalculation(boolean value) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean getForceFormulaRecalculation() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default SpreadsheetVersion getSpreadsheetVersion() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default int addOlePackage(byte[] bytes, String s, String s1, String s2) throws IOException {
		throw new UnsupportedOperationException();
	}

	@Override
	default CellReferenceType getCellReferenceType() {
		throw new UnsupportedOperationException();
	}

	@Override
	default void setCellReferenceType(CellReferenceType cellReferenceType) {
		throw new UnsupportedOperationException();
	}

}
