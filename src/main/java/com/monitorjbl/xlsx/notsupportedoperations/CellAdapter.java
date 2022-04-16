package com.monitorjbl.xlsx.notsupportedoperations;

import com.monitorjbl.xlsx.exceptions.NotSupportedException;

import org.apache.poi.ss.formula.FormulaParseException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.util.CellRangeAddress;

import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

public interface CellAdapter extends Cell {

	/**
	 * Not supported
	 */
	@Override
	default void setCellType(CellType cellType) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setCellValue(double value) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setCellValue(Date value) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setCellValue(LocalDateTime value) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setCellValue(Calendar value) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setCellValue(RichTextString value) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setCellValue(String value) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setCellFormula(String formula) throws FormulaParseException {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setCellValue(boolean value) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setCellErrorValue(byte value) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default byte getErrorCellValue() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setAsActiveCell() {
		throw new NotSupportedException();
	}


	/**
	 * Not supported
	 */
	@Override
	default Comment getCellComment() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setCellComment(Comment comment) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void removeCellComment() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Hyperlink getHyperlink() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setHyperlink(Hyperlink link) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void removeHyperlink() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default CellRangeAddress getArrayFormulaRange() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isPartOfArrayFormulaGroup() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setBlank() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void removeFormula() throws IllegalStateException {
		throw new NotSupportedException();
	}
}
