package com.monitorjbl.xlsx.notsupportedoperations;

import com.monitorjbl.xlsx.exceptions.NotSupportedException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public interface RowAdapter extends Row {
	/* Not supported */

	/**
	 * Not supported
	 */
	@Override
	default Cell createCell(int column) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Cell createCell(int i, CellType cellType) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void removeCell(Cell cell) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setRowNum(int rowNum) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setZeroHeight(boolean zHeight) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default short getHeight() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setHeight(short height) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default float getHeightInPoints() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setHeightInPoints(float height) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isFormatted() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default CellStyle getRowStyle() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setRowStyle(CellStyle style) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default int getOutlineLevel() {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void shiftCellsRight(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void shiftCellsLeft(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {
		throw new NotSupportedException();
	}
}
