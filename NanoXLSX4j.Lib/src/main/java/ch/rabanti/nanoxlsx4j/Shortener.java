/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.styles.Style;

/**
 * Class to provide access to the current worksheet with a shortened syntax<br>
 * Note: The WS object can be null if the workbook was created without a
 * worksheet. The object will be available as soon as the current worksheet is
 * defined
 *
 * @author Raphael Stoeckli
 */
public class Shortener {
	private Worksheet currentWorksheet;
	private final Workbook workbookReference;

	/**
	 * Constructor with workbook reference
	 *
	 * @param reference
	 *            Workbook reference
	 */
	public Shortener(Workbook reference) {
		this.workbookReference = reference;
	}

	/**
	 * Sets the worksheet accessed by the shortener
	 *
	 * @param worksheet
	 *            Current worksheet
	 */
	public void setCurrentWorksheet(Worksheet worksheet) {
		this.workbookReference.setCurrentWorksheet(worksheet);
		this.currentWorksheet = worksheet;
	}

	/**
	 * Sets the worksheet accessed by the shortener, invoked by the workbook
	 *
	 * @param worksheet
	 *            Current worksheet
	 */
	void setCurrentWorksheetInternal(Worksheet worksheet) {
		this.currentWorksheet = worksheet;
	}

	/**
	 * Sets a value into the current cell and moves the cursor to the next cell
	 * (column or row depending on the defined cell direction)
	 *
	 * @param value
	 *            Value to set
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 */
	public void value(Object value) {
		nullCheck();
		this.currentWorksheet.addNextCell(value);
	}

	/**
	 * Sets a value with style into the current cell and moves the cursor to the
	 * next cell (column or row depending on the defined cell direction)
	 *
	 * @param value
	 *            Value to set
	 * @param style
	 *            Style to set
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 */
	public void value(Object value, Style style) {
		nullCheck();
		this.currentWorksheet.addNextCell(value, style);
	}

	/**
	 * Sets a formula into the current cell and moves the cursor to the next cell
	 * (column or row depending on the defined cell direction)
	 *
	 * @param formula
	 *            Formula to set
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 */
	public void formula(String formula) {
		nullCheck();
		this.currentWorksheet.addNextCellFormula(formula);
	}

	/**
	 * Sets a formula with style into the current cell and moves the cursor to the
	 * next cell (column or row depending on the defined cell direction)
	 *
	 * @param formula
	 *            Formula to set
	 * @param style
	 *            Style to set
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 */
	public void formula(String formula, Style style) {
		nullCheck();
		this.currentWorksheet.addNextCellFormula(formula, style);
	}

	/**
	 * Moves the cursor one row down
	 *
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 */
	public void down() {
		nullCheck();
		this.currentWorksheet.goToNextRow();
	}

	/**
	 * Moves the cursor the number of defined rows down
	 *
	 * @param numberOfRows
	 *            Number of rows to move
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 */
	public void down(int numberOfRows) {
		nullCheck();
		this.currentWorksheet.goToNextRow(numberOfRows);
	}

	/**
	 * Moves the cursor the number of defined rows down
	 *
	 * @param numberOfRows
	 *            Number of rows to move
	 * @param keepColumnPosition
	 *            If true, the column position is preserved, otherwise set to 0
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 */
	public void down(int numberOfRows, boolean keepColumnPosition) {
		nullCheck();
		this.currentWorksheet.goToNextRow(numberOfRows, keepColumnPosition);
	}

	/**
	 * Moves the cursor one row up
	 *
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 * @throws RangeException
	 *             Throws a RangeException if the row number is below 0
	 */
	public void up() {
		nullCheck();
		this.currentWorksheet.goToNextRow(-1);
	}

	/**
	 * Moves the cursor the number of defined rows up
	 *
	 * @param numberOfRows
	 *            Number of rows to move
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 * @throws RangeException
	 *             Throws a RangeException if the row number is below 0
	 * @apiNote Values can be also negative. However, this is the equivalent of the
	 *          function {@link #down(int)}}
	 */
	public void up(int numberOfRows) {
		nullCheck();
		this.currentWorksheet.goToNextRow(-1 * numberOfRows);
	}

	/**
	 * Moves the cursor the number of defined rows up
	 *
	 * @param numberOfRows
	 *            Number of rows to move
	 * @param keepColumnPosition
	 *            If true, the column position is preserved, otherwise set to 0
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 * @throws RangeException
	 *             Throws a RangeException if the row number is below 0
	 * @apiNote Values can be also negative. However, this is the equivalent of the
	 *          function {@link #down(int)}}
	 */
	public void up(int numberOfRows, boolean keepColumnPosition) {
		nullCheck();
		this.currentWorksheet.goToNextRow(-1 * numberOfRows, keepColumnPosition);
	}

	/**
	 * Moves the cursor one column to the right
	 *
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 */
	public void right() {
		nullCheck();
		this.currentWorksheet.goToNextColumn();
	}

	/**
	 * Moves the cursor the number of defined columns to the right
	 *
	 * @param numberOfColumns
	 *            Number of columns to move
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 */
	public void right(int numberOfColumns) {
		nullCheck();
		this.currentWorksheet.goToNextColumn(numberOfColumns);
	}

	/**
	 * Moves the cursor the number of defined columns to the right
	 *
	 * @param numberOfColumns
	 *            Number of columns to move
	 * @param keepRowPosition
	 *            If true, the row position is preserved, otherwise set to 0
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 */
	public void right(int numberOfColumns, boolean keepRowPosition) {
		nullCheck();
		this.currentWorksheet.goToNextColumn(numberOfColumns, keepRowPosition);
	}

	/**
	 * Moves the cursor one column to the left
	 *
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 * @throws RangeException
	 *             Throws a RangeException if the column number is below 0
	 */
	public void left() {
		nullCheck();
		this.currentWorksheet.goToNextColumn(-1);
	}

	/**
	 * Moves the cursor the number of defined columns to the left
	 *
	 * @param numberOfColumns
	 *            Number of columns to move
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 * @throws RangeException
	 *             Throws a RangeException if the column number is below 0
	 * @apiNote Values can be also negative. However, this is the equivalent of the
	 *          function {@link #right(int)}}
	 */
	public void left(int numberOfColumns) {
		nullCheck();
		this.currentWorksheet.goToNextColumn(-1 * numberOfColumns);
	}

	/**
	 * Moves the cursor the number of defined columns to the left
	 *
	 * @param numberOfColumns
	 *            Number of columns to move
	 * @param keepRowPosition
	 *            If true, the row position is preserved, otherwise set to 0
	 * @throws WorksheetException
	 *             Throws a WorksheetException if no worksheet was defined
	 * @throws RangeException
	 *             Throws a RangeException if the column number is below 0
	 * @apiNote Values can be also negative. However, this is the equivalent of the
	 *          function {@link #right(int)}}
	 */
	public void left(int numberOfColumns, boolean keepRowPosition) {
		nullCheck();
		this.currentWorksheet.goToNextColumn(-1 * numberOfColumns, keepRowPosition);
	}

	/**
	 * Internal method to check whether the worksheet is null
	 */
	private void nullCheck() {
		if (this.currentWorksheet == null) {
			throw new WorksheetException("No worksheet was defined");
		}
	}

}
