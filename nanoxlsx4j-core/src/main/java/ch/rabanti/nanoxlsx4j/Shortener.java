/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.styles.Style;

/**
 * Class to provide access to the current worksheet with a shortened syntax.
 *
 * <p>Remarks: Note: The WS object can be null if the workbook was created without a worksheet.
 * The object will be available as soon as the current worksheet is defined.
 * </p>
 */
public class Shortener {

    private Worksheet currentWorksheet;
    private final Workbook workbookReference;

    /**
     * Constructor with workbook reference
     *
     * @param reference Workbook reference
     */
    public Shortener(Workbook reference) {
        this.workbookReference = reference;
        this.currentWorksheet = reference.getCurrentWorksheet();

    }

    /**
     * Sets the worksheet accessed by the shortener
     *
     * @param worksheet Current worksheet
     */
    public void SetCurrentWorksheet(Worksheet worksheet) {
        workbookReference.setCurrentWorksheet(worksheet);
        currentWorksheet = worksheet;
    }

    /**
     * Sets the worksheet accessed by the shortener, invoked by the workbook
     *
     * @param currentWorksheet Current worksheet
     */
    void setCurrentWorksheetInternal(Worksheet currentWorksheet) {
        this.currentWorksheet = currentWorksheet;
    }

    /**
     * Sets a value into the current cell and moves the cursor to the next cell (column or row depending on the defined
     * cell direction)
     *
     * @param cellValue Value to set
     * @throws WorksheetException Throws a WorksheetException if no worksheet was defined
     */
    public void value(Object cellValue) {
        nullCheck();
        currentWorksheet.addNextCell(cellValue);
    }

    /**
     * Sets a value with style into the current cell and moves the cursor to the next cell (column or row depending on
     * the defined cell direction)
     *
     * @param cellValue Value to set
     * @param style     Style to apply
     * @throws WorksheetException Throws a WorksheetException if no worksheet was defined
     */
    public void value(Object cellValue, Style style) {
        nullCheck();
        currentWorksheet.addNextCell(cellValue, style);
    }

    /**
     * Sets a formula into the current cell and moves the cursor to the next cell (column or row depending on the
     * defined cell direction)
     *
     * @param cellFormula Formula to set
     * @throws WorksheetException Throws a WorksheetException if no worksheet was defined
     */
    public void formula(String cellFormula) {
        nullCheck();
        currentWorksheet.addNextCellFormula(cellFormula);
    }

    /**
     * Sets a formula with style into the current cell and moves the cursor to the next cell (column or row depending on
     * the defined cell direction)
     *
     * @param cellFormula Formula to set
     * @param style       Style to apply
     * @throws WorksheetException Throws a WorksheetException if no worksheet was defined
     */
    public void formula(String cellFormula, Style style) {
        nullCheck();
        currentWorksheet.addNextCellFormula(cellFormula, style);
    }

    /**
     * Moves the cursor one row down
     */
    public void down() {
        nullCheck();
        currentWorksheet.goToNextRow();
    }

    /**
     * Moves the cursor the number of defined rows down. The column number is reset to 0
     *
     * <p>Remarks: An exception will be thrown if the row number is below 0. Values (number of rows) can be also
     * negative. However, this is the equivalent of the function {@link #up(int, boolean)}</p>
     *
     * @param numberOfRows Number of rows to move
     */
    public void down(int numberOfRows) {
        down(numberOfRows, false);
    }

    /**
     * Moves the cursor the number of defined rows down
     *
     * <p>Remarks: An exception will be thrown if the row number is below 0. Values (number of rows) can be also
     * negative. However, this is the equivalent of the function {@link #up(int, boolean)}</p>
     *
     * @param numberOfRows       Number of rows to move
     * @param keepColumnPosition If true, the column position is preserved, otherwise set to 0
     */
    public void down(int numberOfRows, boolean keepColumnPosition) {
        nullCheck();
        currentWorksheet.goToNextRow(numberOfRows, keepColumnPosition);
    }

    /**
     * Moves the cursor one row up
     *
     * <p>Remarks: An exception will be thrown if the row number is below 0/></p>
     */
    public void up() {
        nullCheck();
        currentWorksheet.goToNextRow(-1);
    }

    /**
     * Moves the cursor the number of defined rows up. The row number is reset to 0
     *
     * <p>Remarks: An exception will be thrown if the row number is below 0. Values can be also negative. However, this
     * is the equivalent of the function {@link #down(int, boolean)}</p>
     *
     * @param numberOfRows Number of rows to move
     */
    public void up(int numberOfRows) {
        up(numberOfRows, false);
    }

    /**
     * Moves the cursor the number of defined rows up
     *
     * <p>Remarks: An exception will be thrown if the row number is below 0. Values can be also negative. However, this
     * is the equivalent of the function {@link #down(int, boolean)}</p>
     *
     * @param numberOfRows       Number of rows to move
     * @param keepColumnPosition If true, the column position is preserved, otherwise set to 0
     */
    public void up(int numberOfRows, boolean keepColumnPosition) {
        nullCheck();
        currentWorksheet.goToNextRow(-1 * numberOfRows, keepColumnPosition);
    }

    /**
     * Moves the cursor one column to the right
     */
    public void right() {
        nullCheck();
        currentWorksheet.goToNextColumn();
    }

    /**
     * Moves the cursor the number of defined columns to the right. The row number is reset to 0
     *
     * @param numberOfColumns Number of columns to move
     */
    public void right(int numberOfColumns) {
        right(numberOfColumns, false);
    }

    /**
     * Moves the cursor the number of defined columns to the right
     *
     * @param numberOfColumns Number of columns to move
     * @param keepRowPosition If true, the row position is preserved, otherwise set to 0
     */
    public void right(int numberOfColumns, boolean keepRowPosition) {
        nullCheck();
        currentWorksheet.goToNextColumn(numberOfColumns, keepRowPosition);
    }

    /**
     * Moves the cursor one column to the left
     *
     * <p>Remarks: An exception will be thrown if the column number is below 0</p>
     */
    public void left() {
        nullCheck();
        currentWorksheet.goToNextColumn(-1);
    }

    /**
     * Moves the cursor the number of defined columns to the left. Te row number is reset to 0
     *
     * <p>Remarks: An exception will be thrown if the column number is below 0. Values can be also negative. However,
     * this is the equivalent of the function {@link #right(int, boolean)}</p>
     *
     * @param numberOfColumns Number of columns to move
     */
    public void left(int numberOfColumns) {
        left(numberOfColumns, false);
    }

    /**
     * Moves the cursor the number of defined columns to the left
     *
     * <p>Remarks: An exception will be thrown if the column number is below 0. Values can be also negative. However,
     * this is the equivalent of the function {@link #right(int, boolean)}</p>
     *
     * @param numberOfColumns    Number of columns to move
     * @param keepRowRowPosition If true, the row position is preserved, otherwise set to 0
     */
    public void left(int numberOfColumns, boolean keepRowRowPosition) {
        nullCheck();
        currentWorksheet.goToNextColumn(-1 * numberOfColumns, keepRowRowPosition);
    }

    /**
     * Internal method to check whether the worksheet is null
     */
    private void nullCheck() {
        if (currentWorksheet == null) {
            throw new WorksheetException("No worksheet was defined");
        }
    }

}
