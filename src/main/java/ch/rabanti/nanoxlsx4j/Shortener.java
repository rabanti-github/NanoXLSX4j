/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2019
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.styles.Style;

/**
 * Class to provide access to the current worksheet with a shortened syntax<br>
 * Note: The WS object can be null if the workbook was created without a worksheet. The object will be available as soon as the current worksheet is defined
 * @author Raphael Stoeckli
 */
    public class Shortener
    {
        private Worksheet currentWorksheet;
        
        /**
         * Default constructor
         */
        public Shortener()
        {}
        
        /**
         * Sets the worksheet accessed by the shortener
         * @param worksheet Current worksheet
         */
        public void setCurrentWorksheet(Worksheet worksheet)
        {
            this.currentWorksheet = worksheet;
        }
        
        /**
         * Sets a value into the current cell and moves the cursor to the next cell (column or row depending on the defined cell direction)
         * @param value Value to set
         * @throws WorksheetException Throws a WorksheetException if no worksheet was defined
         */
        public void value(Object value)
        {
            this.currentWorksheet.addNextCell(value);
        }
        
        /**
         * Sets a value with style into the current cell and moves the cursor to the next cell (column or row depending on the defined cell direction)
         * @param value Value to set
         * @param style Style to set
         * @throws WorksheetException Throws a WorksheetException if no worksheet was defined
         */
        public void value(Object value, Style style)
        {
            this.currentWorksheet.addNextCell(value, style);
        }
        
        /**
         * Sets a formula into the current cell and moves the cursor to the next cell (column or row depending on the defined cell direction)
         * @param formula Formula to set
         * @throws WorksheetException Throws a WorksheetException if no worksheet was defined
         */
        public void formula(String formula)
        {
            this.currentWorksheet.addNextCellFormula(formula);
        }
        
        /**
         * Sets a formula with style into the current cell and moves the cursor to the next cell (column or row depending on the defined cell direction)
         * @param formula Formula to set
         * @param style Style to set
         * @throws WorksheetException Throws a WorksheetException if no worksheet was defined
         */
        public void formula(String formula, Style style)
        {
            this.currentWorksheet.addNextCellFormula(formula, style);
        }
        
        /**
         * Moves the cursor one row down
         * @throws WorksheetException Throws a WorksheetException if no worksheet was defined
         */
        public void down()
        {
            nullCheck();
            this.currentWorksheet.goToNextRow();
        }
        
        /**
         * Moves the cursor the number of defined rows down
         * @param numberOfRows Number of rows to move
         * @throws WorksheetException Throws a WorksheetException if no worksheet was defined
         */
        public void down(int numberOfRows)
        {
            nullCheck();
            this.currentWorksheet.goToNextRow(numberOfRows);
        }        
        
        /**
         * Moves the cursor one column to the right
         * @throws WorksheetException Throws a WorksheetException if no worksheet was defined
         */
        public void right()
        {
            nullCheck();
            this.currentWorksheet.goToNextColumn();
        }
        
        /**
         * Moves the cursor the number of defined columns to the right
         * @param numberOfColumns Number of columns to move
         * @throws WorksheetException Throws a WorksheetException if no worksheet was defined
         */
        public void right(int numberOfColumns)
        {
            nullCheck();
            this.currentWorksheet.goToNextColumn(numberOfColumns);
        }
        
        /**
         * Internal method to check whether the worksheet is null
         */
        private void nullCheck()
        {
            if (this.currentWorksheet == null)
            {
                throw new WorksheetException("UndefinedWorksheetException", "No worksheet was defined");
            }
        }
        
    }
