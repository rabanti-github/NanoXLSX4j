/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2019
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
/**
 * Class representing a cell range (no getters and setters to simplify handling)
 * @author Raphael Stoeckli
 */
    public class Range
    {
// ### P U B L I C  F I E L D S ###        
        /**
         * End address of the range
         */
        public final Address EndAddress;
        /**
         * Start address of the range
         */
        public final Address StartAddress;
        
// ### C O N S T R U C T O R S ###        
        /**
         * Constructor with with addresses as arguments
         * @param start Start address of the range
         * @param end End address of the range
         */
        public Range(Address start, Address end)
        {
            if (end.compareTo(start) < 0){
                throw new RangeException("RangeFormatException", "The end address of the range cannot be smaller than the start address");
            }
            this.StartAddress = start;
            this.EndAddress = end;
        }
        
        /**
         * Constructor with a range string as argument
         * @param range Address range (e.g. 'A1:B12')
         */
        public Range(String range)
        {
            Range r = Cell.resolveCellRange(range);
            if (r.StartAddress.compareTo(r.EndAddress) > 1){
                throw new RangeException("RangeFormatException", "The end address of the range cannot be smaller than the start address");
            }
            this.StartAddress = r.StartAddress;
            this.EndAddress = r.EndAddress;
        }
        
// ### M E T H O D S ###        
        /**
         * Overwritten toString method
         * @return Returns the range (e.g. 'A1:B12')
         */
        @Override
        public String toString()
        {
            return StartAddress.toString() + ":" + EndAddress.toString();
        }

        /**
         * Overwritten equals method
         * @param o Other object to compare
         * @return True if this instance is equal to the other instance
         */
        @Override
        public boolean equals(Object o){
            if(o == this){
                return true;
            }
            if (!(o instanceof Range)){
                return false;
            }
            Range range = (Range)o;
            return this.StartAddress.equals(range.StartAddress) && this.EndAddress.equals(range.EndAddress);
        }

// ### S T A T I C   M E T H O D S ###

        /**
         * Method to build a range object from two addresses. The appropriate start and end address will be determined automatically
         * @param startAddress Proposed start address
         * @param endAddress Proposed end address
         * @return Resolved Range
         */
        public static Range buildRange(Address startAddress, Address endAddress){
            return buildRange(startAddress.Column, startAddress.Row, endAddress.Column, endAddress.Row);
        }

        /**
         * Method to build a range object from zwo column and two row numbers. The appropriate start and end address will be determined automatically
         * @param startColumn Proposed start column
         * @param startRow Proposed start row
         * @param endColumn Proposed end column
         * @param endRow Proposed end row
         * @return Resolved Range
         */
        public static Range buildRange(int startColumn, int startRow, int endColumn, int endRow){
            Address start, end;
            if (startColumn < endColumn && startRow < endRow) {
                start = new Address(startColumn, startRow);
                end = new Address(endColumn, endRow);
            }
            else if (startColumn < endColumn && startRow >= endRow){
                start = new Address(startColumn, endRow);
                end = new Address(endColumn, startRow);
            }
            else if (startColumn >= endColumn && startRow < endRow){
                start = new Address(endColumn, startRow);
                end = new Address(startColumn, endRow);
            }
            else {
                start = new Address(endColumn, endRow);
                end = new Address(startColumn, startRow);
            }
            return new Range(start, end);
        }
        
    } 