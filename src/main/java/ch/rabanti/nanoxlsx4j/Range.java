/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

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
            if (this.StartAddress.equals(range.StartAddress) && this.EndAddress.equals(range.EndAddress)){
                return true;
            }
            else {
                return false;
            }
        }
        
    } 