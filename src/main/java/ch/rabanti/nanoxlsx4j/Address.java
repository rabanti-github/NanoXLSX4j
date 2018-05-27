/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exception.RangeException;

import static ch.rabanti.nanoxlsx4j.Cell.resolveCellAddress;

/**
 * C class representing a cell address (no getters and setters to simplify handling)
 * @author Raphael Stoeckli
 */
    public class Address
    {
       
// ### P U B L I C  F I E L D S ###        
        /**
         * Column of the address (zero-based)
         */        
        public final int Column;
        /**
         * Row of the address (zero-based)
         */
        public final int Row;
        
// ### C O N S T R U C T O R S ###        
        /**
         * Constructor with with row and column as arguments
         * @param column Column of the address (zero-based)
         * @param row Row of the address (zero-based)
         */
        public Address(int column, int row)
        {
            this.Column = column;
            this.Row = row;
        }
        
        public Address(String address)
        {
         Address a = Cell.resolveCellCoordinate(address);
         this.Column = a.Column;
         this.Row = a.Row;
        }

// ### M E T H O D S ###
        
        /**
         * Compares two addresses whether they are equal
         * @param o Compares two addresses whether they are equal
         * @return True if equal
         */
        public boolean equals(Address o)
        {
            if(this.Row == o.Row && this.Column == o.Column)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        
        /**
         * Gets the address as string
         * @return Address as string
         * @throws RangeException Thrown if the column or row is out of range
         */
        public String getAddress()
        {
            return resolveCellAddress(this.Column, this.Row);
        }
        
        /**
         * Returns the address as string or "Illegal Address" in case of an exception
         * @return Address or notification in case of an error
         */
        @Override
        public String toString()
        {
            try
            {
            return getAddress();
            }
            catch(Exception e)
            {
                return "Illegal Address";
            }
        }
    }