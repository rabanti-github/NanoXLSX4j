/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exception.FormatException;
import ch.rabanti.nanoxlsx4j.exception.RangeException;
import ch.rabanti.nanoxlsx4j.exception.StyleException;
import ch.rabanti.nanoxlsx4j.style.Style;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Class representing a cell of a worksheet
 * @author Raphael Stoeckli
 */
public class Cell implements Comparable<Cell>{
    
// ### E N U M S ###
    /**
     * Enum defines the basic data types of a cell
     */
    public enum CellType
    {
        /**
         * Type for single characters and strings
         */
        STRING,
        /**
         * Type for all numeric types (long, integer and float and double)
         */
        NUMBER,
        /**
         * Type for dates and times (Note: Dates before 1900-01-01 are not allowed)
         */
        DATE,
        /**
         * Type for boolean
         */
        BOOL,
        /**
         * Type for Formulas (The cell will be handled differently)
         */
        FORMULA,
        /**
         * Type for empty cells. This type is only used for merged cells (all cells except the first of the cell range)
         */
        EMPTY,
        /**
         * Default Type, not specified
         */
        DEFAULT
    }    
    
// ### P R I V A T E  F I E L D S ###
    
    private Style cellStyle;
    private int columnNumber;
    private CellType dataType;
    private int rowNumber;
    private Object value;
    private Worksheet worksheetReference;
    
// ### G E T T E R S  &  S E T T E R S ###

 /**
     * Gets the combined cell Address as string in the format A1 - XFD1048576
     * @return Cell address
     */
    public String getCellAddress()
    {
        return Cell.resolveCellAddress(this.columnNumber, this.rowNumber);
    }
    /**
     * Sets the combined cell Address as string in the format A1 - XFD1048576
     * @param address Cell address
     * @throws RangeException Thrown in case of a illegal address
     */
    public void setCellAddress(String address)
    {
        Address temp = Cell.resolveCellCoordinate(address);
        this.columnNumber = temp.Column;
        this.rowNumber = temp.Row;
    }
    /**
     * Gets the combined cell address as class
     * @return Cell address
     */
    public Address getCellAddress2()
    {
        return new Address(this.columnNumber, this.rowNumber);
    }
    /**
     * Sets the combined cell address as class
     * @param address Cell address
     */
    public void setCellAddress2(Address address)
    {
        this.setColumnNumber(address.Column);
        this.setRowNumber(address.Row);
    }

    /**
     * Gets the assigned style of the cell
     * @return Assigned style
     */
    public Style getCellStyle() {
        return cellStyle;
    }


    /**
     * Gets the number of the column (zero-based)
     * @return Column number (zero-based)
     */    
    public int getColumnNumber() {
        return columnNumber;
    }

    /**
     * Sets the number of the column (zero-based)
     * @param columnNumber Column number (zero-based)
     */
    public void setColumnNumber(int columnNumber) {
        if (columnNumber < Worksheet.MIN_COLUMN_NUMBER || columnNumber > Worksheet.MAX_COLUMN_NUMBER)
        {
            throw new RangeException("OutOfRangeException","The passed number (" + Integer.toString(columnNumber) + ")is out of range. Range is from " + Integer.toString(Worksheet.MIN_COLUMN_NUMBER) + " to " + Integer.toString(Worksheet.MAX_COLUMN_NUMBER) + " (" + (Integer.toString(Worksheet.MAX_COLUMN_NUMBER + 1)) + " rows).");
        }        
        this.columnNumber = columnNumber;
    }

    /**
     * Gets the type of the cell
     * @return Type of the cell
     */
    public CellType getDataType() {
        return dataType;
    }

    /**
     * Sets the type of the cell
     * @param dataType Type of the cell
     */
    public void setDataType(CellType dataType) {
        this.dataType = dataType;
    }
    /**
     * Gets the number of the row (zero-based)
     * @return Row number (zero-based)
     */
    public int getRowNumber() {
        return rowNumber;
    }
    /**
     * Sets the number of the row (zero-based)
     * @param rowNumber Row number (zero-based)
     */
    public void setRowNumber(int rowNumber) {
        if (rowNumber < Worksheet.MIN_ROW_NUMBER || rowNumber > Worksheet.MAX_ROW_NUMBER)
        {
            throw new RangeException("OutOfRangeException","The passed number (" + Integer.toString(rowNumber) + ")is out of range. Range is from " + Integer.toString(Worksheet.MIN_ROW_NUMBER) + " to " + Integer.toString(Worksheet.MAX_ROW_NUMBER) + " (" + (Integer.toString(Worksheet.MAX_ROW_NUMBER + 1)) + " rows).");
        }
        this.rowNumber = rowNumber;
    }
    /**
     * Gets the value of the cell (generic object type)
     * @return Value of the cell
     */
    public Object getValue() {
        return value;
    }
    /**
     * Sets the value of the cell (generic object type)
     * @param value Value of the cell
     */
    public void setValue(Object value) {
        this.value = value;
    } 
    
    /**
     * Gets or sets the parent worksheet reference
     * @return Worksheet reference
     */
    public Worksheet getWorksheetReference()
    {
        return this.worksheetReference;
    }
    
    /**
     * Sets the parent worksheet reference
     * @param reference Worksheet reference
     */
    public void setWorksheetReference(Worksheet reference)
    {
        this.worksheetReference = reference;
    }
    
    
// ### C O N S T R U C T O R S ###
    
    /**
     * Default constructor. Cells created with this constructor do not have a link to a worksheet initially
     */
    public Cell()
    {
        this.worksheetReference = null;
        this.dataType = CellType.DEFAULT;
    }
    /**
     * Constructor with value and cell type. Cells created with this constructor do not have a link to a worksheet initially
     * @param value Value of the cell
     * @param type Type of the cell
     */
    public Cell(Object value, CellType type)
    {
        this.dataType = type;
        this.value = value;
        if (type == CellType.DEFAULT)
        {
            resolveCellType();
        }
    }

    /**
     * Constructor with value, cell type and address. The worksheet reference is set to null and must be assigned later
     * @param value Value of the cell
     * @param type Type of the cell
     * @param address Address of the cell
     */
    public Cell(Object value, CellType type, String address)
    {
        this.dataType = type;
        this.value = value;
        this.setCellAddress(address);
        this.worksheetReference = null;
        if (type == CellType.DEFAULT)
        {
            resolveCellType();
        }
    }

    /**
     * Constructor with value, cell type, row number, column number and the link to a worksheet
     * @param value Value of the cell
     * @param type Type of the cell
     * @param column Column number of the cell (zero-based)
     * @param row Row number of the cell (zero-based)
     * @param reference Worksheet reference
     */
    public Cell(Object value, CellType type, int column, int row, Worksheet reference)
    {
        this.dataType = type;
        this.value = value;
        this.columnNumber = column;
        this.rowNumber = row;
        this.worksheetReference = reference;
        if (type == CellType.DEFAULT)
        {
            resolveCellType();
        }
    }
    
// ### M E T H O D S ###
    
    /**
     * Implemented compareTo method
     * @param o Object to compare
     * @return 0 if values are the same, -1 if this object is smaller, 1 if it is bigger
     */
    @Override
    public int compareTo(Cell o) {
        if (this.rowNumber == o.rowNumber)
        {
            return Integer.compare(this.columnNumber, o.getColumnNumber());
        }
        else
        {
            return Integer.compare(this.rowNumber, o.getRowNumber());
        }
    }
   
    /**
     * Removes the assigned style from the cell
     * @throws StyleException Thrown if the workbook to remove was not found in the style sheet collection
     */
    public void removeStyle()
    {
        if (this.worksheetReference == null)
        {
            throw new StyleException("MissingReferenceException","No worksheet reference was defined while trying to remove a style from a cell");
        }
        if (this.worksheetReference.getWorkbookReference() == null)
        {
            throw new StyleException("MissingReferenceException","No workbook reference was defined on the worksheet while trying to remove a style from a cell");
        }
        if (this.cellStyle != null)
        {
            String styleName = this.cellStyle.getName();
            this.cellStyle = null;
            this.worksheetReference.getWorkbookReference().removeStyle(styleName, true);
        }
    }
    
     /**
      * Method resets the Cell type and tries to find the actual type. This is used if a Cell was created with the CellType DEFAULT. CellTypes FORMULA and EMPTY will skip this method
      */
    public void resolveCellType()
    {
        if(this.value == null)
        {
            this.setDataType(CellType.EMPTY);
            value = "";
            return;
        } // the following section is intended to be as similar as possible to PicoXLSX for C#
        if (this.dataType == CellType.FORMULA || this.dataType == CellType.EMPTY) {return;}
        else if (value instanceof Boolean)      { this.dataType = CellType.BOOL; }
        else if (value instanceof Byte)         { this.dataType = CellType.NUMBER; } // sbyte not existing in Java
        else if (value instanceof BigDecimal)   { this.dataType = CellType.NUMBER; } // decimal
        else if (value instanceof Double)       { this.dataType = CellType.NUMBER; }
        else if (value instanceof Float)        { this.dataType = CellType.NUMBER; }
        else if (value instanceof Integer)      { this.dataType = CellType.NUMBER; } // uint not existing in Java
        else if (value instanceof Long)         { this.dataType = CellType.NUMBER; }
        else if (value instanceof Short)        { this.dataType = CellType.NUMBER; } // ushort not existing in Java
        else if (value instanceof Date)         { this.dataType = CellType.DATE; }
        else { this.dataType = CellType.STRING; } // Default (char, string, object)
    }
    /**
     * Sets the lock state of the cell
     * @param isLocked If true, the cell will be locked if the worksheet is protected
     * @param isHidden If true, the value of the cell will be invisible if the worksheet is protected
     */
    public void setCellLockedState(boolean isLocked, boolean isHidden)
    {
        Style lockStyle;
        if (this.cellStyle == null)
        {
            lockStyle = new Style();
        }
        else
        {
            lockStyle = this.cellStyle.copyStyle();
        }
        lockStyle.getCellXf().setLocked(isLocked);
        lockStyle.getCellXf().setHidden(isHidden);
        try
        {
            this.setStyle(lockStyle);
        }
        catch(Exception e)
        {
            // Should never happen
        }
    }
    
    /**
     * Sets the style of the cell
     * @param style style to assign
     * @return If the passed style already exists in the workbook, the existing one will be returned, otherwise the passed one
     * @throws StyleException Thrown if the style is not referenced in the workbook
     */
    public Style setStyle(Style style)
    {
       if (this.worksheetReference == null)
       {
           throw new StyleException("MissingReferenceException","No worksheet reference was defined while trying to set a style to a cell");
       }
       if (this.worksheetReference.getWorkbookReference() == null)
       {
           throw new StyleException("MissingReferenceException","No workbook reference was defined on the worksheet while trying to set a style to a cell");
       }
       if (style == null)
       {
           throw new StyleException("MissingReferenceException","No style to assign was defined");
       }
       Style s = this.worksheetReference.getWorkbookReference().addStyle(style);
       this.cellStyle = s;
       return s;
    }
    
// ### S T A T I C   M E T H O D S ###
    
    /**
     * Get a list of cell addresses from a cell range
     * @param startColumn Start column (zero based)
     * @param startRow Start row (zero based)
     * @param endColumn End column (zero based)
     * @param endRow End row (zero based)
     * @return List of cell addresses
     */
    public static List<Address> getCellRange(int startColumn, int startRow, int endColumn, int endRow)
    {
        Address start = new Address(startColumn, startRow);
        Address end = new Address(endColumn, endRow);
        return getCellRange(start, end);       
    }
    
    /**
     * Converts a List of supported objects into a list of cells
     * @param <T> Generic data type
     * @param list List of generic objects
     * @return List of cells
     */
    public  static <T> List<Cell> convertArray(List<T> list)
    {
        List<Cell> output = new ArrayList<>();
        Cell c;
        Object o;
        for (int i = 0; i < list.size(); i++)
        {
            o = list.get(i);
            if (o instanceof Boolean)         { c = new Cell(o, CellType.BOOL);   }
            if (o instanceof Byte)            { c = new Cell(o, CellType.NUMBER); }
            else if (o instanceof BigDecimal) { c = new Cell(o, CellType.NUMBER); }
            else if (o instanceof Double)     { c = new Cell(o, CellType.NUMBER); }
            else if (o instanceof Float)      { c = new Cell(o, CellType.NUMBER); }
            else if (o instanceof Integer)    { c = new Cell(o, CellType.NUMBER); }
            else if (o instanceof Long)       { c = new Cell(o, CellType.NUMBER); }
            else if (o instanceof Short)      { c = new Cell(o, CellType.NUMBER); }
            else if (o instanceof Date)       { c = new Cell(o, CellType.DATE);   }
            else if (o instanceof String)     { c = new Cell(o, CellType.STRING); }
            else
            {
                c = new Cell(o, CellType.DEFAULT);
            }
            output.add(c);
        }
        return output;
    }
    
    /**
     * Gets a list of cell addresses from a cell range (format A1:B3 or AAD556:AAD1000)
     * @param range Range to process
     * @return List of cell addresses
     * @throws FormatException Thrown if the passed address range is malformed
     */
    public static List<Address> getCellRange(String range)
    {
       Range range2 = resolveCellRange(range);
       return getCellRange(range2.StartAddress, range2.EndAddress);
    }
    
    /**
     * Get a list of cell addresses from a cell range
     * @param startAddress Start address as string in the format A1 - XFD1048576
     * @param endAddress End address as string in the format A1 - XFD1048576
     * @return List of cell addresses
     * @throws FormatException Thrown if one of the passed addresses contains malformed information
     * @throws RangeException Thrown if one of the passed addresses is out of range
     */
    public static List<Address> getCellRange(String startAddress, String endAddress)
    {
        Address start = resolveCellCoordinate(startAddress);
        Address end = resolveCellCoordinate(endAddress);
        return getCellRange(start, end);
    }    
    
    /**
     * Get a list of cell addresses from a cell range
     * @param startAddress Start address
     * @param endAddress End address
     * @return List of cell addresses
     */
    public static List<Address> getCellRange(Address startAddress, Address endAddress)
    {
            int startColumn, endColumn, startRow, endRow;
            if (startAddress.Column < endAddress.Column)
            {
                startColumn = startAddress.Column;
                endColumn = endAddress.Column;
            }
            else
            {
                startColumn = endAddress.Column;
                endColumn = startAddress.Column;
            }
            if (startAddress.Row < endAddress.Row)
            {
                startRow = startAddress.Row;
                endRow = endAddress.Row;
            }
            else
            {
                startRow = endAddress.Row;
                endRow = startAddress.Row;
            }
            List<Address> output = new ArrayList<>();
            for (int i = startRow; i <= endRow; i++)
            {
                for (int j = startColumn; j <= endColumn; j++)
                {
                    output.add(new Address(j, i));
                }
            }
            return output;
    }
    
    /**
     * Gets the address of a cell by the column and row number (zero based)
     * @param column Column address of the cell (zero-based)
     * @param row Row address of the cell (zero-based)
     * @return Cell Address as string in the format A1 - XFD1048576
     * @throws RangeException Thrown if the start or end address was out of range
     */
    public static String resolveCellAddress(int column, int row)
    {
            if (row > Worksheet.MAX_ROW_NUMBER || row < Worksheet.MIN_ROW_NUMBER)
            {
                throw new RangeException("OutOfRangeException","The row number (" + Integer.toString(row) + ") is out of range. Range is from " + Integer.toString(Worksheet.MIN_ROW_NUMBER) + " to " + Integer.toString(Worksheet.MAX_ROW_NUMBER) + " (" + (Integer.toString(Worksheet.MIN_ROW_NUMBER) + 1) + " rows).");
            }
            return resolveColumnAddress(column) + Integer.toString(row + 1);    
    }
    
    /**
     * Gets the column and row number (zero based) of a cell by the address
     * @param address Address as string in the format A1 - XFD1048576
     * @return Address object of the passed string
     * @throws FormatException Thrown if the passed address was malformed
     * @throws RangeException Thrown if the resolved address is out of range
     */
    public static Address resolveCellCoordinate(String address)
    {
        int row, column;
        if (Helper.isNullOrEmpty(address))
        {
            throw new FormatException("FormatException","The cell address is null or empty and could not be resolved");
        }
        address = address.toUpperCase();
        Pattern pattern = Pattern.compile("([A-Z]{1,3})([0-9]{1,7})");
        Matcher mx = pattern.matcher(address);
        if (mx.groupCount() != 2)
        {
            throw new FormatException("FormatException","The format of the cell address (" + address + ") is malformed");
        }
        mx.find();
        int digits = Integer.parseInt(mx.group(2));
        column = resolveColumn(mx.group(1));
        row = digits - 1;
        
        if (row > Worksheet.MAX_ROW_NUMBER || row < Worksheet.MIN_ROW_NUMBER)
        {
            throw new RangeException("OutOfRangeException","The row number (" + Integer.toString(row) + ") is out of range. Range is from " + Integer.toString(Worksheet.MIN_ROW_NUMBER) + " to " + Integer.toString(Worksheet.MAX_ROW_NUMBER) + " (" + Integer.toString((Worksheet.MAX_ROW_NUMBER + 1)) + " rows).");
        }     
        if (column > Worksheet.MAX_COLUMN_NUMBER || column < Worksheet.MIN_COLUMN_NUMBER)
        {
            throw new RangeException("OutOfRangeException","The column number (" + Integer.toString(column) + ") is out of range. Range is from " + Integer.toString(Worksheet.MIN_COLUMN_NUMBER) + " to " + Integer.toString(Worksheet.MAX_COLUMN_NUMBER) + " (" + Integer.toString((Worksheet.MAX_COLUMN_NUMBER + 1)) + " columns).");
        }

        return new Address(column, row);
    } 
    /**
     * Resolves a cell range from the format like A1:B3 or AAD556:AAD1000
     * @param range Range to process
     * @return Range object of the passed string range
     * @throws FormatException Thrown if the passed range is malformed
     */
    public static Range resolveCellRange(String range)
    {
        if (Helper.isNullOrEmpty(range))
        {
            throw new FormatException("FormatException","The cell range is null or empty and could not be resolved");
        }
        String[] split = range.split(":");
        if (split.length != 2)
        {
            throw new FormatException("FormatException","The cell range (" + range + ") is malformed and could not be resolved");
        }
        try
        {
            Address start = resolveCellCoordinate(split[0]);
            Address end = resolveCellCoordinate(split[1]);
            return new Range(start, end);
        }
        catch(Exception e)
        {
            throw new FormatException("FormatException","The start address or end address could not be resolved. See inner exception", e);
        }
    }
   
    /**
     * Gets the column number from the column address (A - XFD)
     * @param columnAddress Column address (A - XFD)
     * @return Column number (zero-based)
     * @throws RangeException Thrown if the column is out of range
     */
    public static int resolveColumn(String columnAddress)
    {
        int chr;
        int result = 0;
        int multiplier = 1;
        
        for (int i = columnAddress.length() - 1; i >= 0; i--)
        {
            chr = (int)columnAddress.charAt(i);
            chr = chr - 64;
            result = result + (chr * multiplier);
            multiplier = multiplier * 26;
        }
        if (result - 1 > Worksheet.MAX_COLUMN_NUMBER || result - 1 < Worksheet.MIN_COLUMN_NUMBER)
        {
            throw new RangeException("OutOfRangeException","The column number (" + Integer.toString(result - 1) + ") is out of range. Range is from " + Integer.toString(Worksheet.MIN_COLUMN_NUMBER) + " to " + Integer.toString(Worksheet.MAX_COLUMN_NUMBER) + " (" + Integer.toString((Worksheet.MAX_COLUMN_NUMBER + 1)) + " columns).");
        }        
        return result - 1;
    }
    
    /**
     * Gets the column address (A - XFD)
     * @param columnNumber Column number (zero-based)
     * @return Column address (A - XFD)
     * @throws RangeException Thrown if the passed column number is out of range
     */
    public static String resolveColumnAddress(int columnNumber)
    {
        if (columnNumber > Worksheet.MAX_COLUMN_NUMBER || columnNumber < Worksheet.MIN_COLUMN_NUMBER)
        {
            throw new RangeException("OutOfRangeException","The column number (" + Integer.toString(columnNumber) + ") is out of range. Range is from " + Integer.toString(Worksheet.MIN_COLUMN_NUMBER) + " to " + Integer.toString(Worksheet.MAX_COLUMN_NUMBER) + " (" + Integer.toString((Worksheet.MAX_COLUMN_NUMBER + 1)) + " columns).");
        }
        // A - XFD
        int j = 0;
        int k = 0;
        int l = 0;
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i <= columnNumber; i++)
        {
            if (j > 25)
            {
                k++;
                j = 0;
            }
            if (k > 25)
            {
                l++;
                k = 0;
            }
            j++;
        }
        if (l > 0) { sb.append((char)(l + 64)); }
        if (k > 0) { sb.append((char)(k + 64)); }
        sb.append((char)(j + 64));
        return sb.toString();
    }
}
