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
import ch.rabanti.nanoxlsx4j.exception.WorksheetException;
import ch.rabanti.nanoxlsx4j.style.BasicStyles;
import ch.rabanti.nanoxlsx4j.style.Style;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;



/**
 * Class representing a worksheet of a workbook
 */
public class Worksheet {
    
// ### C O N S T A N T S ###    
    /**
    * Default column width as constant
    */
    public static final float DEFAULT_COLUMN_WIDTH = 10f;
    /**
    * Default row height as constant
    */
    public static final float DEFAULT_ROW_HEIGHT = 15f;
    /**
     * Maximum column number (zero-based)
     */
    public static final int MAX_COLUMN_NUMBER = 16383;
    /**
    * Maximum column width as constant
    */
    public static final float MAX_COLUMN_WIDTH = 255f;
    /**
     * Maximum row number (zero-based)
     */
    public static final int MAX_ROW_NUMBER = 1048575;
    /**
    * Maximum row height as constant
    */
    public static final float MAX_ROW_HEIGHT = 409.5f;    
    /**
     * Minimum column number (zero-based)
     */
    public static final int MIN_COLUMN_NUMBER = 0;
    /**
     * Minimum column width as constant
     */
    public static final float MIN_COLUMN_WIDTH = 0f;
    /**
     * Minimum row number (zero-based)
     */
    public static final int MIN_ROW_NUMBER = 0;
    /**
     * Minimum row height as constant
     */
    public static final float MIN_ROW_HEIGHT = 0f;

// ### E N U M S ###
    /**
     * Enum to define the direction when using AddNextCell method
     */
    public enum CellDirection
    {
        /**
         * The next cell will be on the same row (A1,B1,C1...)
         */
        ColumnToColumn,
        /**
         * The next cell will be on the same column (A1,A2,A3...)
         */
        RowToRow      
    }
    /**
     * Enum to define the possible protection types when protecting a worksheet
     */
    public enum SheetProtectionValue
    {
        // sheet, // Is always on 1 if protected
        /**
         * If selected, the user can edit objects if the worksheets is protected
         */
        objects,
        /**
         * If selected, the user can edit scenarios if the worksheets is protected
         */
        scenarios,
        /**
         * If selected, the user can format cells if the worksheets is protected
         */
        formatCells,
        /**
         * If selected, the user can format columns if the worksheets is protected
         */
        formatColumns,
        /**
         * If selected, the user can format rows if the worksheets is protected
         */
        formatRows,
        /**
         * If selected, the user can insert columns if the worksheets is protected
         */
        insertColumns,
        /**
         * If selected, the user can insert rows if the worksheets is protected
         */
        insertRows,
        /**
         * If selected, the user can insert hyperlinks if the worksheets is protected
         */
        insertHyperlinks,
        /**
         * If selected, the user can delete columns if the worksheets is protected
         */
        deleteColumns,
        /**
         * If selected, the user can delete rows if the worksheets is protected
         */
        deleteRows,
        /**
         * If selected, the user can select locked cells if the worksheets is protected
         */
        selectLockedCells,
        /**
         * If selected, the user can sort cells if the worksheets is protected
         */
        sort,
        /**
         * If selected, the user can use auto filters if the worksheets is protected
         */
        autoFilter,
        /**
         * If selected, the user can use pivot tables if the worksheets is protected
         */
        pivotTables,
        /**
         * If selected, the user can select unlocked cells if the worksheets is protected
         */
        selectUnlockedCells
    }
    
// ### P R I V A T E  F I E L D S ###    
    private Style activeStyle;
    private Range autoFilterRange;
    private Map<String, Cell> cells;
    private Map<Integer, Column> columns;
    private CellDirection currentCellDirection;
    private int currentColumnNumber;
    private int currentRowNumber;
    private float defaultColumnWidth;
    private float defaultRowHeight;
    private Map<Integer, Boolean> hiddenRows;
    private Map<String, Range> mergedCells;
    private Map<Integer, Float> rowHeights;
    private Range selectedCells;
    private int sheetID;
    private String sheetName;
    private String sheetProtectionPassword;
    private List<SheetProtectionValue> sheetProtectionValues;
    private boolean useSheetProtection;
    private boolean useActiveStyle;
    private Workbook workbookReference;
    

// ### G E T T E R S  &  S E T T E R S ###    
   
    /**
     * Sets the column auto filter within the defined column range
     * @param range Range to apply auto filter on. The range could be 'A1:C10' for instance. The end row will be recalculated automatically when saving the file
     * @throws RangeException Thrown if the passed range out of range
     * @exception FormatException Thrown if the passed range is malformed
     */
    public void setAutoFilterRange(String range)
    {
        this.autoFilterRange = Cell.resolveCellRange(range);
        recalculateAutoFilter();
        recalculateColumns();
    }
    /**
     * Gets the range of the auto filter. If null, no auto filters are applied
     * @return Range of auto filter
     */
    public Range getAutoFilterRange() {
        return autoFilterRange;
    }
    /**
     * Gets the cells of the worksheet as map with the cell address as key and the cell object as value
     * @return List of Cell objects
     */
    public Map<String, Cell> getCells() {
        return cells;
    }
    /**
     * Gets all columns with non-standard properties, like auto filter applied or a special width as map with the zero-based column index as key and the column object as value
     * @return map of columns
     */
    public Map<Integer, Column> getColumns() {
        return columns;
    }
    
    /**
     * Gets the direction when using AddNextCell method
     * @return Cell direction
     */
    public CellDirection getCurrentCellDirection() {
        return currentCellDirection;
    }    
    /**
     * Sets the direction when using AddNextCell method
     * @param currentCellDirection Cell direction
     */
    public void setCurrentCellDirection(CellDirection currentCellDirection) {
        this.currentCellDirection = currentCellDirection;
    }
    /**
     * Sets the current column number (zero based)
     * @param columnNumber Column number (zero based)
     * @throws RangeException Thrown if the number is out of the valid range. Range is from 0 to 16383 (16384 columns)
     */
    public void setCurrentColumnNumber(int columnNumber)
    {
        if (columnNumber > MAX_COLUMN_NUMBER || columnNumber < MIN_COLUMN_NUMBER)
        {
            throw new RangeException("OutOfRangeException","The column number (" + Integer.toString(columnNumber) + ") is out of range. Range is from "+ Integer.toString(MIN_COLUMN_NUMBER)+ " to "+ Integer.toString(MAX_COLUMN_NUMBER) +" ("+ Integer.toString(MAX_COLUMN_NUMBER + 1) +" columns).");
        }
        this.currentColumnNumber = columnNumber;
    }  
    /**
     * Sets the current row number (zero based)
     * @param rowNumber Row number (zero based)
     * @throws RangeException Thrown if the number is out of the valid range. Range is from 0 to 1048575 (1048576 rows)
     */
    public void setCurrentRowNumber(int rowNumber)
    {
        if (rowNumber > MAX_ROW_NUMBER || rowNumber < MIN_ROW_NUMBER)
        {
            throw new RangeException("OutOfRangeException","The row number (" + Integer.toString(rowNumber) + ") is out of range. Range is from "+ Integer.toString(MIN_ROW_NUMBER) +" to "+ Integer.toString(MAX_ROW_NUMBER) +" ("+ Integer.toString(MAX_ROW_NUMBER + 1)+" rows).");
        }
        this.currentRowNumber = rowNumber;
    }
    
    /**
     * Gets the current column number (zero based)
     * @return Column number (zero-based)
     */
    public int getCurrentColumnNumber()
    {
        return this.currentColumnNumber;
    }

    /**
     * Gets the current row number (zero based)
     * @return Row number (zero-based)
     */
    public int getCurrentRowNumber()
    {
        return this.currentRowNumber;
    }     
    
    /**
     * Gets the default column width
     * @return Default column width
     */
    public float getDefaultColumnWidth() {
        return defaultColumnWidth;
    } 
    /**
     * Sets the default column width
     * @param defaultColumnWidth Default column width
     * @throws RangeException Throws a RangeException exception if the passed width is out of range (set)
     */
    public void setDefaultColumnWidth(float defaultColumnWidth) {
        if (defaultRowHeight < MIN_COLUMN_WIDTH || defaultRowHeight > MAX_COLUMN_WIDTH)
        {
            throw new RangeException("OutOfRangeException","The passed default row height is out of range (" + Float.toString(MIN_COLUMN_WIDTH) + " to " + Float.toString(MAX_COLUMN_WIDTH) + ")");
        }
        this.defaultColumnWidth = defaultColumnWidth;
    }    
    /*
    * Gets the default Row height
    * @return Default Row height
    */
    public float getDefaultRowHeight() {
        return defaultRowHeight;
    }
    /**
     * Sets the default Row height
     * @param defaultRowHeight Default Row height
     * @throws RangeException Throws a RangeException exception if the passed height is out of range (set)
     */
    public void setDefaultRowHeight(float defaultRowHeight) {
        if (defaultRowHeight < MIN_ROW_HEIGHT || defaultRowHeight > MAX_ROW_HEIGHT)
        {
            throw new RangeException("OutOfRangeException","The passed default row height is out of range (" + Float.toString(MIN_ROW_HEIGHT) + " to " + Float.toString(MAX_ROW_HEIGHT) + ")");
        }
        this.defaultRowHeight = defaultRowHeight;
    }
    /**
     * Gets the hidden rows as map with the zero-based row number as key and a boolean as value. True indicates hidden, false visible. Entries with the value false are not affecting the worksheet. These entries can be removed<br>
     * @return Map with hidden rows
     */
    public Map<Integer, Boolean> getHiddenRows() {
        return hiddenRows;
    }
    /**
     * Gets the merged cells (only references) as map with the cell address as key and the range object as value
     * @return Hashmap with merged cell references
     */
    public Map<String, Range> getMergedCells() {
        return mergedCells;
    }

    /**
     * Gets defined row heights as map with the zero-based row number as key and the height (float from 0 to 409.5) as value
     * @return Map of row heights
     */
    public Map<Integer, Float> getRowHeights() {
        return rowHeights;
    }
    /**
     * Gets the range of selected cells of this worksheet. Null if no cells are selected
     * @return Cell range of the selected cells
     */
    public Range getSelectedCells() {
        return selectedCells;
    }
    /**
     * Sets the selected cells on this worksheet
     * @param range Cell range to select
     * @throws RangeException Thrown if the passed range out of range
     * @exception FormatException Thrown if the passed range is malformed
     */
    public void setSelectedCells(String range)
    {
        this.selectedCells = Cell.resolveCellRange(range);
    }
    /**
     * Sets the selected cells on this worksheet
     * @param range Cell range to select
     */
    public void setSelectedCells(Range range)
    {
        this.selectedCells = range;
    }
    /**
     * Gets the internal ID of the worksheet
     * @return Worksheet ID
     */
    public int getSheetID() {
        return sheetID;
    }
    /**
     * Sets the internal ID of the worksheet
     * @param sheetID Worksheet ID
     */
    public void setSheetID(int sheetID) {
        this.sheetID = sheetID;
    }
    /**
     * Gets the name of the sheet
     * @return Name of the sheet
     */
    public String getSheetName() {
        return sheetName;
    }
    
    /**
     * Gets whether the worksheet is protected
     * @return If true, the worksheet is protected
     */
    public boolean isUseSheetProtection() {
        return useSheetProtection;
    }
    /**
     * Sets whether the worksheet is protected
     * @param useSheetProtection If true, the worksheet is protected
     */
    public void setUseSheetProtection(boolean useSheetProtection) {
        this.useSheetProtection = useSheetProtection;
    }
    
    /**
     * Gets the password used for sheet protection
     * @return Password (UTF-8)
     */
    public String getSheetProtectionPassword() {
        return sheetProtectionPassword;
    }
    /**
     * Sets or removes the password for worksheet protection. If set, UseSheetProtection will be also set to true
     * @param password Password (UTF-8) to protect the worksheet. If the password is null or empty, no password will be used
     */
    public void setSheetProtectionPassword(String password)
    {
        if (Helper.isNullOrEmpty(password) == true)
        {
            this.sheetProtectionPassword = null;
        }
        else
        {
            this.sheetProtectionPassword = password;
            this.useSheetProtection = true;
        }
    }
    /**
     * Gets the list of SheetProtectionValues. These values define the allowed actions if the worksheet is protected
     * @return List of SheetProtectionValues
     */
    public List<SheetProtectionValue> getSheetProtectionValues() {
        return sheetProtectionValues;
    }
    /**
     * Gets the Reference to the parent Workbook
     * @return Workbook reference
     */
    public Workbook getWorkbookReference() {
        return workbookReference;
    }
    /**
     * Sets the Reference to the parent Workbook
     * @param workbookReference Workbook reference
     */
    public void setWorkbookReference(Workbook workbookReference) {
        this.workbookReference = workbookReference;
    }    

    // ### C O N S T R U C T O R S ###
    /**
     * Default constructor. A worksheet created with this constructor cannot be used to assign styles to a cell. This will cause an exception unless a reference to the workbook was set
     * @throws StyleException Thrown if a style is added to the worksheet without w workbook reference
     */
    public Worksheet()
    {
        init();
    }
    /**
     * Constructor with workbook reference
     * @param reference Reference to the parent Workbook
     * @throws FormatException Thrown if the name contains illegal characters or is too long
     */
    public Worksheet(Workbook reference)
    {
        init();
        this.workbookReference = reference;
    }
    /**
     * Constructor with workbook reference, name and sheet ID
     * @param name Name of the worksheet
     * @param id ID of the worksheet (for internal use)
     * @param reference Reference to the parent Workbook
     * @throws FormatException Thrown if the name contains illegal characters or is too long
     */
    public Worksheet(String name, int id, Workbook reference)
    {
        init();
        setSheetName(name);
        this.sheetID = id;
        this.workbookReference = reference;
    }    
    
// ### M E T H O D S ###
       
// ### M E T H O D S  -  A D D N E X T C E L L ###
    
    /**
     * Adds an object to the next cell position. If the type of the value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int, double, float, long, Date, boolean. All other types will be casted into a String using the default toString() method
     * @param value Unspecified value to insert
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCell(Object value)
    {
        addNextCell(castValue(value, this.currentColumnNumber, this.currentRowNumber),true,null);
    }

     /**
     * Adds an object to the next cell position. If the type of the value does not match with one of the supported data types, it will be casted to a String.A prepared object of the type Cell will not be casted but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int, double, float, long, Date, boolean. All other types will be casted into a String using the default toString() method
     * @param value Unspecified value to insert
     * @param style Style object to apply on this cell
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCell(Object value, Style style)
    {
        addNextCell(castValue(value, this.currentColumnNumber, this.currentRowNumber),true,style);
    }
   
    /**
     * Method to insert a generic cell to the next cell position
     * @param cell Cell object to insert
     * @param incremental If true, the address value (row or column) will be incremented, otherwise not
     * @param style If not null, the defined style will be applied to the cell, otherwise no style or the default style will be applied
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    private void addNextCell(Cell cell, boolean incremental, Style style)
    {
        cell.setWorksheetReference(this);
        if (this.activeStyle != null && this.useActiveStyle == true && style == null)
        {
            cell.setStyle(this.activeStyle);
        }
        else if (style != null)
        {
            cell.setStyle(style);
        }
        else if (style == null && cell.getDataType() == Cell.CellType.DATE)
        {
            cell.setStyle(BasicStyles.DateFormat());
        }
        String address = cell.getCellAddress();
        this.cells.put(address, cell);
        if (incremental == true)
        {
            if (this.getCurrentCellDirection() == CellDirection.ColumnToColumn)
            {
                this.currentColumnNumber++;
            }
            else
            {
                this.currentRowNumber++;
            }
        }
        else
        {
            if (this.getCurrentCellDirection() == CellDirection.ColumnToColumn)
            {
                this.currentColumnNumber = cell.getColumnNumber() + 1;
                this.currentRowNumber = cell.getRowNumber();
            }
            else
            {
                this.currentColumnNumber = cell.getColumnNumber();
                this.currentRowNumber = cell.getRowNumber() + 1;
            }
        }
    }
    
// ### M E T H O D S  -  A D D C E L L ###
   
    /**
     * Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int, double, float, long, Date, boolean. All other types will be casted into a String using the default toString() method
     * @param value Unspecified value to insert
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws StyleException Thrown if the active style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Object value, int columnAddress, int rowAddress)
    {
        addNextCell(castValue(value, columnAddress, rowAddress), false, null);
    }
    
    /**
     * Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int, double, float, long, Date, boolean. All other types will be casted into a String using the default toString() method
     * @param value Unspecified value to insert
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @param style Style to apply on the cell
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Object value, int columnAddress, int rowAddress, Style style)
    {
        addNextCell(castValue(value, columnAddress, rowAddress), false, style);

    }    
    
    /**
     * Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int, double, float, long, Date, boolean. All other types will be casted into a String using the default toString() method
     * @param value Unspecified value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws FormatException Thrown if the passed address is malformed
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Object value, String address)
    {
        Address adr = Cell.resolveCellCoordinate(address);
        addCell(value, adr.Column, adr.Row);
    }      
    
    /**
     * Adds an object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int, double, float, long, Date, boolean. All other types will be casted into a String using the default toString() method
     * @param value Unspecified value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @param style Style to apply on the cell
     * @throws FormatException Thrown if the passed address is malformed
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Object value, String address, Style style)
    {
        Address adr = Cell.resolveCellCoordinate(address);
        addCell(value, adr.Column, adr.Row, style);
    }      
    
// ### M E T H O D S  -  A D D C E L L F O R M U L A ###
    
    /**
     * Adds a cell formula as string to the defined cell address
     * @param formula Formula to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws FormatException Thrown if the passed address is malformed
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellFormula(String formula, String address)
    {
        Address adr = Cell.resolveCellCoordinate(address);
        Cell c = new Cell(formula, Cell.CellType.FORMULA, adr.Column, adr.Row, this);
        addNextCell(c, false, null);
    }
    
    /**
     * Adds a cell formula as string to the defined cell address
     * @param formula Formula to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @param style Style to apply on the cell
     * @throws FormatException Thrown if the passed address is malformed
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellFormula(String formula, String address, Style style)
    {
        Address adr = Cell.resolveCellCoordinate(address);
        Cell c = new Cell(formula, Cell.CellType.FORMULA, adr.Column, adr.Row, this);
        addNextCell(c, false, style);
    }    
    
    /**
     * Adds a cell formula as string to the defined cell address
     * @param formula Formula to insert
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellFormula(String formula, int columnAddress, int rowAddress)
    {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, columnAddress, rowAddress, this);
        addNextCell(c, false, null);
    }     
    
    /**
     * Adds a cell formula as string to the defined cell address
     * @param formula Formula to insert
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @param style Style to apply on the cell
     * @throws StyleException Thrown if the passed style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellFormula(String formula, int columnAddress, int rowAddress, Style style)
    {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, columnAddress, rowAddress, this);
        addNextCell(c, false, style);
    }
    
   /**
     * Adds a formula as string to the next cell position
     * @param formula Formula to insert
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCellFormula(String formula)
    {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, this.currentColumnNumber, this.currentRowNumber, this);
        addNextCell(c, true, null);
    }  
    
    /**
     * Adds a formula as string to the next cell position
     * @param formula Formula to insert
     * @param style Style to apply on the cell
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCellFormula(String formula, Style style)
    {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, this.currentColumnNumber, this.currentRowNumber, this);
        addNextCell(c, true, style);
    }   
    
    
// ### M E T H O D S  -  A D D C E L L R A N G E ###

    /**
     * Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int, double, float, long, Date, boolean. All other types will be casted into a String using the default toString() method
     * @param values List of unspecified objects to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellRange(List<Object> values, Address startAddress, Address endAddress)
    {
        addCellRangeInternal(values, startAddress, endAddress, null);
    }  
    
   /**
     * Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int, double, float, long, Date, boolean. All other types will be casted into a String using the default toString() method
     * @param values List of unspecified objects to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @param style Style to apply on the all cells of the range
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellRange(List<Object> values, Address startAddress, Address endAddress, Style style)
    {
        addCellRangeInternal(values, startAddress, endAddress, style);
    }     

    /**
     * Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted<br>
     * The data types in the passed list can be mixed. Recognized are the following data types: Cell (prepared object), String, int, double, float, long, Date, boolean. All other types will be casted into a String using the default toString() method
     * @param values List of unspecified objects to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellRange(List<Object> values, String cellRange)
    {
        Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress, null);
    }       
    
    /**
     * Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String. A prepared object of the type Cell will not be casted but adjusted<br>
     * The data types in the passed list can be mixed. Recognized are the following data types: Cell (prepared object), String, int, double, float, long, Date, boolean. All other types will be casted into a String using the default toString() method
     * @param values List of unspecified objects to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @param style Style to apply on the all cells of the range
     * @throws FormatException Thrown if the passed address is malformed
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellRange(List<Object> values, String cellRange, Style style)
    {
        Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress, style);
    }      
    
    /**
     * Internal function to add a generic list of value to the defined cell range
     * @param <T> Data type of the generic value list
     * @param values List of values
     * @param startAddress Start address
     * @param endAddress End address
     * @param style Style to apply on the all cells of the range. If null, no style or the default style will be applied
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    private <T> void addCellRangeInternal(List<T> values, Address startAddress, Address endAddress, Style style)
    {
        List<Address> addresses = Cell.getCellRange(startAddress, endAddress);
        if (values.size() != addresses.size())
        {
            throw new RangeException("MalformedRangeException","The number of passed values (" + Integer.toString(values.size()) + ") differs from the number of cells within the range (" + Integer.toString(addresses.size()) + ")");
        }
        List<Cell> list = Cell.convertArray(values);
        int len = values.size();
        for(int i = 0; i < len; i++)
        {
            list.get(i).setRowNumber(addresses.get(i).Row);
            list.get(i).setColumnNumber(addresses.get(i).Column);
            list.get(i).setWorksheetReference(this);
            addNextCell(list.get(i), false, style);
        }
    }
    
// ### M E T H O D S  -  R E M O V E C E L L ###
    
    /**
     * Removes a previous inserted cell at the defined address
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @return Returns true if the cell could be removed (existed), otherwise false (did not exist)
     * @throws RangeException Thrown if the resolved cell address is out of range
     */
    public boolean removeCell(int columnAddress, int rowAddress)
    {
        String address = Cell.resolveCellAddress(columnAddress, rowAddress);
        if (this.cells.containsKey(address))
        {
            this.cells.remove(address);
            return true;
        }
        else
        {
            return false;
        }
    }
    /**
     * Removes a previous inserted cell at the defined address
     * @param address Cell address in the format A1 - XFD1048576
     * @return Returns true if the cell could be removed (existed), otherwise false (did not exist)
     * @throws RangeException Thrown if the resolved cell address is out of range
     * @throws FormatException Thrown if the passed address is malformed
     */
    public boolean removeCell(String address)
    {
        Address adr = Cell.resolveCellCoordinate(address);
        return removeCell(adr.Column, adr.Row);
    }    
    
// ### C O M M O N   M E T H O D S ###   
    
    /**
     * Method to add allowed actions if the worksheet is protected. If one or more values are added, UseSheetProtection will be set to true
     * @param typeOfProtection Allowed action on the worksheet or cells
     */
    public void addAllowedActionOnSheetProtection(SheetProtectionValue typeOfProtection)
    {
        if (this.sheetProtectionValues.contains(typeOfProtection) == false)
        {
            if (typeOfProtection == SheetProtectionValue.selectLockedCells && this.sheetProtectionValues.contains(SheetProtectionValue.selectUnlockedCells) == false)
            {
                this.sheetProtectionValues.add(SheetProtectionValue.selectUnlockedCells);
            }
            this.sheetProtectionValues.add(typeOfProtection);
            this.setUseSheetProtection(true);
        }
    }  
      
    /**
     * Sets the defined column as hidden
     * @param columnNumber Column number to hide on the worksheet
     * @throws RangeException Thrown if the passed row number was out of range
     */
    public void addHiddenColumn(int columnNumber)
    {
        setColumnHiddenState(columnNumber, true);
    }   
    /**
     * Sets the defined column as hidden
     * @param columnAddress Column address to hide on the worksheet
     * @throws RangeException Thrown if the passed row number was out of range
     */
    public void addHiddenColumn(String columnAddress)
    {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnHiddenState(columnNumber, true);
    }
    /**
     * Sets the defined row as hidden
     * @param rowNumber Row number to hide on the worksheet
     * @throws RangeException Thrown if the passed column number was out of range
     */
    public void addHiddenRow(int rowNumber)
    {
        setRowHiddenState(rowNumber, true);
    }    
   
    /**
     * Method to cast a value or align an object of the type Cell to the context of the worksheet
     * @param value Unspecified value or object of the type Cell
     * @param column Column index
     * @param row Row index
     * @return Cell object
     */
    private Cell castValue(Object value, int column, int row)
    {
        Cell c;
        if (value instanceof Cell)
        {
            c = (Cell)value;
            c.setWorksheetReference(this);
            c.setCellAddress2(new Address(column, row));
        }
        else
        {
            c = new Cell(value, Cell.CellType.DEFAULT, column, row, this);
        }
        return c;
    }
    
    /**
     * Clears the active style of the worksheet. All later added cells will contain no style unless another active style is set
     */
    public void clearActiveStyle()
    {
        this.useActiveStyle = false;
    }
    
    /**
     * Gets the cell of the specified address
     * @param address Address of the cell
     * @return Cell object
     * @throws WorksheetException Throws a WorksheetException if the cell was not found on the cell table of this worksheet
     */
    public Cell getCell(Address address)
    {
        if (this.cells.containsKey(address.getAddress()) == false)
        {
            throw new WorksheetException("CellNotFoundException", "The cell with the address " + address.getAddress() + " does not exist in this worksheet");
        }
        return this.cells.get(address.getAddress());
    }
    
    /**
     * Gets the cell of the specified column and row number (zero-based)
     * @param columnNumber Column address of the cell
     * @param rowNumber Row address of the cell
     * @return Cell object
     * @throws WorksheetException Throws a WorksheetException if the cell was not found on the cell table of this worksheet
     */
    public Cell getCell(int columnNumber, int rowNumber)
    {
        return getCell(new Address(columnNumber, rowNumber));
    }    
    
    /**
     * Set the current cell address
     * @param columnAddress Column number (zero based)
     * @param width Row number (zero based)
     * @throws RangeException Thrown if the address is out of the valid range. Range is from 0 to 16383 (16384 columns)
     */
    public void setColumnWidth(String columnAddress, float width)
    {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnWidth(columnNumber, width);
    } 
    
    /**
     * Set the current cell address
     * @param address Cell address in the format A1 - XFD1048576
     * @throws RangeException Thrown if the address is out of the valid range. Range is for rows from 0 to 1048575 (1048576 rows) and for columns from 0 to 16383 (16384 columns)
     * @throws FormatException Thrown if the passed address is malformed
     */
    public void setCurrentCellAddress(String address)
    {
        Address adr = Cell.resolveCellCoordinate(address);
        setCurrentCellAddress(adr.Column, adr.Row);
    }  
   
    /**
     * Sets the name of the sheet
     * @param sheetName Name of the sheet
     * @throws FormatException Thrown if the name contains illegal characters or is longer than 31 characters
     */
    public void setSheetName(String sheetName)
    {
        if (Helper.isNullOrEmpty(sheetName))
        {
            throw new FormatException("FormatException","The sheet name must be between 1 and 31 characters");
        }
        if (sheetName.length() > 31)
        {
            throw new FormatException("FormatException","The sheet name must be between 1 and 31 characters");
        }
        Pattern pattern = Pattern.compile("[\\[\\]\\*\\?/\\\\]");
        Matcher mx = pattern.matcher(sheetName);
        if (mx.groupCount() > 0)
        {
            throw new FormatException("FormatException","The sheet name must must not contain the characters [  ]  * ? / \\ ");
        }
        this.sheetName = sheetName;
    }
    
    /**
     * Sets the name of the sheet
     * @param sheetName Name of the sheet
     * @param sanitize If true, the filename will be sanitized automatically according to the specifications of Excel
     * @throws WorksheetException Thrown if no workbook is referenced. This information is necessary to determine whether the name already exists
     */
    public void setSheetName(String sheetName, boolean sanitize)
    {    
       if (this.workbookReference == null)
       {
           throw new WorksheetException("MissingReferenceException", "The worksheet name cannot be sanitized because no workbook is referenced");
       }
       this.sheetName =  Worksheet.sanitizeWorksheetName(sheetName, this.workbookReference);         
    }
    
    /**
     * Moves the current position to the next column
     */
    public void goToNextColumn()
    {
        this.currentColumnNumber++;
        this.currentRowNumber = 0;
    }
    
    /**
     * Moves the current position to the next column with the number of cells to move
     * @param numberOfColumns Number of columns to move
     */
    public void goToNextColumn(int numberOfColumns)
    {
        for(int i = 0; i < numberOfColumns; i++)
        {
            goToNextColumn();
        }
    }
    
    /**
     * Moves the current position to the next row (use for a new line)
     */
    public void goToNextRow()
    {
        this.currentRowNumber++;
        this.currentColumnNumber = 0;
    }
    
    /**
     * Moves the current position to the next row with the number of cells to move (use for a new line)
     * @param numberOfRows Number of rows to move
     */
    public void goToNextRow(int numberOfRows)
    {
        for(int i = 0; i < numberOfRows; i++)
        {
            goToNextRow();
        }
    }
    
    /**
     * Init method for constructors
     */
    private void init()
    {
        this.currentCellDirection = CellDirection.ColumnToColumn;
        this.cells = new HashMap<>();
        this.currentRowNumber = 0;
        this.currentColumnNumber = 0;
        this.defaultColumnWidth = DEFAULT_COLUMN_WIDTH;
        this.defaultRowHeight = DEFAULT_ROW_HEIGHT;
        this.rowHeights = new HashMap<>();
        this.activeStyle = null;
        this.workbookReference = null;
        this.mergedCells = new HashMap<>();    
        this.sheetProtectionValues = new ArrayList<>();
        this.hiddenRows = new HashMap<>();
        this.columns = new HashMap<>();
    }
    
    /**
     * Merges the defined cell range
     * @param cellRange Range to merge
     * @return Returns the validated range of the merged cells (e.g. 'A1:B12')
     */
    public String mergeCells(Range cellRange)
    {
        return mergeCells(cellRange.StartAddress, cellRange.EndAddress);
    }

    /**
     * Merges the defined cell range
     * @param cellRange Range to merge (e.g. 'A1:B12')
     * @return Returns the validated range of the merged cells (e.g. 'A1:B12')
     * @throws ch.rabanti.nanoxlsx4j.exception.FormatException Thrown if the passed address is malformed
     */
    public String mergeCells(String cellRange)
    {
        Range range = Cell.resolveCellRange(cellRange);
        return mergeCells(range.StartAddress, range.EndAddress);
    }    
    
    /**
     * Merges the defined cell range
     * @param startAddress Start address of the merged cell range
     * @param endAddress End address of the merged cell range
     * @return Returns the validated range of the merged cells (e.g. 'A1:B12')
     */
    public String mergeCells(Address startAddress, Address endAddress)
    {
        String key = startAddress.toString() + ":" + endAddress.toString();
        Range value = new Range(startAddress, endAddress);
        if (this.mergedCells.containsKey(key) == false)
        {
            this.mergedCells.put(key, value);
        }
        return key;
    }    
    /**
     * Method to recalculate the auto filter (columns) of this worksheet. This is an internal method. There is no need to use it. It must be public to require access from the LowLevel class
     */
    public void recalculateAutoFilter()
    {
        if (this.autoFilterRange == null) { return; }
        int start = this.autoFilterRange.StartAddress.Column;
        int end = this.autoFilterRange.EndAddress.Column;
        int endRow = 0;
        for(Map.Entry<String, Cell> item  : this.getCells().entrySet())
        {
            if (item.getValue().getColumnNumber() < start || item.getValue().getColumnNumber() > end) { continue; }
            if (item.getValue().getRowNumber() > endRow) {endRow = item.getValue().getRowNumber();}
        }
        Column c;
        for(int i = start; i <= end; i++)
        {
            if (this.columns.containsKey(i) == false)
            {
                c = new Column(i);
                c.setAutoFilter(true);
                this.columns.put(i, c);
            }
            else
            {
                this.getColumns().get(i).setAutoFilter(true);
            }
        }
        this.autoFilterRange = new Range(new Address(start, 0), new Address(end, endRow));
    }
    /**
     * Method to recalculate the collection of columns of this worksheet. This is an internal method. There is no need to use it. It must be public to require access from the LowLevel class
     */
    public void recalculateColumns()
    {
        ArrayList<Integer> columnsToDelete = new ArrayList<>();
        for(Map.Entry<Integer, Column> col  : this.getColumns().entrySet())
        {
            if (col.getValue().hasAutoFilter() == false && col.getValue().isHidden() == false && col.getValue().getWidth() == Worksheet.DEFAULT_COLUMN_WIDTH)
            {
                columnsToDelete.add(col.getKey());
            }
        }
        for(Iterator<Integer> index = columnsToDelete.iterator(); index.hasNext(); )
        {
            this.columns.remove(index.next());
        }
    }
    /**
     * Removes auto filters from the worksheet
     */
    public void removeAutoFilter()
    {
        this.autoFilterRange = null;
    }
    
    /**
     * Sets a previously defined, hidden column as visible again
     * @param columnNumber Column number to make visible again
     * @throws RangeException Thrown if the passed row number was out of range
     */
    public void removeHiddenColumn(int columnNumber)
    {
        setColumnHiddenState(columnNumber, false);
    }
    
    /**
     * Sets a previously defined, hidden column as visible again
     * @param columnAddress Column address to make visible again
     * @throws RangeException Thrown if the passed row number was out of range
     */
    public void removeHiddenColumn(String columnAddress)
    {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnHiddenState(columnNumber, false);
    }
    /**
     * Sets a previously defined, hidden row as visible again
     * @param rowNumber Row number to hide on the worksheet
     * @throws RangeException Thrown if the passed column number was out of range
     */
    public void removeHiddenRow(int rowNumber)
    {
        setRowHiddenState(rowNumber, false);
    }
    /**
     * Removes the defined merged cell range
     * @param range Cell range to remove the merging
     * @throws RangeException Thrown if the passed cell range was not merged earlier
     * @throws FormatException Thrown if the passed address is malformed
     */
    public void removeMergedCells(String range)
    {
        range = range.toUpperCase();
        if (this.mergedCells.containsKey(range) == false)
        {
            throw new RangeException("MissingReferenceException","The cell range " + range + " was not found in the list of merged cell ranges");
        }
        else
        {
            List<Address> addresses = Cell.getCellRange(range);
            Cell cell;
            //foreach(Address address in addresses)
            for(int i = 0; i < addresses.size(); i++)
            {
                if (this.cells.containsKey(addresses.get(i).toString()))
                {
                    cell = this.cells.get(addresses.get(i).toString());
                    cell.setDataType(Cell.CellType.DEFAULT); // resets the type
                    if (cell.getValue() == null)
                    {
                        cell.setValue("");
                    }
                }
            }
            this.mergedCells.remove(range);
        }
    }
    /**
     * Removes the cell selection of this worksheet
     */
    public void removeSelectedCells()
    {
        this.selectedCells = null;
    }
    /**
     * Sets the active style of the worksheet. This style will be assigned to all later added cells
     * @param style Style to set as active style
     * @throws StyleException Thrown if the worksheet has no workbook referenced when trying to set the active style
     */
    public void setActiveStyle(Style style)
    {
        this.useActiveStyle = true;
        this.activeStyle = style;
    }
    
    /**
     * Sets the column auto filter within the defined column range
     * @param startColumn Column number with the first appearance of an auto filter drop down
     * @param endColumn Column number with the last appearance of an auto filter drop down
     * @throws RangeException Thrown if one of the passed column numbers are out of range
     */
    public void setAutoFilter(int startColumn, int endColumn)
    {
        String start = Cell.resolveCellAddress(startColumn, 0);
        String end = Cell.resolveCellAddress(endColumn, 0);
        if (endColumn < startColumn)
        {
            setAutoFilterRange(end + ":" + start);
        }
        else
        {
            setAutoFilterRange(start + ":" + end);
        }
    }
    /**
     * Sets the defined column as hidden or visible
     * @param columnNumber Column number to hide on the worksheet
     * @param state If true, the column will be hidden, otherwise be visible
     * @throws RangeException Thrown if the passed row number was out of range
     */
    private void setColumnHiddenState(int columnNumber, boolean state)
    {
        if (columnNumber > MAX_COLUMN_NUMBER || columnNumber < 0)
        {
            throw new RangeException("OutOfRangeException","The column number (" + Integer.toString(columnNumber) + ") is out of range. Range is from 0 to "+ Integer.toString(MAX_COLUMN_NUMBER) +" ("+ Integer.toString(MAX_COLUMN_NUMBER + 1) +" columns).");
        }
        if (this.columns.containsKey(columnNumber) && state == true)
        {
            this.columns.get(columnNumber).setHidden(state);
        }
        else if (state == true)
        {
            Column c = new Column(columnNumber);
            c.setHidden(state);
            this.columns.put(columnNumber, c);
        }
    } 
    /**
     * Sets the width of the passed column number (zero-based)
     * @param columnNumber Column number (zero-based, from 0 to 16383)
     * @param width Width from 0 to 255.0
     * @throws RangeException Thrown if the address is out of the valid range. Range is from 0 to 16383 (16384 columns)
     */
    public void setColumnWidth(int columnNumber, float width)
    {
        if (columnNumber > MAX_COLUMN_NUMBER || columnNumber < MIN_COLUMN_NUMBER)
        {
            throw new RangeException("OutOfRangeException","The column number (" + Integer.toString(columnNumber) + ") is out of range. Range is from "+ Integer.toString(MIN_COLUMN_NUMBER)+ " to "+ Integer.toString(MAX_COLUMN_NUMBER) +" ("+ Integer.toString(MAX_COLUMN_NUMBER + 1) +" columns).");
        }
        if (width < MIN_COLUMN_WIDTH || width > MAX_COLUMN_WIDTH)
        {
            throw new RangeException("OutOfRangeException","The column width (" + Float.toString(width) + ") is out of range. Range is from "+ Float.toString(MIN_COLUMN_WIDTH)+ " to "+ Float.toString(MAX_COLUMN_WIDTH)+ " (chars).");
        }
        if (this.columns.containsKey(columnNumber))
        {
            this.columns.get(columnNumber).setWidth(width);
        }
        else
        {
            Column c = new Column(columnNumber);
            c.setWidth(width);
            this.columns.put(columnNumber, c);
        }
    }    
    /**
     * Set the current cell address
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws RangeException Thrown if the address is out of the valid range. Range is for rows from 0 to 1048575 (1048576 rows) and for columns from 0 to 16383 (16384 columns)
     */
    public void setCurrentCellAddress(int columnAddress, int rowAddress)
    {
        setCurrentColumnNumber(columnAddress);
        setCurrentRowNumber(rowAddress);
    }
    /**
     * Sets the height of the passed row number (zero-based)
     * @param rowNumber Row number (zero-based, 0 to 1048575)
     * @param height Height from 0 to 409.5
     * @throws RangeException Thrown if the address is out of the valid range. Range is from 0 to 1048575 (1048576 rows)
     */
    public void setRowHeight(int rowNumber, float height)
    {
        if (rowNumber > MAX_ROW_NUMBER || rowNumber < MIN_ROW_NUMBER)
        {
            throw new RangeException("OutOfRangeException","The row number (" + Integer.toString(rowNumber) + ") is out of range. Range is from "+ Integer.toString(MIN_ROW_NUMBER) +" to "+ Integer.toString(MAX_ROW_NUMBER) +" ("+ Integer.toString(MAX_ROW_NUMBER + 1)+" rows).");
        }       
        if (height < 0 || height > 409.5)
        {
            throw new RangeException("OutOfRangeException","The row height (" + Float.toString(height) + ") is out of range. Range is from 0 to 409.5 (equals 546px).");
        }
        this.rowHeights.put(rowNumber, height);
    }
    /**
     * Sets the defined row as hidden or visible
     * @param rowNumber Row number to hide on the worksheet
     * @param state If true, the row will be hidden, otherwise visible
     * @throws RangeException Thrown if the passed column number was out of range
     */
    private void setRowHiddenState(int rowNumber, boolean state)
    {
        if (rowNumber > MAX_ROW_NUMBER || rowNumber < MIN_ROW_NUMBER)
        {
            throw new RangeException("OutOfRangeException","The row number (" + Integer.toString(rowNumber) + ") is out of range. Range is from "+ Integer.toString(MIN_ROW_NUMBER) +" to "+ Integer.toString(MAX_ROW_NUMBER) +" ("+ Integer.toString(MAX_ROW_NUMBER + 1)+" rows).");
        }
        if (this.hiddenRows.containsKey(rowNumber))
        {
            if (state == true)
            {
                this.hiddenRows.put(rowNumber, state);
            }
            else
            {
                this.hiddenRows.remove(rowNumber);
            }
        }
        else if (state == true)
        {
            this.hiddenRows.put(rowNumber, state);
        }
    }
    
    /**
     * Sets the selected cells on this worksheet
     * @param startAddress Start address of the range
     * @param endAddress End address of the range
     */
    public void setSelectedCells(Address startAddress, Address endAddress)
    {
       this.selectedCells = new Range(startAddress, endAddress); 
    }
 
    //### S T A T I C   M E T H O D S
    /**
     * Sanitizes a worksheet name
     * @param input Name to sanitize
     * @param workbook Workbook reference
     * @return Name of the sanitized worksheet
     */
    public static String sanitizeWorksheetName(String input, Workbook workbook)
    {
        if (input == null){input = "Sheet1";}
        int len = input.length();
        if (len > 31) { len = 31; }
        else if (len == 0)
        {
            input = "Sheet1";
        }
        StringBuilder sb = new StringBuilder(31);
        char c;
        for(int i = 0; i < len; i++)
        {
            c = input.charAt(i);
            if (c == '[' || c == ']' || c == '*' || c == '?' || c == '\\' || c == '/')
            { sb.append('_');}
            else
            {  sb.append(c); }
        }
        String name = sb.toString();
        String originalName = name;
        int number = 1;
        while(true)
        {
            if (Worksheet.worksheetExists(name, workbook) == false) { break; } // OK
            if (originalName.length() + (number/10) >= 31)
            {
                name = originalName.substring(0,30-number/10) + Integer.toString(number);
            }
            else
            {
                name = originalName + Integer.toString(number);
            }
            number++;
        }
        return name;
    }
    /**
     * Checks whether a worksheet with the given name exists
     * @param name Name to check
     * @param workbook Workbook reference
     * @return True if the name exits, otherwise false
     */
    private static boolean worksheetExists(String name, Workbook workbook)
    {
        int len = workbook.getWorksheets().size();
        for(int i = 0; i < len; i++)
        {
            if (workbook.getWorksheets().get(i).getSheetName().equals(name))
            {
                return true;
            }
        }
        return false;
    }   
    
}
