/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import java.math.BigDecimal;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Objects;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.enums.Errors;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.internal.FeatureSet;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;
import ch.rabanti.nanoxlsx4j.styles.StyleRepository;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;

public class Cell implements Comparable<Cell> {
    // enums

    /**
     * Enum defines the basic data types of a cell
     */
    public enum CellType {
        /** Type for single characters and strings */
        STRING,
        /**
         * Type for all numeric types (long, integer, float, double, short, byte and decimal; signed and unsigned, if
         * available)
         */
        NUMBER,
        /**
         * Type for dates, represented by {@link Date}  (Note: Dates before 1900-01-01 and after 9999-12-31 are not
         * allowed)
         */
        DATE,
        /** Type for times (Note: Internally handled as OAdate, represented by {@link java.time.Duration}) */
        TIME,
        /** Type for boolean */
        BOOL,
        /** Type for Formulas (The cell will be handled differently) */
        FORMULA,
        /**
         * Type for empty cells. This type is only used for merged cells (all cells except the first of the cell range)
         */
        EMPTY,
        /**
         *
         * <p>Remarks: The preferred value for this type is an {@link ch.rabanti.nanoxlsx4j.enums.FormulaError}. A
         * formula whose cached result is an error remains a {@link CellType#FORMULA} cell and exposes the error through
         * {@link FormulaData#getCachedValue()} and {@link FormulaData#getCachedValueType()}. Explicitly changing a
         * formula cell to this type discards its formula metadata.
         * </p>
         */
        ERROR,
        /* Default Type, not specified */
        DEFAULT
    }

    /**
     * Enum for the referencing style of the address
     */
    public enum AddressType {
        /** Default behavior (e.g. 'C3') */
        DEFAULT,
        /** Row of the address is fixed (e.g. 'C$3') */
        FIXED_ROW,
        /** Column of the address is fixed (e.g. '$C3') */
        FIXED_COLUMN,
        /** Row and column of the address is fixed (e.g. '$C$3') */
        FIXED_ROW_AND_COLUMN,
    }

    /**
     * Enum to define the scope of a passed address string (used in static context)
     */
    public enum AddressScope {
        /** The address represents a single cell or a range of cells */
        ANY,
        /** The address represents a single cell */
        SINGLE_ADDRESS,
        /** The address represents a range of cells */
        RANGE,
        /** The address expression is invalid */
        INVALID
    }

    // private fields
    private Style cellStyle;
    private int columnNumber;
    private CellType dataType;
    private FormulaData formula;
    private int rowNumber;
    private AddressType cellAddressType;
    private Object value;
    private FeatureSet worksheetFeatures;

// getters & setters

    /**
     * Gets the combined cell Address as string in the format A1 - XFD1048576. The address may contain a
     * {@link Cell.AddressType} modifier (e.g. C$50)
     *
     * @return Cell address as string
     */
    public String getCellAddress() {
        return resolveCellAddress(columnNumber, rowNumber, cellAddressType);
    }

    /**
     * Sets the combined cell Address as string in the format A1 - XFD1048576. The address may contain a
     * {@link Cell.AddressType} modifier (e.g. C$50)
     *
     * @param cellAddress Cell address as string
     */
    public void setCellAddress(String cellAddress) {
        Address address = resolveCellCoordinate(cellAddress);
        this.columnNumber = address.column();
        this.rowNumber = address.row();
        this.cellAddressType = address.type();
    }

    /**
     * Gets the combined cell Address as Address object
     *
     * @return Address instance
     */
    public Address getCellAddress2() {
        return new Address(columnNumber, rowNumber, cellAddressType);
    }

    /**
     * Sets the combined cell Address as Address object
     *
     * @param cellAddress Address instance
     */
    public void setCellAddress2(Address cellAddress) {
        setColumnNumber(cellAddress.column());
        setRowNumber(cellAddress.row());
        cellAddressType = cellAddress.type();
    }

    /**
     * Gets the assigned style of the cell
     *
     * @return Cell style
     */
    public Style getCellStyle() {
        return cellStyle;
    }

    /**
     * Gets the number of the column (zero-based)
     *
     * @return Column number
     */
    public int getColumnNumber() {
        return columnNumber;
    }

    /**
     * Sets the number of the column (zero-based)
     *
     * @param columnNumber Column number
     * @throws RangeException Thrown if the column number is out of range
     */
    public void setColumnNumber(int columnNumber) {
        validateColumnNumber(columnNumber);
        this.columnNumber = columnNumber;
    }

    /**
     * Gets the type of the cell
     *
     * @return Type of the cell
     */
    public CellType getDataType() {
        return dataType;
    }

    /**
     * Sets the type of the cell
     *
     * <p>Remarks: Changing the type of an existing cell can create or discard formula metadata and update aggregated
     * workbook features. Prefer assigning {@link Cell#getValue()} when automatic type resolution is intended. Repeated
     * manual transitions to or from {@link CellType#FORMULA} may allocate formula metadata and cause feature-counter
     * propagation.
     * </p>
     */
    public void setDataType(CellType dataType) {
        if (this.dataType == dataType) {
            return;
        }
        if (dataType == CellType.FORMULA) {
            this.setDataType(dataType);
            if (formula == null) {
                this.formula = new FormulaData(getValueAsFormulaExpression());
            } else {
                attachFormulaFeatures();
                synchronizeValueFromFormula();
            }
            return;
        }
        if (this.dataType == CellType.FORMULA) {
            clearFormula();
        }
        this.dataType = dataType;
    }

    /**
     * Gets the number of the row (zero-based)
     *
     * @return Row number
     */
    public int getRowNumber() {
        return rowNumber;
    }

    /**
     * Sets the number of the row (zero-based)
     *
     * @param rowNumber Row number
     * @throws RangeException Thrown if the row number is out of range
     */
    public void setRowNumber(int rowNumber) {
        validateRowNumber(rowNumber);
        this.rowNumber = rowNumber;
    }

    /**
     * Gets the optional address type that can be part of the cell address.
     *
     * @return Address type
     * <p>Remarks: The type has no influence on the behavior of the cell, though. It is preserved to avoid losing
     * information on the address object of the cell</p>
     */
    public AddressType getCellAddressType() {
        return cellAddressType;
    }

    /**
     * Sets the optional address type that can be part of the cell address.
     *
     * @param cellAddressType Address type
     *                        <p>Remarks: The type has no influence on the behavior of the cell, though. It is
     *                        preserved
     *                        to avoid losing information on the address object of the cell</p>
     */
    public void setCellAddressType(AddressType cellAddressType) {
        this.cellAddressType = cellAddressType;
    }

    /**
     * Gets the value of the cell (generic object type)
     *
     */
    public Object getValue() {
        return value;
    }

    /**
     * Sets the value of the cell (generic object type). When setting a value, the {@link Cell#setDataType(CellType)}
     * ()} is automatically resolved
     *
     * <p>Remarks: Assigning a value automatically resolves the cell type and may therefore replace formula metadata.
     * An {@link ch.rabanti.nanoxlsx4j.enums.FormulaError} value resolves to a standalone {@link CellType#ERROR} cell.
     * For formula cells, the assigned value is also synchronized with {@link FormulaData#getExpression()}. Linked
     * formula cells whose {@link FormulaData#getMasterCellAddress()} is set retain their special cached-value
     * behavior.
     * </p>
     */
    public void setValue(Object value) {
        this.value = value;
        resolveCellType();
        if (dataType != CellType.FORMULA || formula == null) {
            return;
        }
        if (formula.getMasterCellAddress() == null) {
            String expression = getValueAsFormulaExpression();
            if (!Objects.equals(formula.getExpression(), expression) && formula.getDefinedNameReference() != null) {
                formula.setDefinedNameReference(null); // Remove additional references
            }
            formula.setExpression(expression);
        }
    }

    /**
     * Gets the Formula object in case of the cell has the DataType {@link CellType#FORMULA}. Default is null, if the
     * cell does not contain a formula
     *
     * <p>Remarks: The plain text of the formula is still set in {@link @link Cell#getValue()}. One exception are
     * linked
     * cells ({@link FormulaData.FormulaType#ARRAY} and {@link FormulaData#getMasterCellAddress()} is set). In this
     * case, the cached value will be in {@link Cell#getValue()} due to compatibility reason.
     * </p>
     */
    public FormulaData getFormula() {
        return formula;
    }

    /**
     * Sets the Formula object internally in case of the cell has the DataType {@link CellType#FORMULA}. Default is
     * null, if the cell does not contain a formula
     *
     * <p>Remarks: The plain text of the formula is still set in {@link @link Cell#getValue()}. One exception are
     * linked
     * cells ({@link FormulaData.FormulaType#ARRAY} and {@link FormulaData#getMasterCellAddress()} is set). In this
     * case, the cached value will be in {@link Cell#getValue()} due to compatibility reason. <br />API note: Do not
     * manually tamper with Formula. There is {@link FeatureSet} inside, responsible for up-stream propagated feature
     * counters.
     * </p>
     */
    void setFormula(FormulaData formula) {
        if (this.formula == formula) {
            return;
        }
        detachFormulaFeatures();
        this.formula = formula;
        synchronizeValueFromFormula();
        attachFormulaFeatures();
    }

    // constructors

    /**
     * Default constructor. Cells created with this constructor do not have a link to a worksheet initially
     */
    public Cell() {
        this.setDataType(CellType.DEFAULT);
    }

    /**
     * Constructor with value and cell type. Cells created with this constructor do not have a link to a worksheet
     * initially
     *
     * <p>Remarks: If the {@link Cell#getDataType()} is defined as {@link CellType#EMPTY} any passed value will be set
     * to null</p>
     *
     * @param value Value of the cell
     * @param type  Type of the cell
     */
    public Cell(Object value, CellType type) {
        if (type == CellType.EMPTY) {
            this.value = null;
        } else {
            this.value = value;
        }
        setDataType(type);
        if (type == CellType.DEFAULT) {
            resolveCellType();
        }
    }

    /**
     * Constructor with value, cell type and address as string. The worksheet reference is set to null and must be
     * assigned later
     *
     * <p>Remarks: If the {@link Cell#getDataType()} is defined as {@link CellType#EMPTY} any passed value will be set
     * to null</p>
     *
     * @param value   Value of the cell
     * @param type    Type of the cell
     * @param address Address of the cell
     */
    public Cell(Object value, CellType type, String address) {
        if (type == CellType.EMPTY) {
            this.value = null;
        } else {
            this.value = value;
        }
        setDataType(type);
        setCellAddress(address);
        if (type == CellType.DEFAULT) {
            resolveCellType();
        }
    }

    /**
     * Constructor with value, cell type and address as struct. The worksheet reference is set to null and must be
     * assigned later
     *
     * <p>Remarks: If the {@link Cell#getDataType()} is defined as {@link CellType#EMPTY} any passed value will be set
     * to null</p>
     *
     * @param value   Value of the cell
     * @param type    Type of the cell
     * @param address Address struct of the cell
     */
    public Cell(Object value, CellType type, Address address) {
        if (type == CellType.EMPTY) {
            this.value = null;
        } else {
            this.value = value;
        }
        setDataType(type);
        columnNumber = address.column();
        rowNumber = address.row();
        this.cellAddressType = address.type();
        if (type == CellType.DEFAULT) {
            resolveCellType();
        }
    }

    /// <summary>
    /// Constructor with value, cell type, row number and column number
    /// </summary>
    /// <param name="value">Value of the cell</param>
    /// <param name="type">Type of the cell</param>
    /// <param name="column">Column number of the cell (zero-based)</param>
    /// <param name="row">Row number of the cell (zero-based)</param>
    public Cell(Object value, CellType type, int column, int row) {
        this(value, type);
        setColumnNumber(column);
        setRowNumber(row);
        this.cellAddressType = AddressType.DEFAULT;
        if (type == CellType.DEFAULT) {
            resolveCellType();
        }
    }

// methods

    /**
     * Removes the assigned style from the cell
     */
    public void removeStyle() {
        this.cellStyle = null;
    }

    /**
     * Sets this cell as a reference to a {@link DefinedName} (workbook- or worksheet-scoped). The cell's
     * {@link Cell#getDataType()} becomes {@link CellType#FORMULA} and its {@link Cell#getValue()} is set to
     * {@link DefinedName#getName()}.
     *
     * @param definedName Defined name to associate with this cell. Must not be null.
     * @return Returns the range object of transposed linked cells if the type is {@link DefinedName.NameType#RANGE}.
     * The value is null otherwise.
     * @throws WorksheetException Thrown if {@code definedName} is null.
     */
    Range setReference(DefinedName definedName) {
        return setReference(definedName, null);
    }

    /**
     * Sets this cell as a reference to a {@link DefinedName} (workbook- or worksheet-scoped). The cell's
     * {@link Cell#getDataType()} becomes {@link CellType#FORMULA} and its {@link Cell#getValue()} is set to
     * {@link DefinedName#getName()}.
     *
     * @param definedName Defined name to associate with this cell. Must not be null.
     * @param cachedValue Optional cached value that will be shown as long as the cell is not refreshed. The value will
     *                    be ignored if the defined name type is {@link DefinedName.NameType#CONSTANT}
     * @return Returns the range object of transposed linked cells if the type is {@link DefinedName.NameType#RANGE}.
     * The value is null otherwise.
     * @throws WorksheetException Thrown if {@code definedName} is null.
     */
    Range setReference(DefinedName definedName, Object cachedValue) {
        if (definedName == null) {
            throw new WorksheetException("The defined name to set as cell reference must not be null.");
        }
        if (this.formula == null) {
            this.formula = new FormulaData();
        }
        FormulaData formula = this.formula;
        Range referenceRange = null;
        formula.setDefinedNameReference(definedName);
        formula.setExpression(definedName.getName());
        if (definedName.getType() == DefinedName.NameType.RANGE) {
            formula.setType(FormulaData.FormulaType.ARRAY);
            referenceRange = transposeDefinedNameArrayRange(definedName.getTextValue());
        }
        if (definedName.getType() == DefinedName.NameType.CONSTANT) {
            formula.setCachedValue(definedName.getTextValue());
            formula.setCachedValueType(FormulaData.resolveCachedValueType(definedName.getValue()));
        } else {
            if (cachedValue == null || (cachedValue instanceof String && ((String) cachedValue).isEmpty())) {
                formula.setCachedValueType(CellType.NUMBER);
            } else {
                formula.setCachedValueType(FormulaData.resolveCachedValueType(cachedValue));
            }
            formula.setCachedValue(ParserUtils.toCachedValueString(cachedValue)); // Force value as plain OOXML string
        }
        this.setDataType(CellType.FORMULA); // Force type
        this.formula = formula;
        this.value = definedName.getName();
        return referenceRange;
    }

    /**
     * Method resets the Cell type and tries to find the actual type. This is used if a Cell was created with the
     * {@link Cell.CellType#DEFAULT} or automatically if a value was set by {@link Cell#setValue(Object)}}.
     * {@link Cell.CellType#FORMULA} will skip this method and {@link Cell.CellType#EMPTY} will discard the value of the
     * cell
     */
    public void resolveCellType() {
        if (this.value == null) {
            this.setDataType(CellType.EMPTY);
            this.value = null;
            return;
        }
        if (this.dataType == CellType.FORMULA) {
            return;
        } // Do not overwrite type
        Object t = this.value;
        if (t instanceof Boolean) {
            this.setDataType(CellType.BOOL);
        } else if (t instanceof Byte) // C# type sbyte not existing
        {
            setDataType(CellType.NUMBER);
        } else if (t instanceof BigDecimal) // c# decimal
        {
            setDataType(CellType.NUMBER);
        } else if (t instanceof Double) {
            setDataType(CellType.NUMBER);
        } else if (t instanceof Float) {
            setDataType(CellType.NUMBER);
        } else if (t instanceof Integer) // C# type uint not existing
        {
            setDataType(CellType.NUMBER);
        } else if (t instanceof Long) // C# type ulong not existing
        {
            setDataType(CellType.NUMBER);
        } else if (t instanceof Short) // C# type ushort not existing
        {
            setDataType(CellType.NUMBER);
        } else if (t instanceof Date) {
            setDataType(CellType.DATE);
            setStyle(BasicStyles.getDateFormat());
        } else if (t instanceof Duration) {
            setDataType(CellType.TIME);
            setStyle(BasicStyles.getTimeFormat());
        } else if (t instanceof Errors) {
            setDataType(CellType.ERROR);
        } else {
            setDataType(CellType.STRING);
        } // Default (char, string, object)
    }

    /**
     * Sets the lock state of the cell
     *
     * <p>Remarks: The listed exception should never happen because the mentioned style is internally generated</p>
     *
     * @param isLocked If true, the cell will be locked if the worksheet is protected
     * @param isHidden If true, the value of the cell will be invisible if the worksheet is protected
     * @throws StyleException Throws a StyleException if the style used to lock cells cannot be referenced
     */
    public void setCellLockedState(boolean isLocked, boolean isHidden) {
        Style lockStyle;
        if (cellStyle == null) {
            lockStyle = new Style();
        } else {
            lockStyle = cellStyle.copyStyle();
        }
        lockStyle.getCurrentCellXf().setLocked(isLocked);
        lockStyle.getCurrentCellXf().setHidden(isHidden);
        setStyle(lockStyle);
    }

    /**
     * Sets the style of the cell
     *
     * @param style Style to assign
     * @return If the passed style already exists in the repository, the existing one will be returned, otherwise the
     * passed one
     */
    public Style setStyle(Style style) {
        return setStyleInternal(style, false);
    }

    /**
     * Sets the style of the cell internally
     *
     * @param style     Style to assign
     * @param unmanaged Internally used: If true, the style repository is not invoked and only the style object of the
     *                  cell is updated. Do not use!
     * @return If the passed style already exists in the repository, the existing one will be returned, otherwise the
     * passed one
     */
    @InternalApi
    public Style setStyleInternal(Style style, boolean unmanaged) {
        if (style == null) {
            throw new StyleException("No style to assign was defined");
        }
        if (unmanaged) {
            this.cellStyle = style;
        } else {
            this.cellStyle = StyleRepository.getInstance().addStyle(style);
        }
        return this.cellStyle;
    }

    /**
     * Copies this cell into a new one. The style is considered if not null.
     *
     * @return Copy of this cell
     */
    Cell copy() {
        Cell copy = new Cell();
        copy.value = this.value;
        copy.setCellAddress(this.getCellAddress());
        copy.setCellAddressType(this.cellAddressType);
        if (this.formula != null) {
            copy.formula = this.formula.copy();
        }
        if (this.cellStyle != null) {
            copy.setStyleInternal(this.cellStyle, true);
        }
        return copy;
    }

    /**
     * Implemented CompareTo method
     *
     * <p>Remarks: Note that this method only compares the row and column numbers,
     * since the values or styles may be completely different types, and therefore hard to compare at all.<br /> The
     * {@link Cell#equals(Object)} method considers values and style, though.
     * </p>
     *
     * @param other Object to compare
     * @return 0 if values are the same, -1 if this object is smaller, 1 if it is bigger
     */
    @Override
    public int compareTo(Cell other) {
        if (other == null) {
            return -1;
        }
        if (rowNumber == other.rowNumber) {
            return Integer.compare(columnNumber, other.columnNumber);
        }

        return Integer.compare(rowNumber, other.rowNumber);
    }

    /**
     * Compares two objects whether they are addresses and equal
     *
     * @param o the reference object with which to compare.
     * @return True if not null, of the same type and equal
     */
    @Override
    public final boolean equals(Object o) {
        if (!(o instanceof Cell cell)) {
            return false;
        }

        return columnNumber == cell.columnNumber && rowNumber == cell.rowNumber &&
                Objects.equals(cellStyle, cell.cellStyle) && dataType == cell.dataType &&
                Objects.equals(formula, cell.formula) && cellAddressType == cell.cellAddressType &&
                Objects.equals(value, cell.value);
    }

    /**
     * Gets the hash code of the cell
     *
     * @return Hash code
     */
    @Override
    public int hashCode() {
        int result = Objects.hashCode(cellStyle);
        result = 31 * result + columnNumber;
        result = 31 * result + Objects.hashCode(dataType);
        result = 31 * result + Objects.hashCode(formula);
        result = 31 * result + rowNumber;
        result = 31 * result + Objects.hashCode(cellAddressType);
        result = 31 * result + Objects.hashCode(value);
        return result;
    }

// static methods

    /**
     * Converts a List of supported objects into a list of cells
     *
     * @param <T>  Generic data type
     * @param list List of generic objects
     * @return List of cells
     */
    public static <T> List<Cell> convertArray(List<T> list) {
        List<Cell> output = new ArrayList<>();
        if (list == null) {
            return output;
        }
        Cell c;
        Object o;
        //Type t;
        for (T item : list) {
            if (item ==
                    null) // DO NOT LISTEN to code suggestions! This is wrong for bool: if (object.Equals(item,
            // default(T)))
            {
                c = new Cell(null, CellType.EMPTY);
                output.add(c);
                continue;
            }
            o = item; // intermediate object is necessary to cast the types below
            // t = item.GetType();
            if (item instanceof Cell) {
                c = (Cell) item;
            } else if (item instanceof Boolean) {
                c = new Cell((boolean) o, CellType.BOOL);
            } else if (item instanceof Byte) {
                c = new Cell((byte) o, CellType.NUMBER);
            } // no C# sbyte available

            else if (item instanceof BigDecimal) {
                c = new Cell((BigDecimal) o, CellType.NUMBER);
            } else if (item instanceof Double) {
                c = new Cell((double) o, CellType.NUMBER);
            } else if (item instanceof Float) {
                c = new Cell((float) o, CellType.NUMBER);
            } else if (item instanceof Integer) // no C# uint available
            {
                c = new Cell((int) o, CellType.NUMBER);
            } else if (item instanceof Long) // no C# ulong available
            {
                c = new Cell((long) o, CellType.NUMBER);
            } else if (item instanceof Short) {
                c = new Cell((short) o, CellType.NUMBER);
            } // no C# ushort available}
            else if (item instanceof Date) {
                c = new Cell((Date) o, CellType.DATE);
                c.setStyle(BasicStyles.getDateFormat());
            } else if (item instanceof Duration) {
                c = new Cell((Duration) o, CellType.TIME);
                c.setStyle(BasicStyles.getTimeFormat());
            } else if (item instanceof String) {
                c = new Cell((String) o, CellType.STRING);
            } else // Default = unspecified object
            {
                c = new Cell(o.toString(), CellType.DEFAULT);
            }
            output.add(c);
        }
        return output;
    }

    /**
     * Gets a list of cell addresses from a cell range (format A1:B3 or AAD556:AAD1000)
     *
     * @param range Range to process
     * @return List of cell addresses
     * @throws FormatException Throws a FormatException if a part of the passed range is malformed
     * @throws RangeException  Throws a RangeException if the range is out of range (A-XFD and 1 to 1048576)
     */
    public static List<Address> getCellRange(String range) {
        Range range2 = resolveCellRange(range);
        return getCellRange(range2.startAddress(), range2.endAddress());
    }

    /**
     * Get a list of cell addresses from a cell range
     *
     * @param startAddress Start address as string in the format A1 - XFD1048576
     * @param endAddress   End address as string in the format A1 - XFD1048576
     * @return List of cell addresses
     * @throws FormatException Throws a FormatException if a part of the passed range is malformed
     * @throws RangeException  Throws a RangeException if the range is out of range (A-XFD and 1 to 1048576)
     */
    public static List<Address> getCellRange(String startAddress, String endAddress) {
        Address start = resolveCellCoordinate(startAddress);
        Address end = resolveCellCoordinate(endAddress);
        return getCellRange(start, end);
    }

    /**
     * Get a list of cell addresses from a cell range
     *
     * @param startColumn Start column (zero based)
     * @param startRow    Start row (zero based)
     * @param endColumn   End column (zero based)
     * @param endRow      End row (zero based)
     * @return List of cell addresses
     * @throws RangeException Throws a RangeException if the value of one passed address parts is out of range (A-XFD
     *                        and 1 to 1048576)
     */
    public static List<Address> getCellRange(int startColumn, int startRow, int endColumn, int endRow) {
        Address start = new Address(startColumn, startRow);
        Address end = new Address(endColumn, endRow);
        return getCellRange(start, end);
    }

    /**
     * Get a list of cell addresses from a cell range
     *
     * @param startAddress Start address
     * @param endAddress   End address
     * @return List of cell addresses
     * @throws FormatException Throws a FormatException if a part of the passed addresses is malformed
     * @throws RangeException  Throws a RangeException if the value of one passed address is out of range (A-XFD and 1
     *                         to 1048576)
     */
    public static List<Address> getCellRange(Address startAddress, Address endAddress) {
        int startColumn;
        int endColumn;
        int startRow;
        int endRow;
        if (startAddress.column() < endAddress.column()) {
            startColumn = startAddress.column();
            endColumn = endAddress.column();
        } else {
            startColumn = endAddress.column();
            endColumn = startAddress.column();
        }
        if (startAddress.row() < endAddress.row()) {
            startRow = startAddress.row();
            endRow = endAddress.row();
        } else {
            startRow = endAddress.row();
            endRow = startAddress.row();
        }
        List<Address> output = new ArrayList<>();
        for (int column = startColumn; column <= endColumn; column++) {
            for (int row = startRow; row <= endRow; row++) {
                output.add(new Address(column, row));
            }
        }
        return output;
    }

    /**
     * Gets the address of a cell by the column and row number (zero based; default referencing)
     *
     * @param column Column number of the cell (zero-based)
     * @param row    Row number of the cell (zero-based)
     * @return Cell Address as string in the format A1 - XFD1048576. Depending on the type, Addresses like '$A55', 'B$2'
     * or '$A$5' are possible outputs
     * @throws RangeException Thrown if the start or end address was out of range
     */
    public static String resolveCellAddress(int column, int row) throws RangeException {
        return resolveCellAddress(column, row, AddressType.DEFAULT);
    }

    /**
     * Gets the address of a cell by the column and row number (zero based)
     *
     * @param column Column number of the cell (zero-based)
     * @param row    Row number of the cell (zero-based)
     * @param type   Referencing type of the address
     * @return Cell Address as string in the format A1 - XFD1048576. Depending on the type, Addresses like '$A55', 'B$2'
     * or '$A$5' are possible outputs
     * @throws RangeException Thrown if the start or end address was out of range
     */
    public static String resolveCellAddress(int column, int row, AddressType type) throws RangeException {
        validateColumnNumber(column);
        validateRowNumber(row);
        return switch (type) {
            case AddressType.FIXED_ROW_AND_COLUMN -> "$" + resolveColumnAddress(column) + "$" + (row + 1);
            case AddressType.FIXED_COLUMN -> "$" + resolveColumnAddress(column) + (row + 1);
            case AddressType.FIXED_ROW -> resolveColumnAddress(column) + "$" + (row + 1);
            default -> resolveColumnAddress(column) + (row + 1);
        };
    }

    //  /**
    //   * Gets the column and row number (zero based) of a cell by the address
    //   *
    //   * @param address Address as string in the format A1 - XFD1048576
    //   * @return Struct with row and column
    //   * @throws FormatException Throws a FormatException if the passed address is malformed
    //   * @throws RangeException Throws a RangeException if the value of the passed address is out of range (A-XFD
    //   and 1 to 1048576)
    //   */
    //  public static Address resolveCellCoordinate(String address)
    //  {
    //      int row;
    //      int column;
    //      AddressType type;
    //      Address addressObject = resolveCellCoordinate(address);
    //      return new Address(addressObject.column(), addressObject.row(), addressObject.type());
    //  }

    /**
     * Gets the column and row number (zero based) of a cell by the address
     *
     * @param address Address as string in the format A1 - XFD1048576
     * @return Struct with row and column
     * @throws FormatException Throws a FormatException if the passed address is malformed
     * @throws RangeException  Throws a RangeException if the value of the passed address is out of range (A-XFD and 1
     *                         to 1048576)
     */
    public static Address resolveCellCoordinate(String address) {
        if (address == null || address.isEmpty()) {
            throw new FormatException("The cell address is null or empty and could not be resolved");
        }

        int i = 0;
        int length = address.length();
        boolean fixedColumn = false;
        boolean fixedRow = false;

        // Optional $ for column
        if (address.charAt(i) == '$') {
            fixedColumn = true;
            i++;
        }

        // Column letters
        int columnStart = i;
        while (i < length && isAsciiLetter(address.charAt(i))) {
            i++;
        }
        if (i == columnStart) {
            throw new FormatException("The format of the cell address (" + address + ") is malformed");
        }

        String columnPart = address.substring(columnStart, i);

        // Optional $ for row
        if (i < length && address.charAt(i) == '$') {
            fixedRow = true;
            i++;
        }

        // Row digits
        int rowStart = i;
        while (i < length && address.charAt(i) >= '0' && address.charAt(i) <= '9') {
            i++;
        }

        if (i == rowStart || i != length) {
            throw new FormatException("The format of the cell address (" + address + ") is malformed");
        }

        int row = Integer.parseInt(address.substring(rowStart)) - 1;
        int column = resolveColumn(columnPart);
        validateRowNumber(row);

        AddressType type;
        if (fixedColumn && fixedRow) {
            type = AddressType.FIXED_ROW_AND_COLUMN;
        } else if (fixedColumn) {
            type = AddressType.FIXED_COLUMN;
        } else if (fixedRow) {
            type = AddressType.FIXED_ROW;
        } else {
            type = AddressType.DEFAULT;
        }
        return new Address(column, row, type);
    }

    /**
     * Resolves a cell range from the format like A1:B3 or AAD556:AAD1000
     *
     * @param range Range to process
     * @return Range object
     * @throws FormatException Throws a FormatException if the start or end address was malformed
     * @throws RangeException  Throws a RangeException if the range is out of range (A-XFD and 1 to 1048576)
     */
    public static Range resolveCellRange(String range) {
        if (range == null || range.isEmpty()) {
            throw new FormatException("The cell range is null or empty and could not be resolved");
        }
        if (!range.contains(":")) {
            Address address = resolveCellCoordinate(range);
            return new Range(address, address);
        }
        String[] split = range.split(":", -1);
        if (split.length != 2) {
            throw new FormatException("The cell range (" + range + ") is malformed and could not be resolved");
        }
        return new Range(resolveCellCoordinate(split[0]), resolveCellCoordinate(split[1]));
    }

    /**
     * Gets the zero-based column number from a column address.
     *
     * @param columnAddress column address in the format A - XFD
     * @return zero-based column number
     * @throws RangeException if the column is out of range
     */
    public static int resolveColumn(String columnAddress) {
        if (columnAddress == null || columnAddress.isEmpty()) {
            throw new RangeException("The passed address was null or empty");
        }
        String normalizedAddress = columnAddress.toUpperCase(java.util.Locale.ROOT);
        int result = 0;
        int multiplier = 1;
        for (int i = normalizedAddress.length() - 1; i >= 0; i--) {
            int character = normalizedAddress.charAt(i) - 64;
            result += character * multiplier;
            multiplier *= 26;
        }
        validateColumnNumber(result - 1);
        return result - 1;
    }

    /**
     * Gets the column address (A - XFD)
     *
     * @param columnNumber Column number (zero-based)
     * @return Column address (A - XFD)
     * @throws RangeException Thrown if the passed column number was out of range
     */
    public static String resolveColumnAddress(int columnNumber) throws RangeException {
        validateColumnNumber(columnNumber);
        // A - XFD
        StringBuilder sb = new StringBuilder();
        columnNumber++;
        while (columnNumber > 0) {
            columnNumber--;
            sb.insert(0, (char) ('A' + (columnNumber % 26)));
            columnNumber /= 26;
        }
        return sb.toString();
    }

    /**
     * Gets the scope of the passed address (string expression). Scope means either single cell address or range
     *
     * @param addressExpression Address expression
     * @return Scope of the address expression
     */
    public static AddressScope getAddressScope(String addressExpression) {
        try {
            resolveCellCoordinate(addressExpression);
            return AddressScope.SINGLE_ADDRESS;
        } catch (Exception e) // any
        {
            try {
                resolveCellRange(addressExpression);
                return AddressScope.RANGE;
            } catch (Exception e2) // any
            {
                return AddressScope.INVALID;
            }
        }
    }

    /**
     * Validates the passed (zero-based) column number. An exception will be thrown if the column is invalid
     *
     * @param column Number to check
     * @throws RangeException Thrown if the passed column number is out of range
     */
    public static void validateColumnNumber(int column) throws RangeException {
        if (column > Worksheet.MAX_COLUM_NUMBER || column < Worksheet.MIN_COLUM_NUMBER) {
            throw new RangeException("The column number (" + column + ") is out of range. Range is from " +
                    Worksheet.MIN_COLUM_NUMBER + " to " + Worksheet.MAX_COLUM_NUMBER + " (" +
                    (Worksheet.MAX_COLUM_NUMBER + 1) + " columns).");
        }
    }

    /**
     * Validates the passed (zero-based) row number. An exception will be thrown if the row is invalid
     *
     * @param row Number to check
     * @throws RangeException Thrown if the passed row number is out of range
     */
    public static void validateRowNumber(int row) throws RangeException {
        if (row > Worksheet.MAX_ROW_NUMBER || row < Worksheet.MIN_ROW_NUMBER) {
            throw new RangeException("The row number (" + row + ") is out of range. Range is from " +
                    Worksheet.MIN_ROW_NUMBER + " to " + Worksheet.MAX_ROW_NUMBER + " (" +
                    (Worksheet.MAX_ROW_NUMBER + 1) + " rows).");
        }
    }

    /**
     * Binds this cell to the aggregate feature set of its worksheet.
     *
     * @param features Worksheet feature set.
     */
    void bindFeatures(FeatureSet features) {
        if (worksheetFeatures == features) {
            return;
        }
        unbindFeatures();
        worksheetFeatures = features;
        attachFormulaFeatures();
    }

    /**
     * Removes this cell's formula contribution from its worksheet feature set.
     */
    void unbindFeatures() {
        detachFormulaFeatures();
        worksheetFeatures = null;
    }

    /**
     * Propagates Cell.Formula.Expression back to Cell.Value
     */
    private void synchronizeValueFromFormula() {
        if (formula.getMasterCellAddress() == null) {
            this.value = formula.getExpression();  // Sync back from formula to cell
        }
    }

    /**
     * Adds a feature set to the formula. Mainly used to add a formula in an existing cell
     */
    private void attachFormulaFeatures() {
        if (worksheetFeatures != null && dataType == CellType.FORMULA && formula != null) {
            formula.getFeatures().add(worksheetFeatures);
        }
    }

    /**
     * Clears a formula and all its metadata from an existing cell
     */
    private void clearFormula() {
        detachFormulaFeatures();
        formula = null;
    }

    /**
     * Removes the feature set of the formula. Mainly used to clear a formula in a existing cell
     */
    private void detachFormulaFeatures() {
        if (worksheetFeatures != null && dataType == CellType.FORMULA && formula != null) {
            formula.getFeatures().Remove(worksheetFeatures);
        }
    }

    /**
     * Gets the value of the cell as formula expression (may lead to syntactical invalid Excel formulas)
     *
     * @return Value as string or null, of no value was set
     */
    private String getValueAsFormulaExpression() {
        if (formula == null) {
            return null;
        }
        return value.toString();
    }

    /**
     * Transposes the range expression of a defined name to the target range of linked cells
     *
     * @param referenceExpression Range expression as string (to be validated first)
     * @return Transposed range of affected liked cells
     */
    private Range transposeDefinedNameArrayRange(String referenceExpression) {
        Range resolvedRange = new Range(referenceExpression);
        int rowCount = resolvedRange.endAddress().row() - resolvedRange.startAddress().row();
        int columnCount = resolvedRange.endAddress().column() - resolvedRange.startAddress().column();
        return new Range(this.columnNumber, this.rowNumber, this.columnNumber + columnCount, this.rowNumber + rowCount);
    }

    private static boolean isAsciiLetter(char character) {
        return character >= 'A' && character <= 'Z' || character >= 'a' && character <= 'z';
    }


}
