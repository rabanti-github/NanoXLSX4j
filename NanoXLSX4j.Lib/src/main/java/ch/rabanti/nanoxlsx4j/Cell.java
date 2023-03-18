/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import java.math.BigDecimal;
import java.time.Duration;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;
import ch.rabanti.nanoxlsx4j.styles.StyleRepository;

/**
 * Class representing a cell of a worksheet
 *
 * @author Raphael Stoeckli
 */
public class Cell implements Comparable<Cell> {

	// ### C O N S T A N T S ###

	private static final int ASCII_OFFSET = 64;

	// ### E N U M S ###

	/**
	 * Enum defines the basic data types of a cell
	 */
	public enum CellType {
		/**
		 * Type for single characters and strings
		 */
		STRING,
		/**
		 * Type for all numeric types (long, short, integer, float, double and
		 * BigDecimal)
		 */
		NUMBER,
		/**
		 * Type for dates and times (Note: Dates before 1900-01-01 are not allowed)
		 */
		DATE,
		/**
		 * Type for times (Note: Internally handled as OAdate, represented by
		 * {@link LocalTime}
		 */
		TIME,
		/**
		 * Type for boolean
		 */
		BOOL,
		/**
		 * Type for Formulas (The cell will be handled differently)
		 */
		FORMULA,
		/**
		 * Type for empty cells. This type is only used for merged cells (all cells
		 * except the first of the cell range)
		 */
		EMPTY,
		/**
		 * Default Type, not specified
		 */
		DEFAULT
	}

	/**
	 * Enum for the referencing style of the address
	 */
	public enum AddressType {
		/**
		 * Default behavior (e.g. 'C3')
		 */
		Default,
		/**
		 * Row of the address is fixed (e.g. 'C$3')
		 */
		FixedRow,
		/**
		 * Column of the address is fixed (e.g. '$C3')
		 */
		FixedColumn,
		/**
		 * Row and column of the address is fixed (e.g. '$C$3')
		 */
		FixedRowAndColumn
	}

	/**
	 * Enum to define the scope of a passed address string (used in static context)
	 */
	public enum AddressScope {
		/**
		 * The address represents a single cell
		 */
		SingleAddress,
		/**
		 * The address represents a range of cells
		 */
		Range,
		/**
		 * The address expression is invalid
		 */
		Invalid
	}

	// ### P R I V A T E F I E L D S ###

	private Style cellStyle;
	private int columnNumber;
	private int rowNumber;
	private CellType dataType;
	private Object value;
	private AddressType cellAddressType = AddressType.Default;

	// ### G E T T E R S & S E T T E R S ###

	/**
	 * Gets the combined cell Address as string in the format A1 - XFD1048576
	 *
	 * @return Cell address
	 */
	public String getCellAddress() {
		return resolveCellAddress(this.columnNumber, this.rowNumber, this.cellAddressType);
	}

	/**
	 * Sets the combined cell Address as string in the format A1 - XFD1048576
	 *
	 * @param address
	 *            Cell address
	 * @throws RangeException
	 *             Thrown in case of a illegal address
	 */
	public void setCellAddress(String address) {
		Address temp = Cell.resolveCellCoordinate(address);
		this.columnNumber = temp.Column;
		this.rowNumber = temp.Row;
		this.cellAddressType = temp.Type;
	}

	/**
	 * Gets the combined cell address as class
	 *
	 * @return Cell address
	 */
	public Address getCellAddress2() {
		return new Address(this.columnNumber, this.rowNumber, this.cellAddressType);
	}

	/**
	 * Sets the combined cell address as class
	 *
	 * @param address
	 *            Cell address
	 */
	public void setCellAddress2(Address address) {
		this.setColumnNumber(address.Column);
		this.setRowNumber(address.Row);
		this.setCellAddressType(address.Type);
	}

	/**
	 * Gets the assigned style of the cell
	 *
	 * @return Assigned style
	 */
	public Style getCellStyle() {
		return cellStyle;
	}

	/**
	 * Gets the number of the column (zero-based)
	 *
	 * @return Column number (zero-based)
	 */
	public int getColumnNumber() {
		return columnNumber;
	}

	/**
	 * Sets the number of the column (zero-based)
	 *
	 * @param columnNumber
	 *            Column number (zero-based)
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
	 * @param dataType
	 *            Type of the cell
	 */
	public void setDataType(CellType dataType) {
		this.dataType = dataType;
	}

	/**
	 * Gets the number of the row (zero-based)
	 *
	 * @return Row number (zero-based)
	 */
	public int getRowNumber() {
		return rowNumber;
	}

	/**
	 * Sets the number of the row (zero-based)
	 *
	 * @param rowNumber
	 *            Row number (zero-based)
	 */
	public void setRowNumber(int rowNumber) {
		validateRowNumber(rowNumber);
		this.rowNumber = rowNumber;
	}

	/**
	 * Gets the optional address type that can be part of the cell address.
	 *
	 * @return Address type
	 * @apiNote The type has no influence on the behavior of the cell, though. It is
	 *          preserved to avoid losing information on the address object of the
	 *          cell
	 */
	public AddressType getCellAddressType() {
		return cellAddressType;
	}

	/**
	 * Sets the optional address type that can be part of the cell address.
	 *
	 * @param cellAddressType
	 *            Address type
	 * @apiNote The type has no influence on the behavior of the cell, though. It is
	 *          preserved to avoid losing information on the address object of the
	 *          cell
	 */
	public void setCellAddressType(AddressType cellAddressType) {
		this.cellAddressType = cellAddressType;
	}

	/**
	 * Gets the value of the cell (generic object type)
	 *
	 * @return Value of the cell
	 */
	public Object getValue() {
		return value;
	}

	/**
	 * Sets the value of the cell (generic object type). When setting a value, the
	 * {@link Cell#dataType } is automatically resolved
	 *
	 * @param value
	 *            Value of the cell
	 */
	public void setValue(Object value) {
		this.value = value;
		resolveCellType();
	}

	// ### C O N S T R U C T O R S ###

	/**
	 * Default constructor. Cells created with this constructor do not have a link
	 * to a worksheet initially
	 */
	public Cell() {
		this.dataType = CellType.DEFAULT;
	}

	/**
	 * Constructor with value and cell type. Cells created with this constructor do
	 * not have a link to a worksheet initially
	 *
	 * @param value
	 *            Value of the cell
	 * @param type
	 *            Type of the cell
	 * @apiNote If the {@link Cell#dataType} is defined as {@link CellType#EMPTY}
	 *          any passed value will be set to null
	 */
	public Cell(Object value, CellType type) {
		this.dataType = type;
		if (type == CellType.EMPTY) {
			this.value = null;
		}
		else {
			this.value = value;
		}
		if (type == CellType.DEFAULT) {
			resolveCellType();
		}
	}

	/**
	 * Constructor with value, cell type and address. The worksheet reference is set
	 * to null and must be assigned later
	 *
	 * @param value
	 *            Value of the cell
	 * @param type
	 *            Type of the cell
	 * @param address
	 *            Address of the cell
	 * @apiNote If the {@link Cell#dataType} is defined as {@link CellType#EMPTY}
	 *          any passed value will be set to null
	 */
	public Cell(Object value, CellType type, String address) {
		this.dataType = type;
		if (type == CellType.EMPTY) {
			this.value = null;
		}
		else {
			this.value = value;
		}
		this.setCellAddress(address);
		if (type == CellType.DEFAULT) {
			resolveCellType();
		}
	}

	/**
	 * Constructor with value, cell type and address class. The worksheet reference
	 * is set to null and must be assigned later
	 *
	 * @param value
	 *            Value of the cell
	 * @param type
	 *            Type of the cell
	 * @param address
	 *            Address class of the cell
	 * @apiNote If the {@link Cell#dataType} is defined as {@link CellType#EMPTY}
	 *          any passed value will be set to null
	 */
	public Cell(Object value, CellType type, Address address) {
		this.dataType = type;
		if (type == CellType.EMPTY) {
			this.value = null;
		}
		else {
			this.value = value;
		}
		this.setCellAddress2(address);
		if (type == CellType.DEFAULT) {
			resolveCellType();
		}
	}

	/**
	 * Constructor with value, cell type, row number and column number
	 *
	 * @param value
	 *            Value of the cell
	 * @param type
	 *            Type of the cell
	 * @param column
	 *            Column number of the cell (zero-based)
	 * @param row
	 *            Row number of the cell (zero-based)
	 * @apiNote If the {@link Cell#dataType} is defined as {@link CellType#EMPTY}
	 *          any passed value will be set to null
	 */
	public Cell(Object value, CellType type, int column, int row) {
		this(value, type);
		this.columnNumber = column;
		this.rowNumber = row;
		this.cellAddressType = AddressType.Default;
		if (type == CellType.DEFAULT) {
			resolveCellType();
		}
	}

	// ### M E T H O D S ###

	/**
	 * Implemented compareTo method
	 *
	 * @param other
	 *            Object to compare
	 * @return 0 if values are the same, -1 if this object is smaller, 1 if it is
	 *         bigger
	 * @implNote Note that this method only compares the row and column numbers,
	 *           since the values or styles may completely different types, and
	 *           therefore hard to compare at all.<br>
	 *           The Equals() method considers values and style, though.
	 */
	@Override
	public int compareTo(Cell other) {
		if (other == null) {
			return -1;
		}
		if (this.rowNumber == other.rowNumber) {
			return Integer.compare(this.columnNumber, other.getColumnNumber());
		}
		return Integer.compare(this.rowNumber, other.getRowNumber());
	}

	@Override
	public boolean equals(Object obj) {
		if (obj == null || !obj.getClass().equals(Cell.class)) {
			return false;
		}
		Cell other = (Cell) obj;
		if (!this.getCellAddress2().equals(other.getCellAddress2())) {
			return false;
		}
		if (this.cellStyle != null && other.getCellStyle() != null && !this.getCellStyle().equals(other.getCellStyle())) {
			return false;
		}
		if (this.getDataType() != other.getDataType()) {
			return false;
		}
        return this.getValue() == null || other.getValue() == null || this.getValue().equals(other.getValue());
    }

	/**
	 * Removes the assigned style from the cell
	 *
	 * @throws StyleException
	 *             Thrown if the workbook to remove was not found in the style sheet
	 *             collection
	 */
	public void removeStyle() {
		this.cellStyle = null;
	}

	/**
	 * Method resets the Cell type and tries to find the actual type. This is used
	 * if a Cell was created with the CellType DEFAULT or automatically if a value
	 * was set by {@link Cell#setValue(Object)}. CellType FORMULA will skip this
	 * method and EMPTY will discard the value of the cell
	 */
	public void resolveCellType() {
		if (this.value == null) {
			this.setDataType(CellType.EMPTY);
			return;
		} // the following section is intended to be as similar as possible to PicoXLSX
			// for C#
		if (this.dataType == CellType.FORMULA) {
        }
		else if (value instanceof Boolean) {
			this.dataType = CellType.BOOL;
		}
		else if (value instanceof Byte) {
			this.dataType = CellType.NUMBER;
		} // sbyte not existing in Java
		else if (value instanceof BigDecimal) {
			this.dataType = CellType.NUMBER;
		} // decimal
		else if (value instanceof Double) {
			this.dataType = CellType.NUMBER;
		}
		else if (value instanceof Float) {
			this.dataType = CellType.NUMBER;
		}
		else if (value instanceof Integer) {
			this.dataType = CellType.NUMBER;
		} // uint not existing in Java
		else if (value instanceof Long) {
			this.dataType = CellType.NUMBER;
		}
		else if (value instanceof Short) {
			this.dataType = CellType.NUMBER;
		} // ushort not existing in Java
		else if (value instanceof Date) {
			this.dataType = CellType.DATE;
			setStyle(BasicStyles.DateFormat());
		}
		else if (value instanceof Duration) {
			this.dataType = CellType.TIME;
			setStyle(BasicStyles.TimeFormat());
		}
		else {
			this.dataType = CellType.STRING;
		} // Default (char, string, object)
	}

	/**
	 * Sets the lock state of the cell
	 *
	 * @param isLocked
	 *            If true, the cell will be locked if the worksheet is protected
	 * @param isHidden
	 *            If true, the value of the cell will be invisible if the worksheet
	 *            is protected
	 * @throws StyleException
	 *             Throws an UndefinedStyleException if the style used to lock cells
	 *             cannot be referenced
	 * @apiNote The listed exception should never happen because the mentioned style
	 *          is internally generated
	 */
	public void setCellLockedState(boolean isLocked, boolean isHidden) {
		Style lockStyle;
		if (this.cellStyle == null) {
			lockStyle = new Style();
		}
		else {
			lockStyle = this.cellStyle.copyStyle();
		}
		lockStyle.getCellXf().setLocked(isLocked);
		lockStyle.getCellXf().setHidden(isHidden);
		this.setStyle(lockStyle);
	}

	/**
	 * Sets the style of the cell
	 *
	 * @param style
	 *            style to assign
	 * @return If the passed style already exists in the repository, the existing
	 *         one will be returned, otherwise the passed one
	 */
	public Style setStyle(Style style) {
		return setStyle(style, false);
	}

	/**
	 * Sets the style of the cell
	 *
	 * @param style
	 *            style to assign
	 * @param unmanaged
	 *            Internally used: If true, the style repository is not invoked and
	 *            only the style object of the cell is updated. Do not use!
	 * @return If the passed style already exists in the repository, the existing
	 *         one will be returned, otherwise the passed one
	 */
	public Style setStyle(Style style, boolean unmanaged) {
		if (style == null) {
			throw new StyleException("No style to assign was defined");
		}
		if (unmanaged) {
			this.cellStyle = style;
		}
		else {
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
		copy.setDataType(this.dataType);
		copy.setColumnNumber(this.columnNumber);
		copy.setRowNumber(this.rowNumber);
		copy.setCellAddressType(this.cellAddressType);
		if (this.cellStyle != null) {
			copy.setStyle(this.cellStyle, true);
		}
		return copy;
	}

	// ### S T A T I C M E T H O D S ###

	/**
	 * Converts a List of supported objects into a list of cells
	 *
	 * @param <T>
	 *            Generic data type
	 * @param list
	 *            List of generic objects
	 * @return List of cells
	 */
	public static <T> List<Cell> convertArray(List<T> list) {
		List<Cell> output = new ArrayList<>();
		if (list == null) {
			return output;
		}
		Cell c;
		for (T o : list) {
			if (o == null) {
				c = new Cell(null, CellType.EMPTY);
			}
			else if (o instanceof Cell) {
				c = (Cell) o;
			}
			else if (o instanceof Boolean) {
				c = new Cell(o, CellType.BOOL);
			}
			else if (o instanceof Byte) {
				c = new Cell(o, CellType.NUMBER);
			}
			else if (o instanceof BigDecimal) {
				c = new Cell(o, CellType.NUMBER);
			}
			else if (o instanceof Double) {
				c = new Cell(o, CellType.NUMBER);
			}
			else if (o instanceof Float) {
				c = new Cell(o, CellType.NUMBER);
			}
			else if (o instanceof Integer) {
				c = new Cell(o, CellType.NUMBER);
			}
			else if (o instanceof Long) {
				c = new Cell(o, CellType.NUMBER);
			}
			else if (o instanceof Short) {
				c = new Cell(o, CellType.NUMBER);
			}
			else if (o instanceof Date) {
				c = new Cell(o, CellType.DATE);
				c.setStyle(BasicStyles.DateFormat());
			}
			else if (o instanceof LocalTime) {
				c = new Cell(o, CellType.TIME);
				c.setStyle(BasicStyles.TimeFormat());
			}
			else if (o instanceof String) {
				c = new Cell(o, CellType.STRING);
			}
			else {
				c = new Cell(o.toString(), CellType.DEFAULT);
			}
			output.add(c);
		}
		return output;
	}

	/**
	 * Gets a list of cell addresses from a cell range (format A1:B3 or
	 * AAD556:AAD1000)
	 *
	 * @param range
	 *            Range to process
	 * @return List of cell addresses
	 * @throws FormatException
	 *             Throws a FormatException if a part of the passed range is
	 *             malformed
	 * @throws RangeException
	 *             Throws a RangeException if the range is out of range (A-XFD and 1
	 *             to 1048576)
	 */
	public static List<Address> getCellRange(String range) {
		Range range2 = resolveCellRange(range);
		return getCellRange(range2.StartAddress, range2.EndAddress);
	}

	/**
	 * Get a list of cell addresses from a cell range
	 *
	 * @param startAddress
	 *            Start address as string in the format A1 - XFD1048576
	 * @param endAddress
	 *            End address as string in the format A1 - XFD1048576
	 * @return List of cell addresses
	 * @throws FormatException
	 *             Throws a FormatException if a part of the passed range is
	 *             malformed
	 * @throws RangeException
	 *             Throws a RangeException if the range is out of range (A-XFD and 1
	 *             to 1048576)
	 */
	public static List<Address> getCellRange(String startAddress, String endAddress) {
		Address start = resolveCellCoordinate(startAddress);
		Address end = resolveCellCoordinate(endAddress);
		return getCellRange(start, end);
	}

	/**
	 * Get a list of cell addresses from a cell range
	 *
	 * @param startColumn
	 *            Start column (zero based)
	 * @param startRow
	 *            Start roe (zero based)
	 * @param endColumn
	 *            End column (zero based)
	 * @param endRow
	 *            End row (zero based)
	 * @return List of cell addresses
	 * @throws RangeException
	 *             Throws a RangeException if the value of one passed address parts
	 *             is out of range (A-XFD and 1 to 1048576)
	 */
	public static List<Address> getCellRange(int startColumn, int startRow, int endColumn, int endRow) {
		Address start = new Address(startColumn, startRow);
		Address end = new Address(endColumn, endRow);
		return getCellRange(start, end);
	}

	/**
	 * Get a list of cell addresses from a cell range
	 *
	 * @param startAddress
	 *            Start address
	 * @param endAddress
	 *            End address
	 * @return List of cell addresses
	 * @throws FormatException
	 *             Throws a FormatException if a part of the passed addresses is
	 *             malformed
	 * @throws RangeException
	 *             Throws a RangeException if the value of one passed address is out
	 *             of range (A-XFD and 1 to 1048576)
	 */
	public static List<Address> getCellRange(Address startAddress, Address endAddress) {
		int startColumn, endColumn, startRow, endRow;
		if (startAddress.Column < endAddress.Column) {
			startColumn = startAddress.Column;
			endColumn = endAddress.Column;
		}
		else {
			startColumn = endAddress.Column;
			endColumn = startAddress.Column;
		}
		if (startAddress.Row < endAddress.Row) {
			startRow = startAddress.Row;
			endRow = endAddress.Row;
		}
		else {
			startRow = endAddress.Row;
			endRow = startAddress.Row;
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
	 * Gets the address of a cell by the column and row number (zero based)
	 *
	 * @param column
	 *            Column address of the cell (zero-based)
	 * @param row
	 *            Row address of the cell (zero-based)
	 * @return Cell Address as string in the format A1 - XFD1048576
	 * @throws RangeException
	 *             Throws a RangeException if the start or end address was out of
	 *             range
	 */
	public static String resolveCellAddress(int column, int row) {
		return resolveCellAddress(column, row, AddressType.Default);
	}

	/**
	 * Gets the address of a cell by the column and row number (zero based)
	 *
	 * @param column
	 *            Column address of the cell (zero-based)
	 * @param row
	 *            Row address of the cell (zero-based)
	 * @param type
	 *            Referencing type of the address
	 * @return Cell Address as string in the format A1 - XFD1048576. Depending on
	 *         the type, Addresses like '$A55', 'B$2' or '$A$5' are possible outputs
	 * @throws RangeException
	 *             Throws a RangeException if the start or end address was out of
	 *             range
	 */
	public static String resolveCellAddress(int column, int row, AddressType type) {
		validateColumnNumber(column);
		validateRowNumber(row);
		switch (type) {
			case FixedRowAndColumn:
				return "$" + resolveColumnAddress(column) + "$" + (row + 1);
			case FixedColumn:
				return "$" + resolveColumnAddress(column) + (row + 1);
			case FixedRow:
				return resolveColumnAddress(column) + "$" + (row + 1);
			default:
				return resolveColumnAddress(column) + (row + 1);
		}
	}

	/**
	 * Gets the column and row number (zero based) of a cell by the address
	 *
	 * @param address
	 *            Address as string in the format A1 - XFD1048576. '$' signs
	 *            indicating fixed rows and / or columns are considered
	 * @return Address object of the passed string
	 * @throws FormatException
	 *             Throws a FormatException if the passed address is malformed
	 * @throws RangeException
	 *             Throws a RangeException if the value of the passed address is out
	 *             of range (A-XFD and 1 to 1048576)
	 */
	public static Address resolveCellCoordinate(String address) {
		int row, column;
		if (Helper.isNullOrEmpty(address)) {
			throw new FormatException("The cell address is null or empty and could not be resolved");
		}
		address = address.toUpperCase();
		Pattern pattern = Pattern.compile("(^(\\$?)([A-Z]{1,3})(\\$?)([0-9]{1,7})$)"); //
		Matcher matcher = pattern.matcher(address);
		if (matcher.matches() && matcher.groupCount() == 5) {
			int digits = Integer.parseInt(matcher.group(5));
			column = resolveColumn(matcher.group(3));
			row = digits - 1;
			validateRowNumber(row);
			if (!matcher.group(2).isEmpty() && !matcher.group(4).isEmpty()) {
				return new Address(column, row, AddressType.FixedRowAndColumn);
			}
			else if (!matcher.group(2).isEmpty() && matcher.group(4).isEmpty()) {
				return new Address(column, row, AddressType.FixedColumn);
			}
			else if (matcher.group(2).isEmpty() && !matcher.group(4).isEmpty()) {
				return new Address(column, row, AddressType.FixedRow);
			}
			else {
				return new Address(column, row, AddressType.Default);
			}
		}
		else {
			throw new FormatException("The format of the cell address (" + address + ") is malformed");
		}
	}

	/**
	 * Resolves a cell range from the format like A1:B3 or AAD556:AAD1000
	 *
	 * @param range
	 *            Range to process
	 * @return Range object of the passed string range
	 * @throws FormatException
	 *             Thrown if the passed range is malformed
	 */
	public static Range resolveCellRange(String range) {
		if (Helper.isNullOrEmpty(range)) {
			throw new FormatException("The cell range is null or empty and could not be resolved");
		}
		String[] split = range.split(":");
		if (split.length != 2) {
			throw new FormatException("The cell range (" + range + ") is malformed and could not be resolved");
		}
		try {
			Address start = resolveCellCoordinate(split[0]);
			Address end = resolveCellCoordinate(split[1]);
			return new Range(start, end);
		}
		catch (Exception e) {
			throw new FormatException("The start address or end address could not be resolved. See inner exception", e);
		}
	}

	/**
	 * Gets the column number from the column address (A - XFD)
	 *
	 * @param columnAddress
	 *            Column address (A - XFD)
	 * @return Column number (zero-based)
	 * @throws RangeException
	 *             Thrown if the column is out of range
	 */
	public static int resolveColumn(String columnAddress) {
		if (Helper.isNullOrEmpty(columnAddress)) {
			throw new RangeException("The passed address was null or empty");
		}
		columnAddress = columnAddress.toUpperCase();
		int chr;
		int result = 0;
		int multiplier = 1;

		for (int i = columnAddress.length() - 1; i >= 0; i--) {
			chr = columnAddress.charAt(i);
			chr = chr - 64;
			result = result + (chr * multiplier);
			multiplier = multiplier * 26;
		}
		validateColumnNumber(result - 1);
		return result - 1;
	}

	/**
	 * Gets the column address (A - XFD)
	 *
	 * @param columnNumber
	 *            Column number (zero-based)
	 * @return Column address (A - XFD)
	 * @throws RangeException
	 *             Thrown if the passed column number is out of range
	 */
	public static String resolveColumnAddress(int columnNumber) {
		validateColumnNumber(columnNumber);
		// A - XFD
		int j = 0;
		int k = 0;
		int l = 0;
		StringBuilder sb = new StringBuilder();
		for (int i = 0; i <= columnNumber; i++) {
			if (j > 25) {
				k++;
				j = 0;
			}
			if (k > 25) {
				l++;
				k = 0;
			}
			j++;
		}
		if (l > 0) {
			sb.append((char) (l + 64));
		}
		if (k > 0) {
			sb.append((char) (k + 64));
		}
		sb.append((char) (j + 64));
		return sb.toString();
	}

	/**
	 * Gets the scope of the passed address (string expression). Scope means either
	 * single cell address or range
	 *
	 * @param addressExpression
	 *            Address expression
	 * @return Scope of the address expression
	 */
	public static AddressScope getAddressScope(String addressExpression) {
		try {
			resolveCellCoordinate(addressExpression);
			return AddressScope.SingleAddress;
		}
		catch (Exception ex) {
			try {
				resolveCellRange(addressExpression);
				return AddressScope.Range;
			}
			catch (Exception ex2) {
				return AddressScope.Invalid;
			}
		}

	}

	/**
	 * Validates the passed (zero-based) column number. an exception will be thrown
	 * if the column is invalid
	 *
	 * @param columnNumber
	 *            Number to check
	 * @throws RangeException
	 *             Thrown if the passed column number is out of range
	 */
	static void validateColumnNumber(int columnNumber) {
		if (columnNumber > Worksheet.MAX_COLUMN_NUMBER || columnNumber < Worksheet.MIN_COLUMN_NUMBER) {
			throw new RangeException("The column number (" +
					columnNumber +
					") is out of range. Range is from " +
					Worksheet.MIN_COLUMN_NUMBER +
					" to " +
					Worksheet.MAX_COLUMN_NUMBER +
					" (" +
					(Worksheet.MAX_COLUMN_NUMBER + 1) +
					" columns).");
		}
	}

	/**
	 * Validates the passed (zero-based) row number. an exception will be thrown if
	 * the row is invalid
	 *
	 * @param rowNumber
	 *            Number to check
	 * @throws RangeException
	 *             Thrown if the passed row number is out of range
	 */
	static void validateRowNumber(int rowNumber) {
		if (rowNumber > Worksheet.MAX_ROW_NUMBER || rowNumber < Worksheet.MIN_ROW_NUMBER) {
			throw new RangeException("The row number (" +
					rowNumber +
					") is out of range. Range is from " +
					Worksheet.MIN_ROW_NUMBER +
					" to " +
					Worksheet.MAX_ROW_NUMBER +
					" (" +
					(Worksheet.MAX_ROW_NUMBER + 1) +
					" rows).");
		}
	}

}
