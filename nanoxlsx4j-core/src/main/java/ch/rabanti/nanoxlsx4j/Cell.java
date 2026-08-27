package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;

import java.util.ArrayList;
import java.util.List;

public class Cell {

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
    public enum AddressScope
    {
        /** The address represents a single cell or a range of cells */
        ANY,
        /** The address represents a single cell */
        SINGLE_ADDRESS,
         /** The address represents a range of cells */
        RANGE,
         /** The address expression is invalid */
        INVALID
    }

    // TODO can this be made an immutable list
    /**
     * Get a list of cell addresses from a cell range
     *
     * @param startAddress Start address
     * @param endAddress End address
     * @return List of cell addresses
     * @throws FormatException Throws a FormatException if a part of the passed addresses is malformed
     * @throws RangeException Throws a RangeException if the value of one passed address is out of range (A-XFD and 1 to 1048576)
     */
    public static List<Address> getCellRange(Address startAddress, Address endAddress)
    {
        int startColumn;
        int endColumn;
        int startRow;
        int endRow;
        if (startAddress.column() < endAddress.column())
        {
            startColumn = startAddress.column();
            endColumn = endAddress.column();
        }
        else
        {
            startColumn = endAddress.column();
            endColumn = startAddress.column();
        }
        if (startAddress.row() < endAddress.row())
        {
            startRow = startAddress.row();
            endRow = endAddress.row();
        }
        else
        {
            startRow = endAddress.row();
            endRow = startAddress.row();
        }
        List<Address> output = new ArrayList<>();
        for (int column = startColumn; column <= endColumn; column++)
        {
            for (int row = startRow; row <= endRow; row++)
            {
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
     * Gets the column and row number (zero based) of a cell by the address
     *
     * @param address Address as string in the format A1 - XFD1048576
     * @return Struct with row and column
     * @throws FormatException Throws a FormatException if the passed address is malformed
     * @throws RangeException Throws a RangeException if the value of the passed address is out of range (A-XFD and 1 to 1048576)
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
     * @throws RangeException Throws a RangeException if the range is out of range (A-XFD and 1 to 1048576)
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

    private static boolean isAsciiLetter(char character) {
        return character >= 'A' && character <= 'Z' || character >= 'a' && character <= 'z';
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

}
