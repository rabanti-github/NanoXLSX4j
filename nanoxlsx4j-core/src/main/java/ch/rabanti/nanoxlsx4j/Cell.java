package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.RangeException;

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
