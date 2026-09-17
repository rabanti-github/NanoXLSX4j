/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import java.math.BigDecimal;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;

/**
 * Class for handling of basic Excel formulas
 */
public class BasicFormulas {

    /**
     * Returns a cell with an average formula
     *
     * @param range Cell range to apply the average operation to
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell average(Range range) {
        return average(null, range);
    }

    /**
     * Returns a cell with an average formula
     *
     * @param target Target worksheet of the average operation. Can be null if on the same worksheet
     * @param range  Cell range to apply the average operation to
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell average(Worksheet target, Range range) {
        return getBasicFormula(target, range, "AVERAGE", null);
    }

    /**
     * Returns a cell with a ceil formula
     *
     * @param address  Address to apply the ceil operation to
     * @param decimals Number of decimals (digits)
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell ceil(Address address, int decimals) {
        return ceil(null, address, decimals);
    }

    /**
     * Returns a cell with a ceil formula
     *
     * @param target   Target worksheet of the ceil operation. Can be null if on the same worksheet
     * @param address  Address to apply the ceil operation to
     * @param decimals Number of decimals (digits)
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell ceil(Worksheet target, Address address, int decimals) {
        return getBasicFormula(target, new Range(address, address), "ROUNDUP", ParserUtils.toString(decimals));
    }

    /**
     * Returns a cell with a floor formula
     *
     * @param address  Address to apply the floor operation to
     * @param decimals Number of decimals (digits)
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell floor(Address address, int decimals) {
        return floor(null, address, decimals);
    }

    /**
     * Returns a cell with a floor formula
     *
     * @param target   Target worksheet of the floor operation. Can be null if on the same worksheet
     * @param address  Address to apply the floor operation to
     * @param decimals Number of decimals (digits)
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell floor(Worksheet target, Address address, int decimals) {
        return getBasicFormula(target, new Range(address, address), "ROUNDDOWN", ParserUtils.toString(decimals));
    }

    /**
     * Returns a cell with a max formula
     *
     * @param range Cell range to apply the max operation to
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell max(Range range) {
        return max(null, range);
    }

    /**
     * Returns a cell with a max formula
     *
     * @param target Target worksheet of the max operation. Can be null if on the same worksheet
     * @param range  Cell range to apply the max operation to
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell max(Worksheet target, Range range) {
        return getBasicFormula(target, range, "MAX", null);
    }

    /**
     * Returns a cell with a median formula
     *
     * @param range Cell range to apply the median operation to
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell median(Range range) {
        return median(null, range);
    }

    /**
     * Returns a cell with a median formula
     *
     * @param target Target worksheet of the median operation. Can be null if on the same worksheet
     * @param range  Cell range to apply the median operation to
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell median(Worksheet target, Range range) {
        return getBasicFormula(target, range, "MEDIAN", null);
    }

    /**
     * Returns a cell with a min formula
     *
     * @param range Cell range to apply the min operation to
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell min(Range range) {
        return min(null, range);
    }

    /**
     * Returns a cell with a min formula
     *
     * @param target Target worksheet of the min operation. Can be null if on the same worksheet
     * @param range  Cell range to apply the median operation to
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell min(Worksheet target, Range range) {
        return getBasicFormula(target, range, "MIN", null);
    }

    /**
     * Returns a cell with a round formula
     *
     * @param address  Address to apply the round operation to
     * @param decimals Number of decimals (digits)
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell round(Address address, int decimals) {
        return round(null, address, decimals);
    }

    /**
     * Returns a cell with a round formula
     *
     * @param target   Target worksheet of the round operation. Can be null if on the same worksheet
     * @param address  Address to apply the round operation to
     * @param decimals Number of decimals (digits)
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell round(Worksheet target, Address address, int decimals) {
        return getBasicFormula(target, new Range(address, address), "ROUND", ParserUtils.toString(decimals));
    }

    /**
     * Returns a cell with a sum formula
     *
     * @param range Cell range to get a sum of
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell sum(Range range) {
        return sum(null, range);
    }

    /**
     * Returns a cell with a sum formula
     *
     * @param target Target worksheet of the sum operation. Can be null if on the same worksheet
     * @param range  Cell range to get a sum of
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    public static Cell sum(Worksheet target, Range range) {
        return getBasicFormula(target, range, "SUM", null);
    }

    /**
     * Function to generate a VLOOKUP Excel function
     *
     * @param number      Numeric value for the lookup. Valid types are byte, short, int, long, float, double and
     *                    BigDecimal
     * @param range       Matrix of the lookup
     * @param columnIndex Column index of the target column within the range (1 based)
     * @param exactMatch  If true, an exact match is applied to the lookup
     * @return Prepared Cell object, ready to be added to a worksheet
     * @throws FormatException Thrown if the value or column index is invalid
     */
    public static Cell vLookup(Object number, Range range, int columnIndex, boolean exactMatch) {
        return vLookup(number, null, range, columnIndex, exactMatch);
    }

    /**
     * Function to generate a VLOOKUP Excel function
     *
     * @param number      Numeric value for the lookup. Valid types are byte, short, int, long, float, double and
     *                    BigDecimal
     * @param rangeTarget Target worksheet of the matrix. Can be null if on the same worksheet
     * @param range       Matrix of the lookup
     * @param columnIndex Column index of the target column within the range (1 based)
     * @param exactMatch  If true, an exact match is applied to the lookup
     * @return Prepared Cell object, ready to be added to a worksheet
     * @throws FormatException Thrown if the value or column index is invalid
     */
    public static Cell vLookup(
            Object number, Worksheet rangeTarget, Range range, int columnIndex, boolean exactMatch) {
        return getVLookup(null, null, number, rangeTarget, range, columnIndex, exactMatch, true);
    }

    /**
     * Function to generate a VLOOKUP Excel function
     *
     * @param address     Query address of a cell as source of the lookup
     * @param range       Matrix of the lookup
     * @param columnIndex Column index of the target column within the range (1 based)
     * @param exactMatch  If true, an exact match is applied to the lookup
     * @return Prepared Cell object, ready to be added to a worksheet
     * @throws FormatException Thrown if the column index is invalid
     */
    public static Cell vLookup(Address address, Range range, int columnIndex, boolean exactMatch) {
        return vLookup(null, address, null, range, columnIndex, exactMatch);
    }

    /**
     * Function to generate a VLOOKUP Excel function
     *
     * @param queryTarget Target worksheet of the query argument. Can be null if on the same worksheet
     * @param address     Query address of a cell as source of the lookup
     * @param rangeTarget Target worksheet of the matrix. Can be null if on the same worksheet
     * @param range       Matrix of the lookup
     * @param columnIndex Column index of the target column within the range (1 based)
     * @param exactMatch  If true, an exact match is applied to the lookup
     * @return Prepared Cell object, ready to be added to a worksheet
     * @throws FormatException Thrown if the column index is invalid
     */
    public static Cell vLookup(
            Worksheet queryTarget,
            Address address,
            Worksheet rangeTarget,
            Range range,
            int columnIndex,
            boolean exactMatch
    ) {
        return getVLookup(queryTarget, address, 0, rangeTarget, range, columnIndex, exactMatch, false);
    }

    /**
     * Function to generate a VLOOKUP Excel function
     *
     * @param queryTarget   Target worksheet of the query argument. Can be null if on the same worksheet
     * @param address       In case of a reference lookup, query address of a cell
     * @param number        In case of a numeric lookup, number for the lookup
     * @param rangeTarget   Target worksheet of the matrix. Can be null if on the same worksheet
     * @param range         Matrix of the lookup
     * @param columnIndex   Column index of the target column within the range (1 based)
     * @param exactMatch    If true, an exact match is applied to the lookup
     * @param numericLookup If true, the lookup is numeric, otherwise it is a cell reference
     * @return Prepared Cell object, ready to be added to a worksheet
     * @throws FormatException Thrown if the value or column index is invalid
     */
    private static Cell getVLookup(
            Worksheet queryTarget,
            Address address,
            Object number,
            Worksheet rangeTarget,
            Range range,
            int columnIndex,
            boolean exactMatch,
            boolean numericLookup
    ) {
        int rangeWidth = Math.abs(range.endAddress().column() - range.startAddress().column()) + 1;
        if (columnIndex < 1 || columnIndex > rangeWidth) {
            throw new FormatException(
                    "The column index on range " + range + " can only be between 1 and " + rangeWidth
            );
        }

        String lookupArgument;
        if (numericLookup) {
            lookupArgument = getNumericLookupArgument(number);
        } else if (queryTarget != null) {
            lookupArgument = queryTarget.getSheetName() + "!" + address;
        } else {
            lookupArgument = address.toString();
        }

        String rangeArgument;
        if (rangeTarget != null) {
            rangeArgument = rangeTarget.getSheetName() + "!" + range;
        } else {
            rangeArgument = range.toString();
        }

        String matchArgument = exactMatch ? "TRUE" : "FALSE";
        return new Cell(
                "VLOOKUP(" + lookupArgument + "," + rangeArgument + "," +
                        ParserUtils.toString(columnIndex) + "," + matchArgument + ")",
                Cell.CellType.FORMULA
        );
    }

    private static String getNumericLookupArgument(Object number) {
        if (number == null) {
            throw new FormatException(
                    "The lookup variable can only be a cell address or a numeric value. The passed value was null."
            );
        }
        if (number instanceof Byte value) {
            return ParserUtils.toString(value);
        } else if (number instanceof BigDecimal value) {
            return ParserUtils.toString(value);
        } else if (number instanceof Double value) {
            return ParserUtils.toString(value);
        } else if (number instanceof Float value) {
            return ParserUtils.toString(value);
        } else if (number instanceof Integer value) {
            return ParserUtils.toString(value);
        } else if (number instanceof Long value) {
            return ParserUtils.toString(value);
        } else if (number instanceof Short value) {
            return ParserUtils.toString(value);
        }
        throw new FormatException(
                "The lookup variable can only be a cell address or a numeric value. The value '" + number +
                        "' is invalid."
        );
    }

    /**
     * Function to generate a basic Excel function with one cell range as parameter and an optional post argument
     *
     * @param target       Target worksheet of the cell reference. Can be null if on the same worksheet
     * @param range        Main argument as cell range. If applied on one cell, the start and end address are identical
     * @param functionName Internal Excel function name
     * @param postArg      Optional argument
     * @return Prepared Cell object, ready to be added to a worksheet
     */
    private static Cell getBasicFormula(Worksheet target, Range range, String functionName, String postArg) {
        String arg1;
        String arg2;
        String prefix;
        if (postArg == null) {
            arg2 = "";
        } else {
            arg2 = "," + postArg;
        }
        if (target != null) {
            prefix = target.getSheetName() + "!";
        } else {
            prefix = "";
        }
        if (range.startAddress().equals(range.endAddress())) {
            arg1 = prefix + range.startAddress();
        } else {
            arg1 = prefix + range;
        }
        return new Cell(functionName + "(" + arg1 + arg2 + ")", Cell.CellType.FORMULA);
    }

    //------------------
    private BasicFormulas() {
        // Do not instantiate
    }
}
