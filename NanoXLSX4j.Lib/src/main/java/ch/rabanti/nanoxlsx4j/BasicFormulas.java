/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import java.math.BigDecimal;

import ch.rabanti.nanoxlsx4j.Cell.CellType;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

/**
 * Class for handling of basic Excel formulas
 *
 * @author Raphael Stoeckli
 */

public final class BasicFormulas {
	/**
	 * Returns a cell with a average formula
	 *
	 * @param range
	 *            Cell range to apply the average operation to
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Average(Range range) {
		return Average(null, range);
	}

	/**
	 * Returns a cell with a average formula
	 *
	 * @param target
	 *            Target worksheet of the average operation. Can be null if on the
	 *            same worksheet
	 * @param range
	 *            Cell range to apply the average operation to
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Average(Worksheet target, Range range) {
		return getBasicFormula(target, range, "AVERAGE", null);
	}

	/**
	 * Returns a cell with a ceil formula
	 *
	 * @param address
	 *            Address to apply the ceil operation to
	 * @param decimals
	 *            Number of decimals (digits)
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Ceil(Address address, int decimals) {
		return Ceil(null, address, decimals);
	}

	/**
	 * Returns a cell with a ceil formula
	 *
	 * @param target
	 *            Target worksheet of the ceil operation. Can be null if on the same
	 *            worksheet
	 * @param address
	 *            Address to apply the ceil operation to
	 * @param decimals
	 *            Number of decimals (digits)
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Ceil(Worksheet target, Address address, int decimals) {
		return getBasicFormula(target, new Range(address, address), "ROUNDUP", Integer.toString(decimals));
	}

	/**
	 * Returns a cell with a floor formula
	 *
	 * @param address
	 *            Address to apply the floor operation to
	 * @param decimals
	 *            Number of decimals (digits)
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Floor(Address address, int decimals) {
		return Floor(null, address, decimals);
	}

	/**
	 * Returns a cell with a floor formula
	 *
	 * @param target
	 *            Target worksheet of the floor operation. Can be null if on the
	 *            same worksheet
	 * @param address
	 *            Address to apply the floor operation to
	 * @param decimals
	 *            Number of decimals (digits)
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Floor(Worksheet target, Address address, int decimals) {
		return getBasicFormula(target, new Range(address, address), "ROUNDDOWN", Integer.toString(decimals));
	}

	/**
	 * Returns a cell with a max formula
	 *
	 * @param range
	 *            Cell range to apply the max operation to
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Max(Range range) {
		return Max(null, range);
	}

	/**
	 * Returns a cell with a max formula
	 *
	 * @param target
	 *            Target worksheet of the max operation. Can be null if on the same
	 *            worksheet
	 * @param range
	 *            Cell range to apply the max operation to
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Max(Worksheet target, Range range) {
		return getBasicFormula(target, range, "MAX", null);
	}

	/**
	 * Returns a cell with a median formula
	 *
	 * @param range
	 *            Cell range to apply the median operation to
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Median(Range range) {
		return Median(null, range);
	}

	/**
	 * Returns a cell with a median formula
	 *
	 * @param target
	 *            Target worksheet of the median operation. Can be null if on the
	 *            same worksheet
	 * @param range
	 *            Cell range to apply the median operation to
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Median(Worksheet target, Range range) {
		return getBasicFormula(target, range, "MEDIAN", null);
	}

	/**
	 * Returns a cell with a min formula
	 *
	 * @param range
	 *            Cell range to apply the min operation to
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Min(Range range) {
		return Min(null, range);
	}

	/**
	 * Returns a cell with a min formula
	 *
	 * @param target
	 *            Target worksheet of the min operation. Can be null if on the same
	 *            worksheet
	 * @param range
	 *            Cell range to apply the median operation to
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Min(Worksheet target, Range range) {
		return getBasicFormula(target, range, "MIN", null);
	}

	/**
	 * Returns a cell with a round formula
	 *
	 * @param address
	 *            Address to apply the round operation to
	 * @param decimals
	 *            Number of decimals (digits)
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Round(Address address, int decimals) {
		return Round(null, address, decimals);
	}

	/**
	 * Returns a cell with a round formula
	 *
	 * @param target
	 *            Target worksheet of the round operation. Can be null if on the
	 *            same worksheet
	 * @param address
	 *            Address to apply the round operation to
	 * @param decimals
	 *            Number of decimals (digits)
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Round(Worksheet target, Address address, int decimals) {
		return getBasicFormula(target, new Range(address, address), "ROUND", Integer.toString(decimals));
	}

	/**
	 * Returns a cell with a sum formula
	 *
	 * @param range
	 *            Cell range to get a sum of
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Sum(Range range) {
		return Sum(null, range);
	}

	/**
	 * Returns a cell with a sum formula
	 *
	 * @param target
	 *            Target worksheet of the sum operation. Can be null if on the same
	 *            worksheet
	 * @param range
	 *            Cell range to get a sum of
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	public static Cell Sum(Worksheet target, Range range) {
		return getBasicFormula(target, range, "SUM", null);
	}

	/**
	 * Function to generate a Vlookup as Excel function
	 *
	 * @param number
	 *            Numeric value for the lookup. Valid types are int, long, float and
	 *            double
	 * @param range
	 *            Matrix of the lookup
	 * @param columnIndex
	 *            Column index of the target column within the range (1 based)
	 * @param exactMatch
	 *            If true, an exact match is applied to the lookup
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 * @throws ch.rabanti.nanoxlsx4j.exceptions.RangeException
	 *             is thrown if the value or column index is invalid
	 */
	public static Cell VLookup(Object number, Range range, int columnIndex, boolean exactMatch) {
		return VLookup(number, null, range, columnIndex, exactMatch);
	}

	/**
	 * Function to generate a Vlookup as Excel function
	 *
	 * @param number
	 *            Numeric value for the lookup. Valid types are int, long, float and
	 *            double
	 * @param rangeTarget
	 *            Target worksheet of the matrix. Can be null if on the same
	 *            worksheet
	 * @param range
	 *            Matrix of the lookup
	 * @param columnIndex
	 *            Column index of the target column within the range (1 based)
	 * @param exactMatch
	 *            If true, an exact match is applied to the lookup
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 * @throws ch.rabanti.nanoxlsx4j.exceptions.RangeException
	 *             is thrown if the value or column index is invalid
	 */
	public static Cell VLookup(Object number, Worksheet rangeTarget, Range range, int columnIndex, boolean exactMatch) {
		return getVLookup(null, null, number, rangeTarget, range, columnIndex, exactMatch, true);
	}

	/**
	 * Function to generate a Vlookup as Excel function
	 *
	 * @param address
	 *            Query address of a cell as string as source of the lookup
	 * @param range
	 *            Matrix of the lookup
	 * @param columnIndex
	 *            Column index of the target column within the range (1 based)
	 * @param exactMatch
	 *            If true, an exact match is applied to the lookup
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 * @throws ch.rabanti.nanoxlsx4j.exceptions.RangeException
	 *             is thrown if the column index is invalid
	 */
	public static Cell VLookup(Address address, Range range, int columnIndex, boolean exactMatch) {
		return VLookup(null, address, null, range, columnIndex, exactMatch);
	}

	/**
	 * Function to generate a Vlookup as Excel function
	 *
	 * @param queryTarget
	 *            Target worksheet of the query argument. Can be null if on the same
	 *            worksheet
	 * @param address
	 *            Query address of a cell as string as source of the lookup
	 * @param rangeTarget
	 *            Target worksheet of the matrix. Can be null if on the same
	 *            worksheet
	 * @param range
	 *            Matrix of the lookup
	 * @param columnIndex
	 *            Column index of the target column within the range (1 based)
	 * @param exactMatch
	 *            If true, an exact match is applied to the lookup
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 * @throws ch.rabanti.nanoxlsx4j.exceptions.RangeException
	 *             is thrown if the column index is invalid
	 */
	public static Cell VLookup(Worksheet queryTarget, Address address, Worksheet rangeTarget, Range range, int columnIndex, boolean exactMatch) {
		return getVLookup(queryTarget, address, 0, rangeTarget, range, columnIndex, exactMatch, false);
	}

	/**
	 * Function to generate a Vlookup as Excel function
	 *
	 * @param queryTarget
	 *            Target worksheet of the query argument. Can be null if on the same
	 *            worksheet
	 * @param address
	 *            In case of a reference lookup, query address of a cell as string
	 * @param number
	 *            In case of a numeric lookup, number for the lookup
	 * @param rangeTarget
	 *            Target worksheet of the matrix. Can be null if on the same
	 *            worksheet
	 * @param range
	 *            Matrix of the lookup
	 * @param columnIndex
	 *            Column index of the target column within the range (1 based)
	 * @param exactMatch
	 *            If true, an exact match is applied to the lookup
	 * @param numericLookup
	 *            If true, the lookup is a numeric lookup, otherwise a reference
	 *            lookup
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 * @throws ch.rabanti.nanoxlsx4j.exceptions.RangeException
	 *             is thrown if the value or column index is invalid
	 * @apiNote Valid numeric values are: Byte, BigDecimal, Double, Float, Integer,
	 *          Long, Short and the corresponding primitive data types.<br>
	 *          Note: BigDecimal may lead to unexpected decimal places. They are
	 *          truncated 18 digits after the decimal point. Consider instantiate
	 *          BigDecimals wit string values to avoid such decimal places.
	 */
	private static Cell getVLookup(Worksheet queryTarget, Address address, Object number, Worksheet rangeTarget, Range range, int columnIndex, boolean exactMatch, boolean numericLookup) {
		int rangeWidth = range.EndAddress.Column - range.StartAddress.Column + 1;
		if (columnIndex < 1 || columnIndex > rangeWidth) {
			throw new FormatException("The column index on range " + range + " can only be between 1 and " + rangeWidth);
		}
		String arg1, arg2, arg3, arg4;
		if (numericLookup) {
			if (number == null) {
				throw new FormatException("The lookup variable can only be a cell address or a numeric value. The passed value was null.");
			}
			if (number instanceof Byte) {
				arg1 = Byte.toString((byte) number);
			}
			else if (number instanceof BigDecimal) {
				if (((BigDecimal) number).compareTo(BigDecimal.ZERO) == 0) {
					arg1 = "0";
				}
				else {
					arg1 = number.toString();
				}
			}
			else if (number instanceof Double) {
				if ((double) number == 0) {
					arg1 = "0";
				}
				else {
					arg1 = Double.toString((double) number);
				}
			}
			else if (number instanceof Float) {
				if ((float) number == 0) {
					arg1 = "0";
				}
				else {
					arg1 = Float.toString((float) number);
				}
			}
			else if (number instanceof Integer) {
				arg1 = Integer.toString((int) number);
			}
			else if (number instanceof Long) {
				arg1 = Long.toString((long) number);
			}
			else if (number instanceof Short) {
				arg1 = Short.toString((short) number);
			}
			else {
				throw new FormatException("The lookup variable can only be a cell address or a numeric value. The value '" + number + "' is invalid.");
			}
		}
		else {
			if (address == null) {
				throw new FormatException("The lookup variable can only be a cell address or a numeric value. The value '" + number + "' is invalid.");
			}
			if (queryTarget != null) {
				arg1 = queryTarget.getSheetName() + "!" + address;
			}
			else {
				arg1 = address.toString();
			}
		}
		arg1 = cutoffDecimals(arg1);
		if (rangeTarget != null) {
			arg2 = rangeTarget.getSheetName() + "!" + range;
		}
		else {
			arg2 = range.toString();
		}
		arg3 = Integer.toString(columnIndex);
		if (exactMatch) {
			arg4 = "TRUE";
		}
		else {
			arg4 = "FALSE";
		}
		return new Cell("VLOOKUP(" + arg1 + "," + arg2 + "," + arg3 + "," + arg4 + ")", CellType.FORMULA);
	}

	/**
	 * Cuts off decimals after the 18th digit
	 *
	 * @param value
	 * @return
	 */
	private static String cutoffDecimals(String value) {
		if (!value.contains(".")) {
			return value;
		}
		String[] parts = value.split("\\.");
		if (parts[1].length() > 18) {
			value = parts[0] + "." + parts[1].substring(0, 18);
		}
		return !value.contains(".") ? value : value.replaceAll("0*$", "").replaceAll("\\.$", "");
	}

	/**
	 * Function to generate a basic Excel function with one cell range as parameter
	 * and an optional post argument
	 *
	 * @param target
	 *            Target worksheet of the cell reference. Can be null if on the same
	 *            worksheet
	 * @param range
	 *            Main argument as cell range. If applied on one cell, the start and
	 *            end address are identical
	 * @param functionName
	 *            Internal Excel function name
	 * @param postArg
	 *            Optional argument
	 * @return Prepared Cell Object, ready to be added to a worksheet
	 */
	private static Cell getBasicFormula(Worksheet target, Range range, String functionName, String postArg) {
		String arg1, arg2, prefix;
		if (postArg == null) {
			arg2 = "";
		}
		else {
			arg2 = "," + postArg;
		}
		if (target != null) {
			prefix = target.getSheetName() + "!";
		}
		else {
			prefix = "";
		}
		if (range.StartAddress.equals(range.EndAddress)) {
			arg1 = prefix + range.StartAddress;
		}
		else {
			arg1 = prefix + range;
		}
		return new Cell(functionName + "(" + arg1 + arg2 + ")", CellType.FORMULA);
	}
}
