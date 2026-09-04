/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils;

import java.util.regex.Pattern;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

public class Validators {

    private static final Pattern HEX_COLOR_PATTERN = Pattern.compile("[a-fA-F0-9]{6,8}");
    private static final Pattern INVALID_WORKSHEET_NAME_PATTERN = Pattern.compile("[\\[\\]*?/\\\\]");

    private Validators() {
        // Do not instantiate
    }

    /**
     * Validates the passed string, whether it is a valid RGB or ARGB value that can be used for Fills, Fonts or other
     * styling components. The method automatically tries to validate for ARGB (8 characters) first, then for RGB (6
     * characters).
     *
     * @param hexCode Hex string to check (no empty values allowed)
     * @throws StyleException Thrown if an invalid hex value is passed
     */
    public static void validateGenericColor(String hexCode) {
        validateGenericColor(hexCode, false);
    }

    /**
     * Validates the passed string, whether it is a valid RGB or ARGB value that can be used for Fills, Fonts or other
     * styling components. The method automatically tries to validate for ARGB (8 characters) first, then for RGB (6
     * characters).
     *
     * @param hexCode    Hex string to check
     * @param allowEmpty If true, null or empty values are allowed as valid values
     * @throws StyleException Thrown if an invalid hex value is passed
     */
    public static void validateGenericColor(String hexCode, boolean allowEmpty) {
        String argbMessage = validateColorInternal(hexCode, true, allowEmpty);
        String rgbMessage = null;
        if (argbMessage != null) {
            rgbMessage = validateColorInternal(hexCode, false, allowEmpty);
            if (rgbMessage != null) {
                throw new StyleException(argbMessage);
            }
        }
    }

    /**
     * Validates the passed string, whether it is a valid RGB or ARGB value that can be used for Fills, Fonts or other
     * styling components
     *
     * @param hexCode  Hex string to check (no empty values allowed)
     * @param useAlpha If true, two additional characters (total 8) are expected as alpha value
     * @throws StyleException Thrown if an invalid hex value is passed
     */
    public static void validateColor(String hexCode, boolean useAlpha) {
        validateColor(hexCode, useAlpha, false);
    }

    /**
     * Validates the passed string, whether it is a valid RGB or ARGB value that can be used for Fills, Fonts or other
     * styling components
     *
     * @param hexCode    Hex string to check
     * @param useAlpha   If true, two additional characters (total 8) are expected as alpha value
     * @param allowEmpty If true, null or empty values are allowed as valid values
     * @throws StyleException Thrown if an invalid hex value is passed
     */
    public static void validateColor(String hexCode, boolean useAlpha, boolean allowEmpty) {
        String message = validateColorInternal(hexCode, useAlpha, allowEmpty);
        if (message != null) {
            throw new StyleException(message);
        }
    }

    /**
     * Validates the passed string, whether it is a valid single cell address or cell range. The address or range can
     * contain modifier characters ({@link Cell.AddressType})
     *
     * @param expression The address expression to validate
     * @throws FormatException A format exception is thrown if the passed address is not a valid cell address or range
     */
    public static void validateCellAddressExpression(String expression) {
        validateCellAddressExpression(expression, Cell.AddressScope.ANY);
    }

    /**
     * Validates the passed string, whether it is a valid single cell address or cell range. The address or range can
     * contain modifier characters ({@link Cell.AddressType})
     *
     * <p>Remarks: If {@code scope} is {@link Cell.AddressScope#RANGE}, an explicit range expression is required; a
     * single address is rejected even though {@link Cell#resolveCellRange(String)} can represent it as a one-cell
     * range. If the scope is {@link Cell.AddressScope#INVALID}, the validation is inverted, so that a valid cell or
     * range will throw an exception.</p>
     *
     * @param expression The address expression to validate
     * @param scope      Optional parameter to validate for a specific address scope (Any, SingleAddress, Range).
     *                   Default is: Any
     * @throws FormatException A format exception is thrown if the passed address is not a valid cell address or range
     */
    public static void validateCellAddressExpression(String expression, Cell.AddressScope scope) {
        boolean isCellAddress = false;
        boolean isRange = false;
        Exception lastException = null;
        try {
            Cell.resolveCellCoordinate(expression);
            isCellAddress = true;
        } catch (Exception ex) {
            if (scope == Cell.AddressScope.SINGLE_ADDRESS) {
                throw new FormatException(ex.getMessage(), ex); // No further checks necessary
            }
            lastException = ex;
        }
        try {
            Cell.resolveCellRange(expression);
            isRange = true;
        } catch (Exception ex) {
            if (scope == Cell.AddressScope.RANGE) {
                throw new FormatException(ex.getMessage(), ex); // No further checks necessary
            }
            lastException = ex;
        }
        if (scope == Cell.AddressScope.RANGE && isCellAddress) {
            FormatException innerException = new FormatException(
                    "The expression (" + expression + ") is a single cell address, but a cell range was expected");
            throw new FormatException(innerException.getMessage(), innerException);
        } else if (scope == Cell.AddressScope.ANY && !isCellAddress && !isRange) {
            throw new FormatException(lastException.getMessage(), lastException); // Not a cell or range
        } else if (scope == Cell.AddressScope.INVALID && (isCellAddress || isRange)) {
            throw new FormatException(
                    "The passed expression is valid cell address or range, but the validation was explicitly inverted");
        }
    }

    /**
     * Validates the passed string, whether it is an expression that can be used as worksheet name.
     *
     * @param name Name to validate
     * @throws FormatException Thrown if the worksheet name is too long or contains illegal characters
     */
    public static void validateWorksheetName(String name) {
        if (ParserUtils.isNullOrEmpty(name) || name.length() > Worksheet.MAX_WORKSHEET_NAME_LENGTH) {
            throw new FormatException(
                    "the worksheet name must be between 1 and " + Worksheet.MAX_WORKSHEET_NAME_LENGTH + " characters");
        }
        if (INVALID_WORKSHEET_NAME_PATTERN.matcher(name).find()) {
            throw new FormatException("the worksheet name must not contain the characters [  ]  * ? / \\ ");
        }
    }

    /**
     * Validates the passed string, whether it is a valid RGB or ARGB value that can be used for Fills, Fonts or other
     * styling components.
     *
     * @param hexCode    Hex string to check
     * @param useAlpha   If true, two additional characters (total 8) are expected as alpha value
     * @param allowEmpty If true, null or empty values are allowed as valid values
     * @return Null, if valid, otherwise, the specific exception message is returned
     */
    private static String validateColorInternal(String hexCode, boolean useAlpha, boolean allowEmpty) {
        if (ParserUtils.isNullOrEmpty(hexCode)) {
            if (allowEmpty) {
                return null;
            }
            return "The color expression cannot be null or empty";
        }

        int length = useAlpha ? 8 : 6;
        if (hexCode.length() != length) {
            return "The value '" + hexCode + "' is invalid. A valid value must contain " + length + " hex characters";
        }
        if (!HEX_COLOR_PATTERN.matcher(hexCode).matches()) {
            return "The expression '" + hexCode + "' is not a valid hex value";
        }
        return null;
    }
}
