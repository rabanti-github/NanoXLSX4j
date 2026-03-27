/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.utils;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

/**
 * Class providing general validator methods
 *
 * @author Raphael Stoeckli
 */
public class Validators {

    private Validators() {
        // Prevents class instantiation
    }

    /**
     * Validates the passed string, whether it is a valid RGB or ARGB value.
     * The method automatically tries to validate for ARGB (8 characters) first, then for RGB (6 characters).
     *
     * @param hexCode Hex string to check
     * @throws StyleException thrown if the value is neither a valid ARGB nor RGB hex string
     */
    public static void validateGenericColor(String hexCode) {
        validateGenericColor(hexCode, false);
    }

    /**
     * Validates the passed string, whether it is a valid RGB or ARGB value.
     * The method automatically tries to validate for ARGB (8 characters) first, then for RGB (6 characters).
     *
     * @param hexCode    Hex string to check
     * @param allowEmpty If true, null or empty is accepted as valid
     * @throws StyleException thrown if the value is neither a valid ARGB nor RGB hex string
     */
    public static void validateGenericColor(String hexCode, boolean allowEmpty) {
        String argbMessage = validateColorInternal(hexCode, true, allowEmpty);
        if (argbMessage != null) {
            String rgbMessage = validateColorInternal(hexCode, false, allowEmpty);
            if (rgbMessage != null) {
                throw new StyleException(argbMessage);
            }
        }
    }

    /**
     * Validates the passed string, whether it is a valid RGB value (6 characters) or ARGB value (8 characters)
     *
     * @param hexCode  Hex string to check
     * @param useAlpha If true, two additional characters (total 8) are expected as alpha value
     * @throws StyleException thrown if an invalid hex value is passed
     */
    public static void validateColor(String hexCode, boolean useAlpha) {
        validateColor(hexCode, useAlpha, false);
    }

    /**
     * Validates the passed string, whether it is a valid RGB value (6 characters) or ARGB value (8 characters)
     *
     * @param hexCode    Hex string to check
     * @param useAlpha   If true, two additional characters (total 8) are expected as alpha value
     * @param allowEmpty If true, null or empty is accepted as valid
     * @throws StyleException thrown if an invalid hex value is passed
     */
    public static void validateColor(String hexCode, boolean useAlpha, boolean allowEmpty) {
        String message = validateColorInternal(hexCode, useAlpha, allowEmpty);
        if (message != null) {
            throw new StyleException(message);
        }
    }

    private static String validateColorInternal(String hexCode, boolean useAlpha, boolean allowEmpty) {
        if (hexCode == null || hexCode.isEmpty()) {
            if (allowEmpty) {
                return null;
            }
            return "The color expression cannot be null or empty";
        }
        int length = useAlpha ? 8 : 6;
        if (hexCode.length() != length) {
            return "The value '" + hexCode + "' is invalid. A valid value must contain " + length + " hex characters";
        }
        if (!hexCode.matches("[a-fA-F0-9]{6,8}")) {
            return "The expression '" + hexCode + "' is not a valid hex value";
        }
        return null;
    }

}
