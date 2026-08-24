package ch.rabanti.nanoxlsx4j.utils;

import java.util.regex.Pattern;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

public class Validators {

    private static final Pattern HEX_COLOR_PATTERN = Pattern.compile("[a-fA-F0-9]{6,8}");

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
