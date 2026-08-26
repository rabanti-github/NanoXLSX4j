/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils;

import java.util.Locale;

public class ParserUtils {
    private ParserUtils() {
        // Do not instantiate
    }

    /**
     * Constant for number conversion. The invariant culture / Locale (represents mostly the US numbering scheme)
     * ensures that no culture-specific punctuations are used when converting numbers to strings, This is especially
     * important for OOXML number values. See also: <a
     * href="https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.invariantculture?view=net-5
     * .0">
     * https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.invariantculture?view=net-5.0</a>
     */
    public static final Locale INVARIANT_CULTURE = Locale.ROOT;

    /**
     * Transforms an integer to an invariant sting
     * <p>
     * This method is mainly a wrapper method but is kept to ensure a similar internal API between NanoXLSX and
     * NanoXLSX4j
     * </p>
     *
     * @param input Integer to transform
     * @return Integer as string
     */
    public static String toString(int input) {
        return Integer.toString(input);
    }

    /**
     * Transforms a string to upper case with null check and invariant culture
     *
     * @param input String to transform
     * @return Upper case string
     */
    public static String toUpper(String input) {
        return !isNullOrEmpty(input) ? input.toUpperCase(INVARIANT_CULTURE) : input;
    }

    /**
     * Transforms a string to lower case with null check and invariant culture
     *
     * @param input String to transform
     * @return Lower case string
     */
    public static String toLower(String input) {
        return !isNullOrEmpty(input) ? input.toLowerCase(INVARIANT_CULTURE) : input;
    }

    /**
     * Convenience method to check in one step whether a string is null or empty
     * @param input String to check
     * @return True if null or empty, otherwise false
     */
    public static boolean isNullOrEmpty(String input) {
        return input == null || input.isEmpty();
    }


}
