/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.utils;

import java.util.Locale;
import java.util.Optional;
import java.util.OptionalDouble;
import java.util.OptionalInt;
import java.util.OptionalLong;

/**
 * Class providing static methods to parse string values to specific types or to print objects as culture-neutral strings.
 * Methods in this class should only be used by the library components and not called by user code.
 *
 * @author Raphael Stoeckli
 */
public class ParserUtils {

    /**
     * Locale equivalent of CultureInfo.InvariantCulture. Ensures culture-neutral string conversions,
     * which is especially important for OOXML number values.
     */
    public static final Locale INVARIANT_LOCALE = Locale.ROOT;

    private ParserUtils() {
        // Prevents class instantiation
    }

    // ### S T R I N G U T I L I T I E S ###

    /**
     * Checks whether a string reference is null or the content is empty
     *
     * @param value value / reference to check
     * @return True if the passed parameter is null or empty, otherwise false
     */
    public static boolean isNullOrEmpty(String value) {
        if (value == null) {
            return true;
        }
        return value.isEmpty();
    }

    /**
     * Determines whether a string starts with a specific value (null-safe)
     *
     * @param input String to check
     * @param value Value to check whether it occurs at the beginning of the input string
     * @return True if the input string starts with the specified value
     */
    public static boolean startsWith(String input, String value) {
        if (input == null && value == null) {
            return true;
        }
        if (input == null || value == null) {
            return false;
        }
        return input.startsWith(value);
    }

    /**
     * Determines whether a string does not start with a specific value (null-safe)
     *
     * @param input String to check
     * @param value Value to check whether it does not occur at the beginning of the input string
     * @return True if the input string does not start with the specified value
     */
    public static boolean notStartsWith(String input, String value) {
        return !startsWith(input, value);
    }

    /**
     * Transforms a string to upper case with null check, using invariant locale
     *
     * @param input String to transform
     * @return Upper case string, or the original value if null or empty
     */
    public static String toUpper(String input) {
        return (input != null && !input.isEmpty()) ? input.toUpperCase(INVARIANT_LOCALE) : input;
    }

    /**
     * Transforms a string to lower case with null check, using invariant locale
     *
     * @param input String to transform
     * @return Lower case string, or the original value if null or empty
     */
    public static String toLower(String input) {
        return (input != null && !input.isEmpty()) ? input.toLowerCase(INVARIANT_LOCALE) : input;
    }

    /**
     * Normalizes all newlines of a string to CR+LF
     *
     * @param value Input value
     * @return Normalized value with CR+LF line endings
     */
    public static String normalizeNewLines(String value) {
        if (value == null || (!value.contains("\n") && !value.contains("\r"))) {
            return value;
        }
        return value.replace("\n\r", "\n").replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n");
    }

    // ### T O S T R I N G O V E R L O A D S ###

    /**
     * Transforms an int to a culture-neutral string
     *
     * @param input Value to transform
     * @return Value as string
     */
    public static String toString(int input) {
        return Integer.toString(input);
    }

    /**
     * Transforms a long to a culture-neutral string
     *
     * @param input Value to transform
     * @return Value as string
     */
    public static String toString(long input) {
        return Long.toString(input);
    }

    /**
     * Transforms a float to a culture-neutral string
     *
     * @param input Value to transform
     * @return Value as string
     */
    public static String toString(float input) {
        return Float.toString(input);
    }

    /**
     * Transforms a double to a culture-neutral string
     *
     * @param input Value to transform
     * @return Value as string
     */
    public static String toString(double input) {
        return Double.toString(input);
    }

    /**
     * Transforms a byte to a culture-neutral string
     *
     * @param input Value to transform
     * @return Value as string
     */
    public static String toString(byte input) {
        return Byte.toString(input);
    }

    /**
     * Transforms a short to a culture-neutral string
     *
     * @param input Value to transform
     * @return Value as string
     */
    public static String toString(short input) {
        return Short.toString(input);
    }

    // ### P A R S E M E T H O D S ###

    /**
     * Parses a float independent from the culture info of the host
     *
     * @param rawValue Raw number as string
     * @return Parsed float
     * @apiNote The method does not check the validity and will throw if an invalid value is passed
     */
    public static float parseFloat(String rawValue) {
        return Float.parseFloat(rawValue);
    }

    /**
     * Parses an int independent from the culture info of the host
     *
     * @param rawValue Raw number as string
     * @return Parsed int
     * @apiNote The method does not check the validity and will throw if an invalid value is passed
     */
    public static int parseInt(String rawValue) {
        return Integer.parseInt(rawValue);
    }

    /**
     * Parses a double independent from the culture info of the host
     *
     * @param rawValue Raw number as string
     * @return Parsed double
     * @apiNote The method does not check the validity and will throw if an invalid value is passed
     */
    public static double parseDouble(String rawValue) {
        return Double.parseDouble(rawValue);
    }

    /**
     * Parses a bool as a binary number either based on an int (0/1) or a string expression (true/false)
     *
     * @param rawValue Raw number or expression as string
     * @return Parsed bool as number (0 = false, 1 = true)
     */
    public static int parseBinaryBool(String rawValue) {
        if (isNullOrEmpty(rawValue)) {
            return 0;
        }
        OptionalInt optInt = tryParseInt(rawValue);
        if (optInt.isPresent()) {
            return optInt.getAsInt() >= 1 ? 1 : 0;
        }
        return rawValue.toLowerCase(INVARIANT_LOCALE).equals("true") ? 1 : 0;
    }

    // ### T R Y P A R S E M E T H O D S ###

    /**
     * Tries to parse an int independent from the culture info of the host
     *
     * @param rawValue Raw number as string
     * @return OptionalInt with the parsed value, or empty if parsing failed
     */
    public static OptionalInt tryParseInt(String rawValue) {
        try {
            return OptionalInt.of(Integer.parseInt(rawValue));
        }
        catch (NumberFormatException e) {
            return OptionalInt.empty();
        }
    }

    /**
     * Tries to parse a long independent from the culture info of the host
     *
     * @param rawValue Raw number as string
     * @return OptionalLong with the parsed value, or empty if parsing failed
     */
    public static OptionalLong tryParseLong(String rawValue) {
        try {
            return OptionalLong.of(Long.parseLong(rawValue));
        }
        catch (NumberFormatException e) {
            return OptionalLong.empty();
        }
    }

    /**
     * Tries to parse a float independent from the culture info of the host
     *
     * @param rawValue Raw number as string
     * @return Optional with the parsed value, or empty if parsing failed
     */
    public static Optional<Float> tryParseFloat(String rawValue) {
        try {
            return Optional.of(Float.parseFloat(rawValue));
        }
        catch (NumberFormatException e) {
            return Optional.empty();
        }
    }

    /**
     * Tries to parse a double independent from the culture info of the host
     *
     * @param rawValue Raw number as string
     * @return OptionalDouble with the parsed value, or empty if parsing failed
     */
    public static OptionalDouble tryParseDouble(String rawValue) {
        try {
            return OptionalDouble.of(Double.parseDouble(rawValue));
        }
        catch (NumberFormatException e) {
            return OptionalDouble.empty();
        }
    }

}
