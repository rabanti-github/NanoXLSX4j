/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils;

import java.math.BigDecimal;
import java.time.Duration;
import java.util.Date;
import java.util.Locale;
import java.util.Optional;
import java.util.regex.Pattern;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.FormulaData;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

public class ParserUtils {
    private static final Pattern INVARIANT_DOUBLE_PATTERN = Pattern.compile(
            "[+-]?(?:(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d*)?|\\.\\d+)(?:[eE][+-]?\\d+)?");

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
     * Transforms a float to an invariant sting
     * <p>
     * This method is mainly a wrapper method but is kept to ensure a similar internal API between NanoXLSX and
     * NanoXLSX4j
     * </p>
     *
     * @param input Float to transform
     * @return Float as string
     */
    public static String toString(float input) {
        return Float.toString(input);
    }

    /**
     * Transforms a byte to an invariant sting
     * <p>
     * This method is mainly a wrapper method but is kept to ensure a similar internal API between NanoXLSX and
     * NanoXLSX4j
     * </p>
     *
     * @param input Byte to transform
     * @return Byte as string
     */
    public static String toString(byte input) {
        return Byte.toString(input);
    }

    /**
     * Transforms a double to an invariant sting
     * <p>
     * This method is mainly a wrapper method but is kept to ensure a similar internal API between NanoXLSX and
     * NanoXLSX4j
     * </p>
     *
     * @param input Double to transform
     * @return Double as string
     */
    public static String toString(double input) {
        return Double.toString(input);
    }

    /**
     * Transforms a BigDecimal to an invariant sting
     * <p>
     * This method is mainly a wrapper method but is kept to ensure a similar internal API between NanoXLSX and
     * NanoXLSX4j
     * </p>
     *
     * @param input BigDecimal to transform
     * @return BigDecimal as string
     */
    public static String toString(BigDecimal input) {
        return input.toPlainString();
    }

    /**
     * Transforms a long to an invariant sting
     * <p>
     * This method is mainly a wrapper method but is kept to ensure a similar internal API between NanoXLSX and
     * NanoXLSX4j
     * </p>
     *
     * @param input Long to transform
     * @return Long as string
     */
    public static String toString(long input) {
        return Long.toString(input);
    }

    /**
     * Transforms a short to an invariant sting
     * <p>
     * This method is mainly a wrapper method but is kept to ensure a similar internal API between NanoXLSX and
     * NanoXLSX4j
     * </p>
     *
     * @param input Short to transform
     * @return Short as string
     */
    public static String toString(short input) {
        return Short.toString(input);
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
     *
     * @param input String to check
     * @return True if null or empty, otherwise false
     */
    public static boolean isNullOrEmpty(String input) {
        return input == null || input.isEmpty();
    }

    /**
     * Convenience method to check in one step whether a string is null, empty or consists only of whitespaces
     *
     * @param input String to check
     * @return True if null, empty or whitespaces otherwise false
     */
    public static boolean isNullOrWhiteSpace(String input) {
        return input == null || input.isBlank();
    }

    /**
     * Null-safe method to compare the equality of two strings, ignoring the case
     *
     * @param a String a
     * @param b String b
     * @return True if equal, otherwise false
     */
    public static boolean equalsIgnoreCase(String a, String b) {
        return a == null ? b == null : a.equalsIgnoreCase(b);
    }

    /**
     * Tries to parse a boolean from its case-insensitive string representation.
     *
     * <p>Integer values such as {@code 0} and {@code 1} are not interpreted as booleans.</p>
     *
     * @param rawValue Raw boolean value
     * @return Parsed boolean, or an empty optional if parsing failed
     */
    public static Optional<Boolean> tryParseBoolean(String rawValue) {
        if (rawValue == null) {
            return Optional.empty();
        }
        String normalizedValue = rawValue.trim();
        if ("true".equalsIgnoreCase(normalizedValue)) {
            return Optional.of(true);
        }
        if ("false".equalsIgnoreCase(normalizedValue)) {
            return Optional.of(false);
        }
        return Optional.empty();
    }

    /**
     * Tries to parse an integer independently of the host locale.
     *
     * @param rawValue Raw integer value
     * @return Parsed integer, or an empty optional if parsing failed
     */
    public static Optional<Integer> tryParseInt(String rawValue) {
        if (rawValue == null) {
            return Optional.empty();
        }
        try {
            return Optional.of(Integer.parseInt(rawValue.trim()));
        } catch (NumberFormatException e) {
            return Optional.empty();
        }
    }

    /**
     * Tries to parse a double independently of the host locale.
     *
     * @param rawValue Raw double value
     * @return Parsed double, or an empty optional if parsing failed
     */
    public static Optional<Double> tryParseDouble(String rawValue) {
        if (rawValue == null) {
            return Optional.empty();
        }
        String normalizedValue = rawValue.trim();
        if (normalizedValue.endsWith("-") || normalizedValue.endsWith("+")) {
            normalizedValue = normalizedValue.substring(normalizedValue.length() - 1)
                    + normalizedValue.substring(0, normalizedValue.length() - 1);
        }
        if (!INVARIANT_DOUBLE_PATTERN.matcher(normalizedValue).matches()) {
            return Optional.empty();
        }
        try {
            return Optional.of(Double.parseDouble(normalizedValue.replace(",", "")));
        } catch (NumberFormatException e) {
            return Optional.empty();
        }
    }

    /**
     * Transforms a given object to a string displayed as cached Values. The common known compatible numeric types, like
     * int, float, sbyte etc. will be transformed to their appropriate string representations. Bool will be either 0 or
     * 1, or to TRUE or FALSE if convertBoolToNumber is set to false. Date or TimeSpan will be transformed to a OADate
     * (numeric) value. Null or an empty string will be transformed to 0. If an unknown object type is passed, its own
     * ToString() method will be used.
     *
     * <p>Remarks: This method transforms values to the Excel-internal OOXML format. It is not meant as a generic
     * ToString() method. Also do not pass nested objects like {@link Cell} or {@link FormulaData}, since they will be
     * handled as unknown object types.</p>
     *
     * @param input Object to transform
     * @return Most appropriate OOXML string form given
     * @throws FormatException Thrown if an invalid Date value was passed. See method
     *                         {@link DataUtils#getOADateTimeString(java.util.Date)} for details
     */
    public static String toCachedValueString(Object input) {
        return toCachedValueString(input, true);
    }

    /**
     * Transforms a given object to a string displayed as cached Values. The common known compatible numeric types, like
     * int, float, sbyte etc. will be transformed to their appropriate string representations. Bool will be either 0 or
     * 1, or to TRUE or FALSE if convertBoolToNumber is set to false. Date or TimeSpan will be transformed to a OADate
     * (numeric) value. Null or an empty string will be transformed to 0. If an unknown object type is passed, its own
     * ToString() method will be used.
     *
     * <p>Remarks: This method transforms values to the Excel-internal OOXML format. It is not meant as a generic
     * ToString() method. Also do not pass nested objects like {@link Cell} or {@link FormulaData}, since they will be
     * handled as unknown object types.</p>
     *
     * @param input               Object to transform
     * @param convertBoolToNumber If set to true, a bool value will be 1 or 0, otherwise TRUE or FALSE. Default is true
     * @return Most appropriate OOXML string form given
     * @throws FormatException Thrown if an invalid Date value was passed. See method
     *                         {@link DataUtils#getOADateTimeString(java.util.Date)} for details
     */
    public static String toCachedValueString(Object input, boolean convertBoolToNumber) {
        if (input == null) {
            return "0";
        } else if (input instanceof String) {
            String stringValue = (String) input;
            return isNullOrEmpty(stringValue) ? "0" : stringValue;
        } else if (input instanceof Boolean) {
            if (convertBoolToNumber) {
                return (boolean) input ? "1" : "0";
            } else {
                return (boolean) input ? "TRUE" : "FALSE";
            }
        } else if (input instanceof Byte) {
            return toString((byte) input);
        } else if (input instanceof BigDecimal) {
            return toString((BigDecimal) input);
        } else if (input instanceof Double) {
            return toString((double) input);
        } else if (input instanceof Float) {
            return toString((float) input);
        } else if (input instanceof Integer) {
            return toString((int) input);
        } else if (input instanceof Long) {
            return toString((long) input);
        } else if (input instanceof Short) {
            return toString((short) input);
        } else if (input instanceof Date) {
            return DataUtils.getOADateTimeString((Date) input);
        } else if (input instanceof Duration) {
            return DataUtils.getOATimeString((Duration) input);
        } else {
            return input.toString(); // Generic string
        }
    }

    /**
     * Tries to parse a raw string as an Excel formula string constant. Escaped double quotes ({@code ""}) are converted
     * to single double quotes ({@code "}).
     *
     * @param expression Raw string expression
     * @return Parsed and unescaped value, or an empty optional if the expression is invalid
     */
    public static Optional<String> tryParseFormulaStringConstant(String expression) {
        return tryParseFormulaStringConstant(expression, false);
    }

    /**
     * Tries to parse a raw string as an Excel formula string constant. Escaped double quotes ({@code ""}) are converted
     * to single double quotes ({@code "}).
     *
     * @param expression             Raw string expression
     * @param enclosingQuotesRemoved Whether the enclosing leading and trailing quotes were already removed
     * @return Parsed and unescaped value, or an empty optional if the expression is invalid
     */
    public static Optional<String> tryParseFormulaStringConstant(
            String expression, boolean enclosingQuotesRemoved) {
        if (expression == null) {
            return Optional.empty();
        }

        int startIndex;
        int endIndex;
        if (enclosingQuotesRemoved) {
            startIndex = 0;
            endIndex = expression.length();
        } else {
            if (expression.length() < 2
                    || expression.charAt(0) != '"'
                    || expression.charAt(expression.length() - 1) != '"') {
                return Optional.empty();
            }
            startIndex = 1;
            endIndex = expression.length() - 1;
        }

        StringBuilder builder = new StringBuilder(endIndex - startIndex);
        for (int i = startIndex; i < endIndex; i++) {
            char current = expression.charAt(i);
            if (current != '"') {
                builder.append(current);
                continue;
            }
            if (i + 1 >= endIndex || expression.charAt(i + 1) != '"') {
                return Optional.empty();
            }
            builder.append('"');
            i++;
        }
        return Optional.of(builder.toString());
    }

    /**
     * Tries to parse a worksheet-qualified address or range expression.
     *
     * @param expression Raw expression to parse
     * @return Parsed worksheet name and reference, or an empty optional if the expression is invalid
     */
    public static Optional<WorksheetQualifiedReference> tryParseWorksheetQualifiedReference(String expression) {
        if (isNullOrEmpty(expression)) {
            return Optional.empty();
        }

        if (expression.charAt(0) != '\'') {
            int separatorIndex = expression.indexOf('!');
            if (separatorIndex <= 0 || separatorIndex == expression.length() - 1) {
                return Optional.empty();
            }
            return Optional.of(new WorksheetQualifiedReference(
                    expression.substring(0, separatorIndex), expression.substring(separatorIndex + 1)));
        }

        StringBuilder builder = new StringBuilder();
        for (int i = 1; i < expression.length(); i++) {
            char current = expression.charAt(i);
            if (current != '\'') {
                builder.append(current);
                continue;
            }
            if (i + 1 < expression.length() && expression.charAt(i + 1) == '\'') {
                builder.append('\'');
                i++;
                continue;
            }
            if (i + 1 >= expression.length() || expression.charAt(i + 1) != '!') {
                return Optional.empty();
            }
            if (i + 2 >= expression.length()) {
                return Optional.empty();
            }
            return Optional.of(new WorksheetQualifiedReference(
                    builder.toString(), expression.substring(i + 2)));
        }
        return Optional.empty();
    }

    /** Parsed worksheet name and its attached address or range expression. */
    public record WorksheetQualifiedReference(String worksheetName, String reference) {
    }

    /**
     * Determines whether a formula contains an external workbook reference without parsing the formula. String
     * literals, structured references, and relative R1C1 references are ignored.
     *
     * <p>This method is primarily intended for internal NanoXLSX4j processing
     * and is not considered part of the stable public API.</p>
     *
     * @param formulaExpression Formula expression without a leading equal sign.
     * @return True if an external workbook reference was found.
     */
    public static boolean containsExternalReference(String formulaExpression) {
        if (isNullOrEmpty(formulaExpression) || formulaExpression.indexOf('[') < 0) {
            return false;
        }

        boolean inStringLiteral = false;
        for (int i = 0; i < formulaExpression.length(); i++) {
            char current = formulaExpression.charAt(i);
            if (current == '"') {
                if (inStringLiteral && i + 1 < formulaExpression.length() && formulaExpression.charAt(i + 1) == '"') {
                    i++;
                    continue;
                }
                inStringLiteral = !inStringLiteral;
                continue;
            }
            if (inStringLiteral || current != '[') {
                continue;
            }

            int closingBracket = formulaExpression.indexOf(']', i + 1);
            if (closingBracket <= i + 1) {
                continue;
            }

            boolean hasWorksheetName = false;
            for (int j = closingBracket + 1; j < formulaExpression.length(); j++) {
                char referenceCharacter = formulaExpression.charAt(j);
                if (referenceCharacter == '!') {
                    if (hasWorksheetName) {
                        return true;
                    }
                    break;
                }
                if (referenceCharacter == '[' || referenceCharacter == ']' || referenceCharacter == '"'
                        || referenceCharacter == '+' || referenceCharacter == '-' || referenceCharacter == '*'
                        || referenceCharacter == '/' || referenceCharacter == '^' || referenceCharacter == '&'
                        || referenceCharacter == '=' || referenceCharacter == '<' || referenceCharacter == '>'
                        || referenceCharacter == ',' || referenceCharacter == ';' || referenceCharacter == '('
                        || referenceCharacter == ')' || referenceCharacter == '{' || referenceCharacter == '}') {
                    break;
                }
                if (!Character.isWhitespace(referenceCharacter) && referenceCharacter != '\'') {
                    hasWorksheetName = true;
                }
            }
            i = closingBracket;
        }
        return false;
    }

}
