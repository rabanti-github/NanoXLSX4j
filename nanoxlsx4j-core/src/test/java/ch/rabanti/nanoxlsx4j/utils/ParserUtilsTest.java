/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils;

import ch.rabanti.nanoxlsx4j.FormulaData;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.MethodSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;

import java.math.BigDecimal;
import java.time.Duration;
import java.time.LocalDate;
import java.time.ZoneOffset;
import java.util.Date;
import java.util.Optional;
import java.util.UUID;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ParserUtilsTest {

    @DisplayName("Test of the ParserUtils toUpper function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "''|''",
                    "NULL|NULL",
                    "123|123",
                    "abc|ABC",
                    "ABC|ABC"
            }
    )
    void toUpperTest(String givenValue, String expectedValue) {
        String value = ParserUtils.toUpper(givenValue);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the ParserUtils toLower function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "''|''",
                    "NULL|NULL",
                    "123|123",
                    "abc|abc",
                    "ABC|abc"
            }
    )
    void toLowerTest(String givenValue, String expectedValue) {
        String value = ParserUtils.toLower(givenValue);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the ParserUtils startsWith function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "HelloWorld|Hello|true",
                    "HelloWorld|world|false",
                    "000|0|true",
                    "''|''|true",
                    "NULL|NULL|true",
                    "NULL|test|false",
                    "test|NULL|false",
                    "012|3|false",
                    "abc|abc|true",
                    "abc|ABC|false",
                    "'   '|' '|true",
                    "'   '|\t|false"
            }
    )
    void startsWithTest(String givenValue, String startValue, boolean expectedStartsWith) {
        boolean startsWith = ParserUtils.startsWith(givenValue, startValue);
        assertEquals(expectedStartsWith, startsWith);
    }

    @DisplayName("Test of the ParserUtils NotStartsWith function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "HelloWorld|Hello|false",
                    "HelloWorld|world|true",
                    "000|0|false",
                    "''|''|false",
                    "NULL|NULL|false",
                    "NULL|test|true",
                    "test|NULL|true",
                    "012|3|true",
                    "abc|abc|false",
                    "abc|ABC|true",
                    "'   '|' '|false",
                    "'   '|\t|true"
            }
    )
    void notStartsWithTest(String givenValue, String startValue, boolean expectedStartsWith) {
        boolean startsWith = ParserUtils.NotStartsWith(givenValue, startValue);
        assertEquals(expectedStartsWith, startsWith);
    }

    @DisplayName("Test of the ParserUtils toString function for integers")
    @ParameterizedTest
    @CsvSource(
            {
                    "-10, -10",
                    "0, 0",
                    "1, 1",
                    "100, 100"}
    )
    void toStringTest(int givenValue, String expectedValue) {
        String value = ParserUtils.toString(givenValue);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the ParserUtils toString function for floats")
    @ParameterizedTest
    @CsvSource(
            {
                    "-10, -10",
                    "0, 0",
                    "1, 1",
                    "100, 100",
                    "0.1, 0.1",
                    "-0.01, -0.01",
                    "100.01, 100.01",
                    "-1.111, -1.111",
                    "Infinity, Infinity",
                    "-Infinity, -Infinity",
                    "NaN, NaN"
            }
    )
    void toStringTest2(float givenValue, String expectedValue) {
        String value = ParserUtils.toString(givenValue);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the ParserUtils toString function for doubles")
    @ParameterizedTest
    @CsvSource(
            {
                    "100.0, 100",
                    "Infinity, Infinity",
                    "-Infinity, -Infinity",
                    "NaN, NaN"
            }
    )
    void toStringDoubleTest(double givenValue, String expectedValue) {
        String value = ParserUtils.toString(givenValue);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the ParserUtils toCachedValueString function for null and strings")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "NULL|0",
                    "''|0",
                    "' '|' '",
                    "text|text",
                    "Grüße 世界|Grüße 世界"
            }
    )
    void toCachedValueStringTest(Object givenValue, String expectedValue) {
        String value = ParserUtils.toCachedValueString(givenValue);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the ParserUtils toCachedValueString function for booleans")
    @ParameterizedTest
    @CsvSource(
            {
                    "true, true, 1",
                    "false, true, 0",
                    "true, false, TRUE",
                    "false, false, FALSE"}
    )
    void toCachedValueStringBoolTest(
            boolean givenValue, boolean givenConvertBoolToNumber, String expectedValue) {
        String value = ParserUtils.toCachedValueString(givenValue, givenConvertBoolToNumber);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the ParserUtils toCachedValueString function for numerical types")
    @ParameterizedTest
    @MethodSource("cachedNumericArguments") // see helper method cachedNumericArguments()
    void toCachedValueStringNumericTest(Object givenValue, String expectedValue) {
        String value = ParserUtils.toCachedValueString(givenValue);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the ParserUtils toCachedValueString function for decimals")
    @Test
    void toCachedValueStringDecimalTest() {
        String value = ParserUtils.toCachedValueString(new BigDecimal("-1234.5678"));
        assertEquals("-1234.5678", value);
    }

    @DisplayName("Test of the ParserUtils toCachedValueString function for dates")
    @Test
    void toCachedValueStringDateTimeTest() {
        Date givenValue = Date.from(LocalDate.of(2021, 1, 1).atStartOfDay(ZoneOffset.UTC).toInstant());
        String value = ParserUtils.toCachedValueString(givenValue);
        assertEquals("44197", value);
    }

    @DisplayName("Test of the ParserUtils toCachedValueString function for durations")
    @Test
    void toCachedValueStringDurationTest() {
        String value = ParserUtils.toCachedValueString(Duration.ofHours(36));
        assertEquals("1.5", value);
    }

    @DisplayName("Test of the ParserUtils toCachedValueString function for unknown object types")
    @Test
    void toCachedValueStringUnknownObjectTest() {
        UUID givenValue = UUID.fromString("12345678-1234-5678-90ab-1234567890ab");
        String value = ParserUtils.toCachedValueString(givenValue);
        assertEquals("12345678-1234-5678-90ab-1234567890ab", value);
    }

    @DisplayName("Test of the failing ParserUtils toCachedValueString function for invalid dates")
    @Test
    void toCachedValueStringDateTimeFailTest() {
        assertThrows(FormatException.class, () -> ParserUtils.toCachedValueString(new Date(Long.MIN_VALUE)));
    }

    @DisplayName("Test of the failing ParserUtils toCachedValueString function for unknown object types")
    @Test
    void toCachedValueStringUnknownObjectFailTest() {
        assertThrows(
                IllegalStateException.class,
                () -> ParserUtils.toCachedValueString(new InvalidStringValue())
        );
    }

    @DisplayName("Test of the ParserUtils normalizeNewLines function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "NULL|NULL",
                    "''|''",
                    "test|test",
                    "'test\r\ntest'|'test\r\ntest'",
                    "'test\rtest'|'test\r\ntest'",
                    "'test\ntest'|'test\r\ntest'",
                    "'test\n\rtest'|'test\r\ntest'",
                    "'test\r\ntest \r\ntest'|'test\r\ntest \r\ntest'",
                    "'test\rtest \rtest'|'test\r\ntest \r\ntest'",
                    "'test\ntest \ntest'|'test\r\ntest \r\ntest'",
                    "'test\n\rtest \n\rtest'|'test\r\ntest \r\ntest'",
                    "'\n\r\n\n'|'\r\n\r\n\r\n'"
            }
    )
    void normalizeNewLinesTest(String givenValue, String expectedValue) {
        String value = ParserUtils.normalizeNewLines(givenValue);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the ParserUtils isAsciiDigit function")
    @ParameterizedTest
    @ValueSource(
            chars = {
                    '0',
                    '1',
                    '5',
                    '9'}
    )
    void isAsciiDigitTest(char givenCharacter) {
        boolean match = ParserUtils.isAsciiDigit(givenCharacter);
        assertTrue(match);
    }

    @DisplayName("Test of the ParserUtils isAsciiDigit function on invalid values")
    @ParameterizedTest
    @ValueSource(
            chars = {
                    '/',
                    ':',
                    'a',
                    '\0',
                    '\t',
                    '\n',
                    '\u0661',
                    '\u06F1',
                    '\u0967',
                    '\uFF11',
                    '\u00B2',
                    '\u2460'
            }
    )
    void isAsciiDigitFailTest(char givenCharacter) {
        boolean match = ParserUtils.isAsciiDigit(givenCharacter);
        assertFalse(match);
    }

    @DisplayName("Test of the ParserUtils isNullOrWhiteSpace function")
    @ParameterizedTest
    @NullSource
    @ValueSource(strings = {"", " ", "\t", "\n"})
    void isNullOrWhiteSpaceTest(String givenValue) {
        assertTrue(ParserUtils.isNullOrWhiteSpace(givenValue));
    }

    @DisplayName("Test of the ParserUtils isNullOrWhiteSpace function")
    @ParameterizedTest
    @ValueSource(strings = {"a", " a "})
    void isNullOrWhiteSpaceFailTest(String givenValue) {
        assertFalse(ParserUtils.isNullOrWhiteSpace(givenValue));
    }

    @DisplayName("Test of the ParserUtils equalsIgnoreCase function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "NULL|NULL|true",
                    "NULL|text|false",
                    "text|NULL|false",
                    "''|''|true",
                    "text|text|true",
                    "text|TEXT|true",
                    "text|other|false"
            }
    )
    void equalsIgnoreCaseTest(String a, String b, boolean expectedValue) {
        assertEquals(expectedValue, ParserUtils.equalsIgnoreCase(a, b));
    }

    @DisplayName("Test of the ParserUtils parseFloat function (no error handling)")
    @ParameterizedTest
    @CsvSource(
            {
                    "1, 1",
                    "0, 0",
                    "-1, -1",
                    "-10, -10",
                    "22, 22",
                    "-0.005, -0.005",
                    "0.858, 0.858",
                    "-99998.1234, -99998.1234",
                    "98755142.237, 98755142.237"
            }
    )
    void parseFloatTest(String givenValue, float expectedValue) {
        float value = ParserUtils.parseFloat(givenValue);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the ParserUtils parseInt function (no error handling)")
    @ParameterizedTest
    @CsvSource(
            {
                    "0, 0",
                    "1, 1",
                    "1.0, 1",
                    "-2.0, -2",
                    "0.0, 0",
                    "-1, -1",
                    "42, 42",
                    "-42, -42",
                    "2147483647, 2147483647",
                    "-2147483648, -2147483648"
            }
    )
    void parseIntTest(String givenValue, int expectedValue) {
        int value = ParserUtils.parseInt(givenValue);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the failing ParserUtils parseInt function")
    @ParameterizedTest
    @ValueSource(strings = {"a", "1x1", "--1"})
    void parseIntFailTest(String givenValue) {
        assertThrows(NumberFormatException.class, () -> ParserUtils.parseInt(givenValue));
    }

    @DisplayName("Test of the ParserUtils parseDouble function (no error handling)")
    @ParameterizedTest
    @CsvSource(
            {
                    "0, 0",
                    "1, 1",
                    "1.0, 1",
                    "-2.0, -2",
                    "0.0, 0",
                    "-1, -1",
                    "42, 42",
                    "-42, -42",
                    "-0.0001, -0.0001",
                    "15.258789, 15.258789",
                    "1.7976931348623157E+308, 1.7976931348623157E+308",
                    "-1.7976931348623157E+308, -1.7976931348623157E+308"
            }
    )
    void parseDoubleTest(String givenValue, double expectedValue) {
        double value = ParserUtils.parseDouble(givenValue);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the ParserUtils parseBinaryBoolean function (no error handling)")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "0|0",
                    "1|1",
                    "-1|0",
                    "2|1",
                    "false|0",
                    "FALSE|0",
                    "False|0",
                    "''|0",
                    "NULL|0",
                    "no|0",
                    "true|1",
                    "TRUE|1",
                    "True|1"
            }
    )
    void parseBinaryBooleanTest(String givenValue, int expectedValue) {
        int value = ParserUtils.parseBinaryBoolean(givenValue);
        assertEquals(expectedValue, value);
    }

    @DisplayName("Test of the ParserUtils tryParseBoolean function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "true|true",
                    "TRUE|true",
                    "false|false",
                    "' FALSE '|false",
                    "NULL|NULL",
                    "''|NULL",
                    "1|NULL",
                    "yes|NULL"
            }
    )
    void tryParseBooleanTest(String givenValue, Boolean expectedValue) {
        Optional<Boolean> value = ParserUtils.tryParseBoolean(givenValue);
        assertEquals(Optional.ofNullable(expectedValue), value);
    }

    @DisplayName("Test of the ParserUtils tryParseFloat function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "1|1|true",
                    "0|0|true",
                    "-1|-1|true",
                    "-10|-10|true",
                    "22|22|true",
                    "-0.005|-0.005|true",
                    "0.858|0.858|true",
                    "-99998.1234|-99998.1234|true",
                    "98755142.237|98755142.237|true",
                    "''|0|false",
                    "' '|0|false",
                    "NULL|0|false",
                    "a|0|false",
                    "1x1|0|false",
                    "0.0x|0|false",
                    "-22.5f4|0|false"
            }
    )
    void tryParseFloatTest(String givenValue, float expectedValue, boolean expectedMatch) {
        Optional<Float> value = ParserUtils.tryParseFloat(givenValue);
        assertEquals(expectedMatch, value.isPresent());
        value.ifPresent(parsedValue -> assertEquals(expectedValue, parsedValue));
    }

    @DisplayName("Test of the ParserUtils tryParseInt function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "0|0|true",
                    "1|1|true",
                    "-1|-1|true",
                    "42|42|true",
                    "-42|-42|true",
                    "2147483647|2147483647|true",
                    "''|0|false",
                    "' '|0|false",
                    "NULL|0|false",
                    "a|0|false",
                    "1x1|0|false"
            }
    )
    void tryParseIntTest(String givenValue, int expectedValue, boolean expectedMatch) {
        Optional<Integer> value = ParserUtils.tryParseInt(givenValue);
        assertEquals(expectedMatch, value.isPresent());
        value.ifPresent(parsedValue -> assertEquals(expectedValue, parsedValue));
    }

    // TryParseUintTest has no Java counterpart; Java does not expose an unsigned-int parser.

    @DisplayName("Test of the ParserUtils tryParseLong function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "0|0|true",
                    "1|1|true",
                    "-1|-1|true",
                    "42|42|true",
                    "-42|-42|true",
                    "9223372036854775807|9223372036854775807|true",
                    "''|0|false",
                    "' '|0|false",
                    "NULL|0|false",
                    "a|0|false",
                    "1x1|0|false"
            }
    )
    void tryParseLongTest(String givenValue, long expectedValue, boolean expectedMatch) {
        Optional<Long> value = ParserUtils.tryParseLong(givenValue);
        assertEquals(expectedMatch, value.isPresent());
        value.ifPresent(parsedValue -> assertEquals(expectedValue, parsedValue));
    }

    // TryParseUlongTest has no Java counterpart; Java does not expose an unsigned-long parser.

    @DisplayName("Test of the ParserUtils tryParseBigDecimal function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "1|1|true",
                    "0|0|true",
                    "-1|-1|true",
                    "-10|-10|true",
                    "22|22|true",
                    "-0.0000005|-0.0000005|true",
                    "0.858|0.858|true",
                    "-99998.1234|-99998.1234|true",
                    "98755142.2111137|98755142.2111137|true",
                    "''|0|false",
                    "' '|0|false",
                    "NULL|0|false",
                    "a|0|false",
                    "1x1|0|false",
                    "0.0x|0|false",
                    "-22.5f4|0|false"
            }
    )
    void tryParseBigDecimalTest(String givenValue, String expectedValue, boolean expectedMatch) {
        Optional<BigDecimal> value = ParserUtils.tryParseBigDecimal(givenValue);
        assertEquals(expectedMatch, value.isPresent());
        value.ifPresent(parsedValue -> assertEquals(new BigDecimal(expectedValue), parsedValue));
    }

    @DisplayName("Test of the ParserUtils tryParseDouble function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "1|1|true",
                    "0|0|true",
                    "-1|-1|true",
                    "-10|-10|true",
                    "22|22|true",
                    "-0.0000005|-0.0000005|true",
                    "0.858|0.858|true",
                    "-99998.1234|-99998.1234|true",
                    "98755142.2111137|98755142.2111137|true",
                    "''|0|false",
                    "' '|0|false",
                    "NULL|0|false",
                    "a|0|false",
                    "1x1|0|false",
                    "0.0x|0|false",
                    "-22.5f4|0|false"
            }
    )
    void tryParseDoubleTest(String givenValue, double expectedValue, boolean expectedMatch) {
        Optional<Double> value = ParserUtils.tryParseDouble(givenValue);
        assertEquals(expectedMatch, value.isPresent());
        value.ifPresent(parsedValue -> assertEquals(expectedValue, parsedValue));
    }

    @DisplayName("Test of the ParserUtils tryParseDouble function with defined number styles")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            value = {
                    "1,234 | 1234",
                    "123- | -123",
                    " 1.25E+3  | 1250"
            }
    )
    void tryParseDoubleTest2(String givenValue, double expectedValue) {
        Optional<Double> value = ParserUtils.tryParseDouble(givenValue);
        assertEquals(Optional.of(expectedValue), value);
    }

    // NumberStyles.Float-specific rows have no Java counterpart; the Java API exposes only invariant Any behavior.

    @DisplayName("Test of the ParserUtils tryParseFormulaStringConstant function without options")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            value = {
                    "'\"\"'|''",
                    "'\"text\"'|text",
                    "'\"He said \"\"Hello\"\"\"'|'He said \"Hello\"'"
            }
    )
    void tryParseFormulaStringConstantWithoutOptionsTest(String givenExpression, String expectedValue) {
        Optional<String> value = ParserUtils.tryParseFormulaStringConstant(givenExpression);
        assertEquals(Optional.of(expectedValue), value);
    }

    @DisplayName("Test of the successful ParserUtils tryParseFormulaStringConstant function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            value = {
                    "'\"\"'|''|false",
                    "'\"text\"'|text|false",
                    "'\"text with spaces\"'|'text with spaces'|false",
                    "'\"Grüße 世界\"'|'Grüße 世界'|false",
                    "'\"line 1\nline 2\"'|'line 1\nline 2'|false",
                    "'\"He said \"\"Hello\"\"\"'|'He said \"Hello\"'|false",
                    "'\"\"\"\"'|'\"'|false",
                    "''|''|true",
                    "text|text|true",
                    "日本語|日本語|true",
                    "'He said \"\"Hello\"\"'|'He said \"Hello\"'|true",
                    "'\"\"'|'\"'|true"
            }
    )
    void tryParseFormulaStringConstantTest(
            String givenExpression, String expectedValue, boolean givenEnclosingQuotesRemoved) {
        Optional<String> value = ParserUtils.tryParseFormulaStringConstant(
                givenExpression, givenEnclosingQuotesRemoved);
        assertTrue(value.isPresent());
        assertEquals(expectedValue, value.orElseThrow());
    }

    @DisplayName("Test of the failing ParserUtils tryParseFormulaStringConstant function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "NULL|false",
                    "''|false",
                    "'\"'|false",
                    "text|false",
                    "'\"text'|false",
                    "'text\"'|false",
                    "'\"unpaired \" quote\"'|false",
                    "'\"'|true",
                    "'unpaired \" quote'|true"
            }
    )
    void tryParseFormulaStringConstantFailTest(String givenExpression, boolean givenEnclosingQuotesRemoved) {
        Optional<String> value = ParserUtils.tryParseFormulaStringConstant(
                givenExpression, givenEnclosingQuotesRemoved);
        assertTrue(value.isEmpty());
    }

    @DisplayName("Test of the successful ParserUtils tryParseWorksheetQualifiedReference function")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            quoteCharacter = '"',
            value = {
                    "Sheet1!A1|Sheet1|A1",
                    "Tabelle Übersicht!$XFD$1048576|Tabelle Übersicht|$XFD$1048576",
                    "工作表!A1:B2|工作表|A1:B2",
                    "Sheet1!A1!B2|Sheet1|A1!B2",
                    "'Sheet 1'!A1|Sheet 1|A1",
                    "'Übersicht 世界'!$A$1:$B$2|Übersicht 世界|$A$1:$B$2",
                    "'Owner''s Sheet'!C3|Owner's Sheet|C3",
                    "''''!A1|'|A1",
                    "''!A1|\"\"|A1"
            }
    )
    void tryParseWorksheetQualifiedReferenceTest(
            String givenExpression, String expectedWorksheetName, String expectedReference) {
        Optional<ParserUtils.WorksheetQualifiedReference> value =
                ParserUtils.tryParseWorksheetQualifiedReference(givenExpression);
        assertTrue(value.isPresent());
        assertEquals(expectedWorksheetName, value.orElseThrow().worksheetName());
        assertEquals(expectedReference, value.orElseThrow().reference());
    }

    @DisplayName("Test of the failing ParserUtils tryParseWorksheetQualifiedReference function")
    @ParameterizedTest
    @NullSource
    @ValueSource(
            strings = {
                    "",
                    "Sheet1",
                    "!A1",
                    "Sheet1!",
                    "'",
                    "'Sheet 1",
                    "'Sheet 1'",
                    "'Sheet 1' A1",
                    "'Sheet 1'!",
                    "'Owner''s Sheet"
            }
    )
    void tryParseWorksheetQualifiedReferenceFailTest(String givenExpression) {
        Optional<ParserUtils.WorksheetQualifiedReference> value =
                ParserUtils.tryParseWorksheetQualifiedReference(givenExpression);
        assertTrue(value.isEmpty());
    }

    @DisplayName("Test of several numerical parse and tryParse functions for their minimum values")
    @Test
    void parseMinTest() {
        Optional<BigDecimal> dValue = ParserUtils.tryParseBigDecimal("-79228162514264337593543950335");
        assertEquals(new BigDecimal("-79228162514264337593543950335"), dValue.orElseThrow());

        Optional<Long> lValue = ParserUtils.tryParseLong("-9223372036854775808");
        assertEquals(Long.MIN_VALUE, lValue.orElseThrow());

        Optional<Integer> iValue = ParserUtils.tryParseInt("-2147483648");
        assertEquals(Integer.MIN_VALUE, iValue.orElseThrow());

        assertEquals(Integer.MIN_VALUE, ParserUtils.parseInt("-2147483648"));

        Optional<Float> fValue = ParserUtils.tryParseFloat("-3.40282347E+38");
        assertEquals(-Float.MAX_VALUE, fValue.orElseThrow());

        assertEquals(-Float.MAX_VALUE, ParserUtils.parseFloat("-3.40282347E+38"));

        Optional<Double> dbValue = ParserUtils.tryParseDouble("-1.7976931348623157E+308");
        assertEquals(-Double.MAX_VALUE, dbValue.orElseThrow());
    }

    @DisplayName("Test of several numerical parse and tryParse functions for their maximum values")
    @Test
    void parseMaxTest() {
        Optional<BigDecimal> dValue = ParserUtils.tryParseBigDecimal("79228162514264337593543950335");
        assertEquals(new BigDecimal("79228162514264337593543950335"), dValue.orElseThrow());

        Optional<Long> lValue = ParserUtils.tryParseLong("9223372036854775807");
        assertEquals(Long.MAX_VALUE, lValue.orElseThrow());

        Optional<Integer> iValue = ParserUtils.tryParseInt("2147483647");
        assertEquals(Integer.MAX_VALUE, iValue.orElseThrow());

        assertEquals(Integer.MAX_VALUE, ParserUtils.parseInt("2147483647"));

        Optional<Float> fValue = ParserUtils.tryParseFloat("3.40282347E+38");
        assertEquals(Float.MAX_VALUE, fValue.orElseThrow());

        assertEquals(Float.MAX_VALUE, ParserUtils.parseFloat("3.40282347E+38"));

        Optional<Double> dbValue = ParserUtils.tryParseDouble("1.7976931348623157E+308");
        assertEquals(Double.MAX_VALUE, dbValue.orElseThrow());
    }

    @DisplayName("Test of valid and invalid external link identifiers")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "[0]|true",
                    "[1]|true",
                    "[001]|true",
                    "[1234567890]|true",
                    "NULL|false",
                    "''|false",
                    "[]|false",
                    "[a]|false",
                    "[1a]|false",
                    "[ 1]|false",
                    "[1 ]|false",
                    "[\0]|false",
                    "[\t]|false",
                    "[\u0661]|false",
                    "1]|false",
                    "[1|false",
                    "prefix[1]|false",
                    "[1]suffix|false"
            }
    )
    void isValidExternalLinkIdTest(String givenIdentifier, boolean expectedMatch) {
        boolean match = ParserUtils.isValidExternalLinkId(givenIdentifier);
        assertEquals(expectedMatch, match);
    }

    @DisplayName("Test of successfully reading external link identifiers")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            quoteCharacter = '"',
            ignoreLeadingAndTrailingWhitespace = false,
            value = {
                    "[0]Sheet1!A1|0|[0]",
                    "SUM([12]Sheet_Name!$A$1)|4|[12]",
                    "'[003]Sheet name'!A1|1|[003]",
                    "+[4]!ExternalName|1|[4]",
                    "=[5]#REF!A1|1|[5]",
                    " [6]工作表!A1|1|[6]"
            }
    )
    void tryReadExternalLinkIdTest(String givenExpression, int givenStartIndex, String expectedIdentifier) {
        Optional<Integer> identifierLength = ParserUtils.tryReadExternalLinkId(givenExpression, givenStartIndex);
        assertTrue(identifierLength.isPresent());
        assertEquals(expectedIdentifier.length(), identifierLength.orElseThrow());
        assertEquals(
                expectedIdentifier,
                givenExpression.substring(givenStartIndex, givenStartIndex + identifierLength.orElseThrow())
        );
    }

    @DisplayName("Test of failing to read malformed or incorrectly bounded external link identifiers")
    @ParameterizedTest
    @CsvSource(
            delimiter = '|',
            nullValues = "NULL",
            value = {
                    "NULL|0",
                    "''|0",
                    "[1]Sheet|-1",
                    "[1]Sheet|8",
                    "A1|0",
                    "[|0",
                    "[]Sheet|0",
                    "[a]Sheet|0",
                    "[\u0661]Sheet|0",
                    "[12|0",
                    "[12Sheet|0",
                    "[12a]Sheet|0",
                    "[1]|0",
                    "[1] Sheet|0",
                    "[1]\tSheet|0",
                    "Table[1]Column|5",
                    "1[1]Column|1",
                    "_[1]Column|1",
                    "\\[1]Column|1",
                    ".[1]Column|1",
                    "名[1]Column|1",
                    "[1]\"Sheet|0",
                    "[1][Sheet|0",
                    "[1]]Sheet|0",
                    "[1](Sheet|0",
                    "[1])Sheet|0",
                    "[1],Sheet|0",
                    "[1];Sheet|0",
                    "[1]+Sheet|0",
                    "[1]-Sheet|0",
                    "[1]*Sheet|0",
                    "[1]/Sheet|0",
                    "[1]^Sheet|0",
                    "[1]&Sheet|0",
                    "[1]=Sheet|0",
                    "[1]<Sheet|0",
                    "[1]>Sheet|0",
                    "[1]%Sheet|0",
                    "[1]:Sheet|0"
            }
    )
    void tryReadExternalLinkIdFailTest(String givenExpression, int givenStartIndex) {
        Optional<Integer> identifierLength = ParserUtils.tryReadExternalLinkId(givenExpression, givenStartIndex);
        assertTrue(identifierLength.isEmpty());
    }

    @DisplayName("Test of external workbook reference detection in formulas")
    @ParameterizedTest
    @ValueSource(
            strings = {
                    "[1]Sheet1!A1",
                    "SUM([12]Sheet_Name!$A$1)",
                    "'[1]Sheet 1'!$A$1",
                    "[Book.xlsx]Sheet1!A1",
                    "SUM('[Book.xlsx]Owner''s Sheet'!A1)",
                    "[1]Sheet1!A1+[2]Sheet2!B2",
                    "'..\\[Book.xlsx]Sheet1'!$A$1",
                    "'../[Book.xlsx]Sheet1'!$A$1",
                    "'C:\\temp\\[book one.xlsx]Sheet 1'!$A$1",
                    "SUM('C:\\temp\\[book one.xlsx]Sheet 1'!$A$1,'..\\[other.xlsx]Data'!$B$2)"
            }
    )
    void containsExternalReferenceTest(String expression) {
        FormulaData data = new FormulaData(expression);
        assertTrue(data.hasExternalReferences());
        assertTrue(ParserUtils.containsExternalReference(expression));
    }

    @DisplayName("Test of expressions without external workbook references")
    @ParameterizedTest
    @NullSource
    @ValueSource(
            strings = {
                    "",
                    "SUM(A1:A2)",
                    "Table1[Column]",
                    "Table1[1]",
                    "R[1]C[1]",
                    "[",
                    "[]Sheet1!A1",
                    "[1]",
                    "[1]!A1",
                    "[1]Sheet1+A1",
                    "Table1[Column]+Sheet1!A1",
                    "\"[1]Sheet1!A1\"",
                    "INDIRECT(\"[1]Sheet1!A1\")",
                    "\"escaped \"\"[1]Sheet1!A1\"\" text\""
            }
    )
    void containsExternalReferenceNegativeTest(String expression) {
        FormulaData data = new FormulaData(expression);
        assertFalse(data.hasExternalReferences());
        assertFalse(ParserUtils.containsExternalReference(expression));
    }

    private static Stream<Arguments> cachedNumericArguments() {
        return Stream.of(
                Arguments.of((short) 255, "255"), Arguments.of((byte) -128, "-128"),
                Arguments.of((short) -32768, "-32768"), Arguments.of(65535, "65535"),
                Arguments.of(Integer.MIN_VALUE, "-2147483648"),
                Arguments.of(4294967295L, "4294967295"), Arguments.of(Long.MIN_VALUE, "-9223372036854775808"),
                Arguments.of(new BigDecimal("18446744073709551615"), "18446744073709551615"),
                Arguments.of(-1.25f, "-1.25"),
                Arguments.of(-1234.5d, "-1234.5")
        );
    }

    private static class InvalidStringValue {
        @Override
        public String toString() {
            throw new IllegalStateException("The value cannot be converted to a string");
        }
    }
}
