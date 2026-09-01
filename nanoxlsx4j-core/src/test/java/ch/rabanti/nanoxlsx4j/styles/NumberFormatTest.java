/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Optional;
import java.util.stream.Stream;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.MethodSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;

class NumberFormatTest {

    private final NumberFormat exampleStyle;

    NumberFormatTest() {
        exampleStyle = new NumberFormat();
        exampleStyle.setCustomFormatCode("#.###");
        exampleStyle.setNumber(NumberFormat.FormatNumber.FORMAT_10);
        exampleStyle.setCustomFormatId(170);
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the FormatNumber property")
    @EnumSource(NumberFormat.FormatNumber.class) // adds all enum values
    void formatNumberTest(NumberFormat.FormatNumber number) {
        NumberFormat numberFormat = new NumberFormat();
        assertEquals(NumberFormat.DEFAULT_NUMBER, numberFormat.getNumber()); // default is none
        numberFormat.setNumber(number);
        assertEquals(number, numberFormat.getNumber());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the CustomFormatCode property")
    @ValueSource(strings = {"//", "#.###"})
    void customFormatCodeTest(String value) {
        NumberFormat numberFormat = new NumberFormat();
        assertEquals("", numberFormat.getCustomFormatCode());
        numberFormat.setCustomFormatCode(value);
        assertEquals(value, numberFormat.getCustomFormatCode());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing set function of the CustomFormatCode property on invalid values")
    @NullSource // Adds null as value
    @ValueSource(strings = "")
    void customFormatCodeFailTest(String value) {
        NumberFormat numberFormat = new NumberFormat();
        FormatException exception = assertThrows(FormatException.class, () -> numberFormat.setCustomFormatCode(value));
        assertEquals(FormatException.class, exception.getClass());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the CustomFormatID property")
    @ValueSource(ints = {164, 200})
    void customFormatIDTest(int value) {
        NumberFormat numberFormat = new NumberFormat();
        assertEquals(164, numberFormat.getCustomFormatId());
        numberFormat.setCustomFormatId(value);
        assertEquals(value, numberFormat.getCustomFormatId());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing set function of the CustomFormatID property (invalid values)")
    @ValueSource(ints = {163, 0, -100})
    void customFormatIDFailTest(int value) {
        NumberFormat numberFormat = new NumberFormat();
        StyleException exception = assertThrows(StyleException.class, () -> numberFormat.setCustomFormatId(value));
        assertEquals(StyleException.class, exception.getClass());
    }

    @ParameterizedTest
    @DisplayName("Test of the get function of the IsCustomFormat property")
    @CsvSource({
            "NONE, false",
            "FORMAT_10, false",
            "CUSTOM, true"
    })
    void isCustomFormatTest(NumberFormat.FormatNumber number, boolean expectedResult) {
        NumberFormat numberFormat = new NumberFormat();
        assertFalse(numberFormat.isCustomFormat());
        numberFormat.setNumber(number);
        assertEquals(expectedResult, numberFormat.isCustomFormat());
    }

    @ParameterizedTest
    @DisplayName("Test of the IsDateFormat method")
    @CsvSource({
            "NONE, false",
            "FORMAT_1, false",
            "FORMAT_2, false",
            "FORMAT_3, false",
            "FORMAT_4, false",
            "FORMAT_5, false",
            "FORMAT_6, false",
            "FORMAT_7, false",
            "FORMAT_8, false",
            "FORMAT_9, false",
            "FORMAT_10, false",
            "FORMAT_11, false",
            "FORMAT_12, false",
            "FORMAT_13, false",
            "FORMAT_14, true",
            "FORMAT_15, true",
            "FORMAT_16, true",
            "FORMAT_17, true",
            "FORMAT_18, false",
            "FORMAT_19, false",
            "FORMAT_20, false",
            "FORMAT_21, false",
            "FORMAT_22, true",
            "FORMAT_37, false",
            "FORMAT_38, false",
            "FORMAT_39, false",
            "FORMAT_40, false",
            "FORMAT_45, false",
            "FORMAT_46, false",
            "FORMAT_47, false",
            "FORMAT_48, false",
            "FORMAT_49, false",
            "CUSTOM, false"
    })
    void isDateFormatTest(NumberFormat.FormatNumber number, boolean expectedDate) {
        assertEquals(expectedDate, NumberFormat.isDateFormat(number));
    }

    @ParameterizedTest
    @DisplayName("Test of the IsTimeFormat method")
    @CsvSource({
            "NONE, false",
            "FORMAT_1, false",
            "FORMAT_2, false",
            "FORMAT_3, false",
            "FORMAT_4, false",
            "FORMAT_5, false",
            "FORMAT_6, false",
            "FORMAT_7, false",
            "FORMAT_8, false",
            "FORMAT_9, false",
            "FORMAT_10, false",
            "FORMAT_11, false",
            "FORMAT_12, false",
            "FORMAT_13, false",
            "FORMAT_14, false",
            "FORMAT_15, false",
            "FORMAT_16, false",
            "FORMAT_17, false",
            "FORMAT_18, true",
            "FORMAT_19, true",
            "FORMAT_20, true",
            "FORMAT_21, true",
            "FORMAT_22, false",
            "FORMAT_37, false",
            "FORMAT_38, false",
            "FORMAT_39, false",
            "FORMAT_40, false",
            "FORMAT_45, true",
            "FORMAT_46, true",
            "FORMAT_47, true",
            "FORMAT_48, false",
            "FORMAT_49, false",
            "CUSTOM, false"
    })
    void isTimeFormatTest(NumberFormat.FormatNumber number, boolean expectedTime) {
        assertEquals(expectedTime, NumberFormat.isTimeFormat(number));
    }

    @ParameterizedTest
    @DisplayName("Test of the TryParseFormatNumber method")
    @CsvSource({
            "0, DEFINED_FORMAT, NONE",
            "-1, INVALID, NONE",
            "22, DEFINED_FORMAT, FORMAT_22",
            "23, UNDEFINED, NONE",
            "163, UNDEFINED, NONE",
            "164, DEFINED_FORMAT, CUSTOM",
            "165, CUSTOM_FORMAT, CUSTOM",
            "700, CUSTOM_FORMAT, CUSTOM"
    })
    void tryParseFormatNumberTest(int givenNumber, NumberFormat.FormatRange expectedRange,
            NumberFormat.FormatNumber expectedFormatNumber) {
        NumberFormat.NumberFormatEvaluation evaluation = NumberFormat.tryParseFormatNumber(givenNumber);
        assertEquals(expectedRange, evaluation.range());
        assertEquals(expectedFormatNumber, evaluation.formatNumber());
    }

    @Test
    @DisplayName("Test of the Equals method")
    void equalsTest() {
        NumberFormat style2 = (NumberFormat) exampleStyle.copy();
        assertTrue(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Number)")
    void equalsTest2() {
        NumberFormat style2 = (NumberFormat) exampleStyle.copy();
        style2.setNumber(NumberFormat.FormatNumber.FORMAT_15);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of CustomFormatCode)")
    void equalsTest2b() {
        NumberFormat style2 = (NumberFormat) exampleStyle.copy();
        style2.setCustomFormatCode("hh-mm-ss");
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of CustomFormatID)")
    void equalsTest2c() {
        NumberFormat style2 = (NumberFormat) exampleStyle.copy();
        style2.setCustomFormatId(180);
        assertFalse(exampleStyle.equals(style2));
    }

    @ParameterizedTest
    @DisplayName("Test of the Equals method (inequality on null or different objects)")
    @MethodSource("differentObjects") // see helper method differentObjects()
    void equalsTest3(Object obj) {
        assertFalse(exampleStyle.equals(obj));
    }

    @ParameterizedTest
    @DisplayName("Test of the Equals method when the origin object is null or not of the same type")
    @MethodSource("differentOriginObjects") // see helper method differentOriginObjects()
    void equalsTest5(Object origin) {
        NumberFormat copy = (NumberFormat) exampleStyle.copy();
        assertFalse(copy.equals(origin));
    }

    @Test
    @DisplayName("Test of the GetHashCode method (equality of two identical objects)")
    void getHashCodeTest() {
        NumberFormat copy = (NumberFormat) exampleStyle.copy();
        copy.setInternalId(Optional.of(99)); // Should not influence
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @Test
    @DisplayName("Test of the GetHashCode method (inequality of two different objects)")
    void getHashCodeTest2() {
        NumberFormat copy = (NumberFormat) exampleStyle.copy();
        copy.setNumber(NumberFormat.FormatNumber.FORMAT_14);
        assertNotEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @Test
    @DisplayName("Test of the constant of the default custom format start number")
    void defaultFontNameTest() {
        assertEquals(164, NumberFormat.CUSTOM_FORMAT_START_NUMBER); // Expected 164
    }

    @Test
    @DisplayName("Test of the CompareTo method")
    void compareToTest() {
        NumberFormat numberFormat = new NumberFormat();
        NumberFormat other = new NumberFormat();
        numberFormat.setInternalId(Optional.empty());
        other.setInternalId(Optional.empty());
        assertEquals(-1, numberFormat.compareTo(other));
        numberFormat.setInternalId(Optional.of(5));
        assertEquals(1, numberFormat.compareTo(other));
        assertEquals(1, numberFormat.compareTo(null));
        other.setInternalId(Optional.of(5));
        assertEquals(0, numberFormat.compareTo(other));
        other.setInternalId(Optional.of(4));
        assertEquals(1, numberFormat.compareTo(other));
        other.setInternalId(Optional.of(6));
        assertEquals(-1, numberFormat.compareTo(other));
    }

    // For code coverage
    @Test
    @DisplayName("Test of the ToString function")
    void toStringTest() {
        NumberFormat numberFormat = new NumberFormat();
        String s1 = numberFormat.toString();
        numberFormat.setNumber(NumberFormat.FormatNumber.FORMAT_11);
        assertNotEquals(s1, numberFormat.toString()); // An explicit value comparison is probably not sensible
    }

    private static Stream<Arguments> differentObjects() {
        return Stream.of(Arguments.of((Object) null), Arguments.of("text"), Arguments.of(true));
    }

    private static Stream<Arguments> differentOriginObjects() {
        return Stream.of(Arguments.of((Object) null), Arguments.of(true), Arguments.of("origin"));
    }
}
