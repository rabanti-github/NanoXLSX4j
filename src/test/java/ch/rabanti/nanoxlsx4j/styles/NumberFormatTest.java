package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.TestUtils;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.*;

public class NumberFormatTest {

    private NumberFormat exampleStyle;

    public NumberFormatTest() {
        exampleStyle = new NumberFormat();
        exampleStyle.setCustomFormatCode("#.###");
        exampleStyle.setNumber(NumberFormat.FormatNumber.format_10);
        exampleStyle.setCustomFormatID(170);
    }

    @DisplayName("Test of the get and set function of the formatNumber field")
    @ParameterizedTest(name = "Given value {0} should lead to defined field")
    @CsvSource({
            "none",
            "format_1",
            "format_2",
            "format_3",
            "format_4",
            "format_5",
            "format_6",
            "format_7",
            "format_8",
            "format_9",
            "format_10",
            "format_11",
            "format_12",
            "format_13",
            "format_14",
            "format_15",
            "format_16",
            "format_17",
            "format_18",
            "format_19",
            "format_20",
            "format_21",
            "format_22",
            "format_37",
            "format_38",
            "format_39",
            "format_40",
            "format_45",
            "format_46",
            "format_47",
            "format_48",
            "format_49",
            "custom",
    })
    void formatNumberTest(NumberFormat.FormatNumber number) {
        NumberFormat numberFormat = new NumberFormat();
        assertEquals(NumberFormat.DEFAULT_NUMBER, numberFormat.getNumber()); // default is none
        numberFormat.setNumber(number);
        assertEquals(number, numberFormat.getNumber());
    }

    @DisplayName("Test of the get and set function of the customFormatCode filed")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource({
            "STRING, ''",
            "STRING, '#.###'",
    })
    void customFormatCodeTest(String sourceType, String sourceValue) {
        String value = (String) TestUtils.createInstance(sourceType, sourceValue);
        NumberFormat numberFormat = new NumberFormat();
        assertEquals("", numberFormat.getCustomFormatCode());
        numberFormat.setCustomFormatCode(value);
        assertEquals(value, numberFormat.getCustomFormatCode());
    }

    @DisplayName("Test of the get and set function of the customFormatID field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource({
            "164",
            "200",
    })
    void customFormatIDTest(int value) {
        NumberFormat numberFormat = new NumberFormat();
        assertEquals(164, numberFormat.getCustomFormatID());
        numberFormat.setCustomFormatID(value);
        assertEquals(value, numberFormat.getCustomFormatID());
    }

    @DisplayName("Test of the failing set function of the CustomFormatID field (invalid values)")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource({
            "163",
            "0",
            "-100",
    })
    void customFormatIDFailTest(int value) {
        NumberFormat numberFormat = new NumberFormat();
        assertThrows(Exception.class, () -> numberFormat.setCustomFormatID(value));
    }

    @DisplayName("Test of the get function of the ssCustomFormat field")
    @ParameterizedTest(name = "Given value {0} should lead to the result {1}")
    @CsvSource({
            "none, false",
            "format_10, false",
            "custom, true",
    })
    void isCustomFormatTest(NumberFormat.FormatNumber number, boolean expectedResult) {
        NumberFormat numberFormat = new NumberFormat();
        assertFalse(numberFormat.isCustomFormat());
        numberFormat.setNumber(number);
        assertEquals(expectedResult, numberFormat.isCustomFormat());
    }

    @DisplayName("Test of the isDateFormat method")
    @ParameterizedTest(name = "Given value {0} should lead to the result {1}")
    @CsvSource({
            "none, false",
            "format_1, false",
            "format_2, false",
            "format_3, false",
            "format_4, false",
            "format_5, false",
            "format_6, false",
            "format_7, false",
            "format_8, false",
            "format_9, false",
            "format_10, false",
            "format_11, false",
            "format_12, false",
            "format_13, false",
            "format_14, true",
            "format_15, true",
            "format_16, true",
            "format_17, true",
            "format_18, false",
            "format_19, false",
            "format_20, false",
            "format_21, false",
            "format_22, true",
            "format_37, false",
            "format_38, false",
            "format_39, false",
            "format_40, false",
            "format_45, false",
            "format_46, false",
            "format_47, false",
            "format_48, false",
            "format_49, false",
            "custom, false",
    })
    void isDateFormatTest(NumberFormat.FormatNumber number, boolean expectedDate) {
        assertEquals(expectedDate, NumberFormat.isDateFormat(number));
    }

    @DisplayName("Test of the isTimeFormat method")
    @ParameterizedTest(name = "Given value {0} should lead to the result {1}")
    @CsvSource({
            "none, false",
            "format_1, false",
            "format_2, false",
            "format_3, false",
            "format_4, false",
            "format_5, false",
            "format_6, false",
            "format_7, false",
            "format_8, false",
            "format_9, false",
            "format_10, false",
            "format_11, false",
            "format_12, false",
            "format_13, false",
            "format_14, false",
            "format_15, false",
            "format_16, false",
            "format_17, false",
            "format_18, true",
            "format_19, true",
            "format_20, true",
            "format_21, true",
            "format_22, false",
            "format_37, false",
            "format_38, false",
            "format_39, false",
            "format_40, false",
            "format_45, true",
            "format_46, true",
            "format_47, true",
            "format_48, false",
            "format_49, false",
            "custom, false",
    })
    void isTimeFormatTest(NumberFormat.FormatNumber number, boolean expectedTime) {
        assertEquals(expectedTime, NumberFormat.isTimeFormat(number));
    }

    @DisplayName("Test of the TryParseFormatNumber method")
    @ParameterizedTest(name = "Given value {0} should lead to the expected range {1} and the format number {2}")
    @CsvSource({
            "0, defined_format, none",
            "-1, invalid, none",
            "22, defined_format, format_22",
            "23, undefined, none",
            "163, undefined, none",
            "164, defined_format, custom",
            "165, custom_format, custom",
            "700, custom_format, custom",
    })
    void tryParseFormatNumberTest(int givenNumber, NumberFormat.NumberFormatEvaluation.FormatRange expectedRange, NumberFormat.FormatNumber expectedFormatNumber) {
        NumberFormat.FormatNumber number;
        NumberFormat.NumberFormatEvaluation evaluation = NumberFormat.tryParseFormatNumber(givenNumber);
        assertEquals(expectedRange, evaluation.getRange());
        assertEquals(expectedFormatNumber, evaluation.getFormatNumber());
    }

    @DisplayName("Test of the Equals method")
    @Test()
    void equalsTest() {
        NumberFormat style2 = (NumberFormat) exampleStyle.copy();
        assertTrue(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of number)")
    @Test()
    void equalsTest2() {
        NumberFormat style2 = (NumberFormat) exampleStyle.copy();
        style2.setNumber(NumberFormat.FormatNumber.format_15);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of customFormatCode)")
    @Test()
    void equalsTest2b() {
        NumberFormat style2 = (NumberFormat) exampleStyle.copy();
        style2.setCustomFormatCode("hh-mm-ss");
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of customFormatID)")
    @Test()
    void equalsTest2c() {
        NumberFormat style2 = (NumberFormat) exampleStyle.copy();
        style2.setCustomFormatID(180);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality on null or different objects)")
    @ParameterizedTest(name = "Given value '{1}' should lead to an unequal result")
    @CsvSource({
            "NULL, ''",
            "STRING, 'text'",
            "BOOLEAN, 'true'",
    })
    void equalsTest3(String sourceType, String sourceValue) {
        Object obj = TestUtils.createInstance(sourceType, sourceValue);
        assertFalse(exampleStyle.equals(obj));
    }

    @DisplayName("Test of the Equals method when the origin object is null or not of the same type")
    @ParameterizedTest(name = "Given value '{1}' should lead to an unequal result")
    @CsvSource({
            "NULL, ''",
            "BOOLEAN, 'true'",
            "STRING, 'origin'",
    })
    void equalsTest5(String sourceType, String sourceValue) {
        Object origin = TestUtils.createInstance(sourceType, sourceValue);
        NumberFormat copy = (NumberFormat) exampleStyle.copy();
        assertFalse(copy.equals(origin));
    }

    @DisplayName("Test of the hashCode method (equality of two identical objects)")
    @Test()
    void getHashCodeTest() {
        NumberFormat copy = (NumberFormat) exampleStyle.copy();
        copy.setInternalID(99);  // Should not influence
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }


    @DisplayName("Test of the hashCode method (inequality of two different objects)")
    @Test()
    void getHashCodeTest2() {
        NumberFormat copy = (NumberFormat) exampleStyle.copy();
        copy.setNumber(NumberFormat.FormatNumber.format_14);
        assertNotEquals(exampleStyle.hashCode(), copy.hashCode());
    }


    @DisplayName("Test of the constant of the default custom format start number")
    @Test()
    void defaultFontNameTest() {
        assertEquals(164, NumberFormat.CUSTOMFORMAT_START_NUMBER); // Expected 164
    }

    // For code coverage

    @DisplayName("Test of the ToString function")
    @Test()
    void toStringTest() {
        NumberFormat numberFormat = new NumberFormat();
        String s1 = numberFormat.toString();
        numberFormat.setNumber(NumberFormat.FormatNumber.format_11);
        assertNotEquals(s1, numberFormat.toString()); // An explicit value comparison is probably not sensible
    }


}
