package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

public class FillTest {

    private final Fill exampleStyle;

    public FillTest() {
        exampleStyle = new Fill();
        exampleStyle.setBackgroundColor("FFAABB00");
        exampleStyle.setForegroundColor("1188FF00");
        exampleStyle.setIndexedColor(99);
        exampleStyle.setPatternFill(Fill.PatternValue.darkGray);
    }

    @DisplayName("Test of the default values")
    @Test()
    void defaultValuesTest() {
        assertEquals("FF000000", Fill.DEFAULT_COLOR);
        assertEquals(64, Fill.DEFAULT_INDEXED_COLOR);
        assertEquals(Fill.PatternValue.none, Fill.DEFAULT_PATTERN_FILL);
    }

    @DisplayName("Test of the constructor with colors")
    @Test()
    void constructorTest() {
        Fill fill = new Fill();
        assertEquals(Fill.DEFAULT_INDEXED_COLOR, fill.getIndexedColor());
        assertEquals(Fill.DEFAULT_PATTERN_FILL, fill.getPatternFill());
        assertEquals(Fill.DEFAULT_COLOR, fill.getForegroundColor());
        assertEquals(Fill.DEFAULT_COLOR, fill.getBackgroundColor());
    }

    @DisplayName("Test of the constructor")
    @Test()
    void constructorTest2() {
        Fill fill = new Fill("FFAABBCC",
                             "FF001122");
        assertEquals(Fill.DEFAULT_INDEXED_COLOR, fill.getIndexedColor());
        assertEquals(Fill.PatternValue.solid, fill.getPatternFill());
        assertEquals("FFAABBCC", fill.getForegroundColor());
        assertEquals("FF001122", fill.getBackgroundColor());
    }

    @DisplayName("Test of the constructor with color and fill type")
    @ParameterizedTest(name = "Given color {0} fill type {1} should lead to a valid fill object")
    @CsvSource(
            {
                    "FFAABBCC, fillColor, FFAABBCC, FF000000",
                    "FF112233, patternColor, FF000000, FF112233",}
    )
    void constructorTest3(String color, Fill.FillType fillType, String expectedForeground, String expectedBackground) {
        Fill fill = new Fill(color, fillType);
        assertEquals(Fill.DEFAULT_INDEXED_COLOR, fill.getIndexedColor());
        assertEquals(Fill.PatternValue.solid, fill.getPatternFill());
        assertEquals(expectedForeground, fill.getForegroundColor());
        assertEquals(expectedBackground, fill.getBackgroundColor());
    }

    @DisplayName("Test of the failing constructor")
    @ParameterizedTest(name = "Given foreground {1} or background {3} should lead to an exception")
    @CsvSource(
            {
                    "STRING, '', STRING , 'FF000000'",
                    "STRING, 'FF000000', STRING , ''",
                    "NULL, '', STRING , 'FF000000'",
                    "STRING, 'FF000000', NULL , ''",
                    "STRING, '', STRING , ''",
                    "NULL, '', NULL , ''",
                    "STRING, 'FF00000000', STRING , 'FFAABBCC'",
                    "STRING, 'FF000000', STRING , 'FFAABBCCCC'",
                    "STRING, 'FF0000', STRING , 'FFAABBCC'",
                    "STRING, 'FF000000', STRING , 'FFAABB'",
                    "STRING, 'x', STRING , 'FFAABBCC'",
                    "STRING, 'FF000000', STRING , 'x'",
                    "STRING, 'x', STRING , 'y'",}
    )
    void constructorFailTest(String foregroundSourceType, String foregroundSourceValue, String backgroundSourceType, String backgroundSourceValue) {
        String foreground = (String) TestUtils.createInstance(foregroundSourceType, foregroundSourceValue);
        String background = (String) TestUtils.createInstance(backgroundSourceType, backgroundSourceValue);
        assertThrows(StyleException.class, () -> new Fill(foreground, background));
    }

    @DisplayName("Test of the failing constructor with color and fill type")
    @ParameterizedTest(name = "Given values {0} and {1} should lead to an exception")
    @CsvSource(
            {
                    "STRING, '', fillColor",
                    "NULL, '', fillColor",
                    "STRING, 'x', fillColor",
                    "STRING, 'FFAABBCCDD', fillColor",
                    "STRING, 'FFAABB', fillColor",
                    "STRING, '', patternColor",
                    "NULL, '', patternColor",
                    "STRING, 'x', patternColor",
                    "STRING, 'FFAABBCCDD', patternColor",
                    "STRING, 'FFAABB', patternColor",}
    )
    void constructorFailTest2(String sourceType, String sourceValue, Fill.FillType fillType) {
        String color = (String) TestUtils.createInstance(sourceType, sourceValue);
        assertThrows(StyleException.class, () -> new Fill(color, fillType));
    }

    @DisplayName("Test of the get and set function of the backgroundColor field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource(
            {
                    "77CCBB00",
                    "00000000",}
    )
    void backgroundColorTest(String value) {
        Fill fill = new Fill();
        assertEquals(Fill.DEFAULT_COLOR, fill.getBackgroundColor());
        fill.setBackgroundColor(value);
        assertEquals(value, fill.getBackgroundColor());
    }

    @DisplayName("Test of the failing set function of the backgroundColor field with invalid values")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource(
            {
                    "STRING, '77BB00'",
                    "STRING, '0002200000'",
                    "STRING, ''",
                    "NULL, ''",
                    "STRING, 'XXXXXXXX'",}
    )
    void backgroundColorFailTest(String sourceType, String sourceValue) {
        String value = (String) TestUtils.createInstance(sourceType, sourceValue);
        Fill fill = new Fill();
        assertThrows(Exception.class, () -> fill.setBackgroundColor(value));
    }

    @DisplayName("Test of the get and set function of the foregroundColor field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource(
            {
                    "77CCBB00",
                    "FFFFFFFF",}
    )
    void foregroundColorTest(String value) {
        Fill fill = new Fill();
        assertEquals(Fill.DEFAULT_COLOR, fill.getForegroundColor());
        fill.setForegroundColor(value);
        assertEquals(value, fill.getForegroundColor());
    }

    @DisplayName("Test of the failing set function of the foregroundColor field with invalid values")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource(
            {
                    "STRING, '77BB00'",
                    "STRING, '0002200000'",
                    "STRING, ''",
                    "NULL, ''",
                    "STRING, 'XXXXXXXX'",}
    )
    void foregroundColorFailTest(String sourceType, String sourceValue) {
        String value = (String) TestUtils.createInstance(sourceType, sourceValue);
        Fill fill = new Fill();
        assertThrows(Exception.class, () -> fill.setForegroundColor(value));
    }

    @DisplayName("Test of the get and set function of the indexedColor field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource(
            {
                    "0",
                    "256",
                    "-10",}
    )
    void indexedColorTest(int value) {
        Fill fill = new Fill();
        assertEquals(Fill.DEFAULT_INDEXED_COLOR, fill.getIndexedColor()); // 64 is default
        fill.setIndexedColor(value);
        assertEquals(value, fill.getIndexedColor());
    }

    @DisplayName("Test of the get and set function of the patternFill field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource(
            {
                    "darkGray",
                    "gray0625",
                    "gray125",
                    "lightGray",
                    "mediumGray",
                    "none",
                    "solid",}
    )
    void patternFillTest(Fill.PatternValue value) {
        Fill fill = new Fill();
        assertEquals(Fill.DEFAULT_PATTERN_FILL, fill.getPatternFill()); // default is none
        fill.setPatternFill(value);
        assertEquals(value, fill.getPatternFill());
    }

    @DisplayName("Test of the setColor function")
    @ParameterizedTest(name = "Given color {0} and fill type {1} should lead to a fill with background color {2} and foreground color {3}")
    @CsvSource(
            {
                    "FFAABBCC, fillColor, FFAABBCC, FF000000",
                    "FF112233, patternColor, FF000000, FF112233",}
    )
    void setColorTest(String color, Fill.FillType fillType, String expectedForeground, String expectedBackground) {
        Fill fill = new Fill();
        assertEquals(Fill.DEFAULT_COLOR, fill.getForegroundColor());
        assertEquals(Fill.DEFAULT_COLOR, fill.getBackgroundColor());
        assertEquals(Fill.PatternValue.none, fill.getPatternFill());
        fill.setColor(color, fillType);
        assertEquals(Fill.DEFAULT_INDEXED_COLOR, fill.getIndexedColor());
        assertEquals(Fill.PatternValue.solid, fill.getPatternFill());
        assertEquals(expectedForeground, fill.getForegroundColor());
        assertEquals(expectedBackground, fill.getBackgroundColor());
    }

    @DisplayName("Test of the copyFill function")
    @Test()
    void copyFillTest() {
        Fill copy = exampleStyle.copy();
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @DisplayName("Test of the validateColor function")
    @ParameterizedTest(name = "Given value {1} with allowed alpha: {2} and allowed empty: {3} should lead be valid: {4}")
    @CsvSource(
            {
                    "STRING, '', false, false, false",
                    "NULL, '', false, false, false",
                    "STRING, '', true, false, false",
                    "NULL, '', true, false, false",
                    "STRING, '', false, true, true",
                    "NULL, '', false, true, true",
                    "STRING, '', true, true, true",
                    "NULL, '', true, true, true",
                    "STRING, FFAABBCC, false, false, false",
                    "STRING, FFAABBCC, true, false, true",
                    "STRING, FFAABBCC, false, true, false",
                    "STRING, FFAABBCC, true, true, true",
                    "STRING, FFAABB, false, false, true",
                    "STRING, FFAABB, true, false, false",
                    "STRING, FFAA, true, false, false",
                    "STRING, FFAA, false, false, false",
                    "STRING, FFAA, true, true, false",
                    "STRING, FFAACCDDDD, true, false, false",
                    "STRING, FFAACCDDDD, false, false, false",
                    "STRING, FFAACCDDDD, true, true, false",
                    "STRING, FFAACCDDDD, true, true, false",}
    )
    void validateColorTest(String sourceType, String sourceValue, boolean useAlpha, boolean allowEmpty, boolean expectedValid) {
        String color = (String) TestUtils.createInstance(sourceType, sourceValue);
        if (expectedValid) {
            // Should not throw
            Fill.validateColor(color, useAlpha, allowEmpty);
        } else {
            assertThrows(StyleException.class, () -> Fill.validateColor(color, useAlpha, allowEmpty));
        }

    }

    @DisplayName("Test of the equals method")
    @Test()
    void equalsTest() {
        Fill style2 = exampleStyle.copy();
        assertEquals(exampleStyle, style2);
    }

    @DisplayName("Test of the equals method (inequality of backgroundColor)")
    @Test()
    void equalsTest2a() {
        Fill style2 = exampleStyle.copy();
        style2.setBackgroundColor("66880000");
        assertNotEquals(exampleStyle, style2);
    }

    @DisplayName("Test of the equals method (inequality of foregroundColor)")
    @Test()
    void equalsTest2b() {
        Fill style2 = exampleStyle.copy();
        style2.setForegroundColor("AA330000");
        assertNotEquals(exampleStyle, style2);
    }

    @DisplayName("Test of the equals method (inequality of indexedColor)")
    @Test()
    void equalsTest2c() {
        Fill style2 = exampleStyle.copy();
        style2.setIndexedColor(78);
        assertNotEquals(exampleStyle, style2);
    }

    @DisplayName("Test of the equals method (inequality of patternFill)")
    @Test()
    void equalsTest2d() {
        Fill style2 = exampleStyle.copy();
        style2.setPatternFill(Fill.PatternValue.solid);
        assertNotEquals(exampleStyle, style2);
    }

    @DisplayName("Test of the equals method (inequality on null or different objects)")
    @ParameterizedTest(name = "Given value {1} should lead to an equal object")
    @CsvSource(
            {
                    "NULL, ''",
                    "STRING, 'text'",
                    "BOOLEAN, 'true'",}
    )
    void equalsTest3(String sourceType, String sourceValue) {
        Object obj = TestUtils.createInstance(sourceType, sourceValue);
        assertNotEquals(exampleStyle, obj);
    }

    @DisplayName("Test of the equals method when the origin object is null or not of the same type")
    @ParameterizedTest(name = "Given value {1} should lead to an equal object")
    @CsvSource(
            {
                    "NULL, ''",
                    "BOOLEAN, 'true'",
                    "STRING, 'origin'",}
    )
    void equalsTest5(String sourceType, String sourceValue) {
        Object origin = TestUtils.createInstance(sourceType, sourceValue);
        Fill copy = exampleStyle.copy();
        assertNotEquals(copy, origin);
    }

    @DisplayName("Test of the hashCode method (equality of two identical objects)")
    @Test()
    void hashCodeTest() {
        Fill copy = exampleStyle.copy();
        copy.setInternalID(99); // Should not influence
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
        assertEquals(exampleStyle.hashCode(), copy.hashCode()); // For code coverage
    }

    @DisplayName("Test of the hashCode method (inequality of two different objects)")
    @Test()
    void hashCodeTest2() {
        Fill copy = exampleStyle.copy();
        copy.setBackgroundColor("778800FF");
        assertNotEquals(exampleStyle.hashCode(), copy.hashCode());
        assertNotEquals(exampleStyle.hashCode(), copy.hashCode()); // For code coverage
    }

    @DisplayName("Test of the compareTo method")
    @Test()
    void compareToTest() {
        Fill fill = new Fill();
        Fill other = new Fill();
        fill.setInternalID(null);
        other.setInternalID(null);
        assertEquals(-1, fill.compareTo(other));
        fill.setInternalID(5);
        assertEquals(1, fill.compareTo(other));
        assertEquals(1, fill.compareTo(null));
        other.setInternalID(5);
        assertEquals(0, fill.compareTo(other));
        other.setInternalID(4);
        assertEquals(1, fill.compareTo(other));
        other.setInternalID(6);
        assertEquals(-1, fill.compareTo(other));
    }

    // For code coverage
    @DisplayName("Test of the ToString function")
    @Test()
    void toStringTest() {
        Fill fill = new Fill();
        String s1 = fill.toString();
        fill.setForegroundColor("FFAABBCC");
        assertNotEquals(s1, fill.toString()); // An explicit value comparison is probably not sensible
    }

    @DisplayName("Test of the getValue function of the PatternValue enum")
    @Test()
    void patternValueTest() {
        assertEquals(0, Fill.PatternValue.none.getValue());
        assertEquals(1, Fill.PatternValue.solid.getValue());
        assertEquals(2, Fill.PatternValue.darkGray.getValue());
        assertEquals(3, Fill.PatternValue.mediumGray.getValue());
        assertEquals(4, Fill.PatternValue.lightGray.getValue());
        assertEquals(5, Fill.PatternValue.gray0625.getValue());
        assertEquals(6, Fill.PatternValue.gray125.getValue());
    }

    @DisplayName("Test of the getValue function of the FillType enum")
    @Test()
    void fillTypeValueTest() {
        assertEquals(0, Fill.FillType.patternColor.getValue());
        assertEquals(1, Fill.FillType.fillColor.getValue());
    }

}
