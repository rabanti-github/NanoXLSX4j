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
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Optional;
import java.util.stream.Stream;

import ch.rabanti.nanoxlsx4j.colors.AutoColor;
import ch.rabanti.nanoxlsx4j.colors.Color;
import ch.rabanti.nanoxlsx4j.colors.IndexedColor;
import ch.rabanti.nanoxlsx4j.colors.SrgbColor;
import ch.rabanti.nanoxlsx4j.colors.SystemColor;
import ch.rabanti.nanoxlsx4j.colors.ThemeColor;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.themes.Theme;
import ch.rabanti.nanoxlsx4j.utils.Validators;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.MethodSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;

class FillTest {

    private final Fill exampleStyle;
    private final Fill comparisonStyle;

    FillTest() {
        exampleStyle = new Fill();
        exampleStyle.setBackgroundColor("FFAABB00");
        exampleStyle.setForegroundColor(Color.createIndexed(IndexedColor.Value.BRIGHT_GREEN));
        exampleStyle.setPatternFill(Fill.PatternValue.DARK_GRAY);

        comparisonStyle = new Fill();
        exampleStyle.setBackgroundColor("77CCBB00");
        exampleStyle.setForegroundColor(Color.createIndexed(IndexedColor.Value.BLUE_4));
        exampleStyle.setPatternFill(Fill.PatternValue.LIGHT_GRAY);
    }

    @Test
    @DisplayName("Test of the default values")
    void defaultValuesTest() {
        assertEquals("FF000000", Fill.DEFAULT_COLOR.getArgbValue());
        assertEquals(IndexedColor.DEFAULT_INDEXED_COLOR, Fill.DEFAULT_INDEXED_COLOR.getIndexedColor().getColorValue());
        assertEquals(Fill.PatternValue.NONE, Fill.DEFAULT_PATTERN_FILL);
    }

    @Test
    @DisplayName("Test of the constructor")
    void constructorTest() {
        Fill fill = new Fill();
        assertEquals(Fill.DEFAULT_PATTERN_FILL, fill.getPatternFill());
        assertEquals(Fill.DEFAULT_COLOR, fill.getForegroundColor());
        assertEquals(Fill.DEFAULT_COLOR, fill.getBackgroundColor());
    }

    @Test
    @DisplayName("Test of the constructor with colors")
    void constructorTest2() {
        Fill fill = new Fill("FFAABBCC", "FF001122");
        assertEquals(Fill.PatternValue.SOLID, fill.getPatternFill());
        assertEquals("FFAABBCC", fill.getForegroundColor().getArgbValue());
        assertEquals("FF001122", fill.getBackgroundColor().getArgbValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the constructor with color and fill type")
    @CsvSource({
            "FFAABBCC, FILL_COLOR, FFAABBCC, FF000000",
            "FF112233, PATTERN_COLOR, FF000000, FF112233"
    })
    void constructorTest3(String color, Fill.FillType fillType, String expectedForeground, String expectedBackground) {
        Fill fill = new Fill(color, fillType);
        assertEquals(Fill.PatternValue.SOLID, fill.getPatternFill());
        assertEquals(expectedForeground, fill.getForegroundColor().getArgbValue());
        assertEquals(expectedBackground, fill.getBackgroundColor().getArgbValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing constructor")
    @CsvSource(value = {
            "'', FF000000",
            "FF000000, ''",
            "NULL, FF000000",
            "FF000000, NULL",
            "'', ''",
            "NULL, NULL",
            "FF00000000, FFAABBCC",
            "FF000000, FFAABBCCCC",
            "FF0000, FAABBCC",
            "FF0000, '#FFAABBCC'",
            "F000000, FFAABB",
            "'#FF000000', FFAABB",
            "x, FFAABBCC",
            "FF000000, x",
            "x, y"
    }, nullValues = "NULL")
    void constructorFailTest(String foreground, String background) {
        assertThrows(StyleException.class, () -> new Fill(foreground, background));
    }

    @ParameterizedTest
    @DisplayName("Test of the failing constructor with color and fill type")
    @CsvSource(value = {
            "'', FILL_COLOR",
            "NULL, FILL_COLOR",
            "x, FILL_COLOR",
            "FFAABBCCDD, FILL_COLOR",
            "FAABB, FILL_COLOR",
            "'#FFAABB', FILL_COLOR",
            "'', PATTERN_COLOR",
            "NULL, PATTERN_COLOR",
            "x, PATTERN_COLOR",
            "FFAABBCCDD, PATTERN_COLOR",
            "FAABB, PATTERN_COLOR",
            "'#FFAABB', PATTERN_COLOR"
    }, nullValues = "NULL")
    void constructorFailTest2(String color, Fill.FillType fillType) {
        assertThrows(StyleException.class, () -> new Fill(color, fillType));
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the BackgroundColor property")
    @ValueSource(strings = {"77CCBB00", "00000000"})
    void backgroundColorTest(String value) {
        Fill fill = new Fill();
        assertEquals(Fill.DEFAULT_COLOR, fill.getBackgroundColor());
        fill.setBackgroundColor(value);
        assertEquals(value, fill.getBackgroundColor().getArgbValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing set function of the BackgroundColor property with invalid values")
    @NullSource
    @ValueSource(strings = {"7BB00", "#77BB00", "0002200000", "", "XXXXXXXX"})
    void backgroundColorFailTest(String value) {
        Fill fill = new Fill();
        StyleException exception = assertThrows(StyleException.class, () -> fill.setBackgroundColor(value));
        assertEquals(StyleException.class, exception.getClass());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the ForegroundColor property")
    @ValueSource(strings = {"77CCBB00", "FFFFFFFF"})
    void foregroundColorTest(String value) {
        Fill fill = new Fill();
        assertEquals(Fill.DEFAULT_COLOR, fill.getForegroundColor());
        fill.setForegroundColor(value);
        assertEquals(value, fill.getForegroundColor().getArgbValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing set function of the ForegroundColor property with invalid values")
    @NullSource
    @ValueSource(strings = {"7BB00", "#77BB00", "0002200000", "", "XXXXXXXX"})
    void foregroundColorFailTest(String value) {
        Fill fill = new Fill();
        StyleException exception = assertThrows(StyleException.class, () -> fill.setForegroundColor(value));
        assertEquals(StyleException.class, exception.getClass());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the IndexedColor property")
    @EnumSource(IndexedColor.Value.class) // Adds all enum value
    void indexedColorTest(IndexedColor.Value value) {
        Fill fill = new Fill();
        assertNull(fill.getForegroundColor().getIndexedColor());
        assertEquals(Color.ColorType.RGB, fill.getForegroundColor().getType()); // RGB is default
        fill.setForegroundColor(Color.createIndexed(value));
        assertEquals(value, fill.getForegroundColor().getIndexedColor().getColorValue());
        assertEquals(Color.ColorType.INDEXED, fill.getForegroundColor().getType());

        fill = new Fill();
        assertNull(fill.getBackgroundColor().getIndexedColor());
        fill.setBackgroundColor(Color.createIndexed(value));
        assertEquals(value, fill.getBackgroundColor().getIndexedColor().getColorValue());
        assertEquals(Color.ColorType.INDEXED, fill.getBackgroundColor().getType());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the PatternFill property")
    @EnumSource(Fill.PatternValue.class) // Adds all enum value
    void patternFillTest(Fill.PatternValue value) {
        Fill fill = new Fill();
        assertEquals(Fill.DEFAULT_PATTERN_FILL, fill.getPatternFill()); // default is none
        fill.setPatternFill(value);
        assertEquals(value, fill.getPatternFill());
    }

    @ParameterizedTest
    @DisplayName("Test of the SetColor function, when using a ARGB value")
    @CsvSource({
            "FFAABBCC, FILL_COLOR, FFAABBCC, FF000000",
            "FF112233, PATTERN_COLOR, FF000000, FF112233"
    })
    void setColorTest(String color, Fill.FillType fillType, String expectedForeground, String expectedBackground) {
        Fill fill = new Fill();
        assertEquals(Fill.DEFAULT_COLOR, fill.getForegroundColor());
        assertEquals(Fill.DEFAULT_COLOR, fill.getBackgroundColor());
        assertEquals(Fill.PatternValue.NONE, fill.getPatternFill());
        fill.setColor(color, fillType);
        assertEquals(Color.ColorType.RGB, fill.getForegroundColor().getType());
        assertEquals(Color.ColorType.RGB, fill.getBackgroundColor().getType());
        assertEquals(Fill.PatternValue.SOLID, fill.getPatternFill());
        assertEquals(expectedForeground, fill.getForegroundColor().getArgbValue());
        assertEquals(expectedBackground, fill.getBackgroundColor().getArgbValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the SetColor function, when using a sRGB color object")
    @CsvSource({
            "FFAABBCC, FILL_COLOR, FFAABBCC, FF000000",
            "FF112233, PATTERN_COLOR, FF000000, FF112233"
    })
    void setColorTest2(String colorValue, Fill.FillType fillType, String expectedForeground,
            String expectedBackground) {
        Fill fill = new Fill();
        SrgbColor color = new SrgbColor(colorValue);
        assertEquals(Fill.DEFAULT_COLOR, fill.getForegroundColor());
        assertEquals(Fill.DEFAULT_COLOR, fill.getBackgroundColor());
        assertEquals(Fill.PatternValue.NONE, fill.getPatternFill());
        fill.setColor(color, fillType);
        assertEquals(Color.ColorType.RGB, fill.getForegroundColor().getType());
        assertEquals(Color.ColorType.RGB, fill.getBackgroundColor().getType());
        assertEquals(Fill.PatternValue.SOLID, fill.getPatternFill());
        assertEquals(expectedForeground, fill.getForegroundColor().getArgbValue());
        assertEquals(expectedBackground, fill.getBackgroundColor().getArgbValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the SetColor function, when using an indexed color object")
    @CsvSource({
            "BLACK, FILL_COLOR, true, FF000000",
            "BLACK, PATTERN_COLOR, false, FF000000",
            "BLUE_4, FILL_COLOR, true, FF000000",
            "INDIGO, PATTERN_COLOR, false, FF000000",
            "SYSTEM_BACKGROUND, FILL_COLOR, true, FF000000",
            "SYSTEM_FOREGROUND, PATTERN_COLOR, false, FF000000"
    })
    void setColorTest3(IndexedColor.Value indexedValue, Fill.FillType fillType, boolean expectedForegroundHasColor,
            String expectedOppositeColor) {
        Fill fill = new Fill();
        IndexedColor color = new IndexedColor(indexedValue);
        assertEquals(Fill.DEFAULT_COLOR, fill.getForegroundColor());
        assertEquals(Fill.DEFAULT_COLOR, fill.getBackgroundColor());
        assertEquals(Fill.PatternValue.NONE, fill.getPatternFill());
        fill.setColor(color, fillType);
        if (expectedForegroundHasColor) {
            assertEquals(Color.ColorType.INDEXED, fill.getForegroundColor().getType());
            assertEquals(Color.ColorType.RGB, fill.getBackgroundColor().getType());
            assertEquals(color.getSrgbColor().getColorValue(), fill.getForegroundColor().getArgbValue());
            assertEquals(expectedOppositeColor, fill.getBackgroundColor().getArgbValue());
        } else {
            assertEquals(Color.ColorType.RGB, fill.getForegroundColor().getType());
            assertEquals(Color.ColorType.INDEXED, fill.getBackgroundColor().getType());
            assertEquals(expectedOppositeColor, fill.getForegroundColor().getArgbValue());
            assertEquals(color.getSrgbColor().getColorValue(), fill.getBackgroundColor().getArgbValue());
        }
        assertEquals(Fill.PatternValue.SOLID, fill.getPatternFill());
    }

    @ParameterizedTest
    @DisplayName("Test of the SetColor function, when using a theme color object")
    @CsvSource({
            "ACCENT_1, FILL_COLOR, true, FF000000",
            "ACCENT_1, PATTERN_COLOR, false, FF000000",
            "DARK_1, FILL_COLOR, true, FF000000",
            "LIGHT_1, PATTERN_COLOR, false, FF000000",
            "HYPERLINK, FILL_COLOR, true, FF000000",
            "FOLLOWED_HYPERLINK, PATTERN_COLOR, false, FF000000"
    })
    void setColorTest4(Theme.ColorSchemeElement themeValue, Fill.FillType fillType,
            boolean expectedForegroundHasColor, String expectedOppositeColor) {
        Fill fill = new Fill();
        ThemeColor color = new ThemeColor(themeValue);
        assertEquals(Fill.DEFAULT_COLOR, fill.getForegroundColor());
        assertEquals(Fill.DEFAULT_COLOR, fill.getBackgroundColor());
        assertEquals(Fill.PatternValue.NONE, fill.getPatternFill());
        fill.setColor(color, fillType);
        if (expectedForegroundHasColor) {
            assertEquals(Color.ColorType.THEME, fill.getForegroundColor().getType());
            assertEquals(Color.ColorType.RGB, fill.getBackgroundColor().getType());
            assertEquals(color.getStringValue(), fill.getForegroundColor().getThemeColor().getStringValue());
            assertEquals(expectedOppositeColor, fill.getBackgroundColor().getArgbValue());
        } else {
            assertEquals(Color.ColorType.RGB, fill.getForegroundColor().getType());
            assertEquals(Color.ColorType.THEME, fill.getBackgroundColor().getType());
            assertEquals(expectedOppositeColor, fill.getForegroundColor().getArgbValue());
            assertEquals(color.getStringValue(), fill.getBackgroundColor().getThemeColor().getStringValue());
        }
        assertEquals(Fill.PatternValue.SOLID, fill.getPatternFill());
    }

    @ParameterizedTest
    @DisplayName("Test of the SetColor function, when using a system color object")
    @CsvSource({
            "ACTIVE_BORDER, FILL_COLOR, true, FF000000",
            "BACKGROUND, PATTERN_COLOR, false, FF000000",
            "BUTTON_FACE, FILL_COLOR, true, FF000000",
            "MENU, PATTERN_COLOR, false, FF000000",
            "WINDOW, FILL_COLOR, true, FF000000",
            "CAPTION_TEXT, PATTERN_COLOR, false, FF000000"
    })
    void setColorTest5(SystemColor.Value systemValue, Fill.FillType fillType, boolean expectedForegroundHasColor,
            String expectedOppositeColor) {
        Fill fill = new Fill();
        SystemColor color = new SystemColor(systemValue);
        assertEquals(Fill.DEFAULT_COLOR, fill.getForegroundColor());
        assertEquals(Fill.DEFAULT_COLOR, fill.getBackgroundColor());
        assertEquals(Fill.PatternValue.NONE, fill.getPatternFill());
        fill.setColor(color, fillType);
        if (expectedForegroundHasColor) {
            assertEquals(Color.ColorType.SYSTEM, fill.getForegroundColor().getType());
            assertEquals(Color.ColorType.RGB, fill.getBackgroundColor().getType());
            assertEquals(color.getStringValue(), fill.getForegroundColor().getSystemColor().getStringValue());
            assertEquals(expectedOppositeColor, fill.getBackgroundColor().getArgbValue());
        } else {
            assertEquals(Color.ColorType.RGB, fill.getForegroundColor().getType());
            assertEquals(Color.ColorType.SYSTEM, fill.getBackgroundColor().getType());
            assertEquals(expectedOppositeColor, fill.getForegroundColor().getArgbValue());
            assertEquals(color.getStringValue(), fill.getBackgroundColor().getSystemColor().getStringValue());
        }
        assertEquals(Fill.PatternValue.SOLID, fill.getPatternFill());
    }

    @ParameterizedTest
    @DisplayName("Test of the SetColor function, when using an auto color object")
    @CsvSource({
            "FILL_COLOR, true, FF000000",
            "PATTERN_COLOR, false, FF000000"
    })
    void setColorTest6(Fill.FillType fillType, boolean expectedForegroundHasColor, String expectedOppositeColor) {
        Fill fill = new Fill();
        AutoColor color = new AutoColor();
        assertEquals(Fill.DEFAULT_COLOR, fill.getForegroundColor());
        assertEquals(Fill.DEFAULT_COLOR, fill.getBackgroundColor());
        assertEquals(Fill.PatternValue.NONE, fill.getPatternFill());
        fill.setColor(color, fillType);
        if (expectedForegroundHasColor) {
            assertEquals(Color.ColorType.AUTO, fill.getForegroundColor().getType());
            assertEquals(Color.ColorType.RGB, fill.getBackgroundColor().getType());
            assertTrue(fill.getForegroundColor().isAuto());
            assertFalse(fill.getBackgroundColor().isAuto());
            assertEquals(expectedOppositeColor, fill.getBackgroundColor().getArgbValue());
        } else {
            assertEquals(Color.ColorType.RGB, fill.getForegroundColor().getType());
            assertEquals(Color.ColorType.AUTO, fill.getBackgroundColor().getType());
            assertFalse(fill.getForegroundColor().isAuto());
            assertTrue(fill.getBackgroundColor().isAuto());
            assertEquals(expectedOppositeColor, fill.getForegroundColor().getArgbValue());
        }
        assertEquals(Fill.PatternValue.SOLID, fill.getPatternFill());
    }

    @ParameterizedTest
    @DisplayName("Test of the SetColor function, when a compound color object was passed")
    @CsvSource({
            "FILL_COLOR, true, FF000000",
            "PATTERN_COLOR, false, FF000000"
    })
    void setColorTest7(Fill.FillType fillType, boolean expectedForegroundHasColor, String expectedOppositeColor) {
        Fill fill = new Fill();
        Color color = Color.createRgb("FF112233");
        assertEquals(Fill.DEFAULT_COLOR, fill.getForegroundColor());
        assertEquals(Fill.DEFAULT_COLOR, fill.getBackgroundColor());
        assertEquals(Fill.PatternValue.NONE, fill.getPatternFill());
        fill.setColor(color, fillType);
        if (expectedForegroundHasColor) {
            assertEquals(Color.ColorType.RGB, fill.getForegroundColor().getType());
            assertEquals(Color.ColorType.RGB, fill.getBackgroundColor().getType());
            assertEquals(color.getArgbValue(), fill.getForegroundColor().getArgbValue());
            assertEquals(expectedOppositeColor, fill.getBackgroundColor().getArgbValue());
        } else {
            assertEquals(Color.ColorType.RGB, fill.getForegroundColor().getType());
            assertEquals(Color.ColorType.RGB, fill.getBackgroundColor().getType());
            assertEquals(expectedOppositeColor, fill.getForegroundColor().getArgbValue());
            assertEquals(color.getArgbValue(), fill.getBackgroundColor().getArgbValue());
        }
        assertEquals(Fill.PatternValue.SOLID, fill.getPatternFill());
    }

    @ParameterizedTest
    @DisplayName("Test of the ValidateColor function")
    @CsvSource(value = {
            "'', false, false, false",
            "NULL, false, false, false",
            "'', true, false, false",
            "NULL, true, false, false",
            "'', false, true, true",
            "NULL, false, true, true",
            "'', true, true, true",
            "NULL, true, true, true",
            "FFAABBCC, false, false, false",
            "FFAABBCC, true, false, true",
            "FFAABBCC, false, true, false",
            "FFAABBCC, true, true, true",
            "FFAABB, false, false, true",
            "FFAABB, true, false, false",
            "FFAA, true, false, false",
            "FFAA, false, false, false",
            "FFAA, true, true, false",
            "FFAACCDDDD, true, false, false",
            "FFAACCDDDD, false, false, false",
            "FFAACCDDDD, true, true, false"
    }, nullValues = "NULL")
    void validateColorTest(String color, boolean useAlpha, boolean allowEmpty, boolean expectedValid) {
        if (expectedValid) {
            // Should not throw
            Validators.validateColor(color, useAlpha, allowEmpty);
        } else {
            assertThrows(StyleException.class, () -> Validators.validateColor(color, useAlpha, allowEmpty));
        }
    }

    @Test
    @DisplayName("Test of the CopyFill function")
    void copyFillTest() {
        Fill copy = exampleStyle.copyFill();
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @ParameterizedTest
    @DisplayName("Test of the implicit operator (create ARGB color by string)")
    @ValueSource(strings = {"FF000000", "AC000000", "FF0000FF", "FFFFFFFF", "FF123456"})
    void implicitOperatorTest(String value) {
        Fill fill = new Fill(value); // Java constructor equivalent of the C# implicit conversion
        Fill expectedFill = new Fill(value, Fill.FillType.FILL_COLOR);
        assertEquals(fill.getForegroundColor(), expectedFill.getForegroundColor());
        assertEquals(fill.getBackgroundColor(), expectedFill.getBackgroundColor());
    }

    @ParameterizedTest
    @DisplayName("Test of the implicit operator (create indexed color by enum value)")
    @ValueSource(strings = {"RED", "BLACK_0", "BLACK", "SYSTEM_FOREGROUND", "SYSTEM_BACKGROUND"})
    void implicitOperatorTest3(IndexedColor.Value value) {
        Fill fill = new Fill(value); // Java constructor equivalent of the C# implicit conversion
        Fill expectedFill = new Fill();
        expectedFill.setForegroundColor(Color.createIndexed(value));
        expectedFill.setPatternFill(Fill.PatternValue.SOLID);
        assertEquals(fill.getForegroundColor(), expectedFill.getForegroundColor());
        assertEquals(fill.getBackgroundColor(), expectedFill.getBackgroundColor());
    }

    @ParameterizedTest
    @DisplayName("Test of the implicit operator (create indexed color by int)")
    @CsvSource({
            "10, RED",
            "0, BLACK_0",
            "8, BLACK",
            "64, SYSTEM_FOREGROUND",
            "65, SYSTEM_BACKGROUND"
    })
    void implicitOperatorTest4(int givenValue, IndexedColor.Value expectedValue) {
        Fill fill = new Fill(givenValue); // Java constructor equivalent of the C# implicit conversion
        Fill expectedFill = new Fill();
        expectedFill.setForegroundColor(Color.createIndexed(expectedValue));
        expectedFill.setPatternFill(Fill.PatternValue.SOLID);
        assertEquals(fill.getForegroundColor(), expectedFill.getForegroundColor());
        assertEquals(fill.getBackgroundColor(), expectedFill.getBackgroundColor());
    }

    @Test
    @DisplayName("Test of the Equals method")
    void equalsTest() {
        Fill style2 = (Fill) exampleStyle.copy();
        assertTrue(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of BackgroundColor)")
    void equalsTest2a() {
        Fill style2 = (Fill) exampleStyle.copy();
        style2.setBackgroundColor("66880000");
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of ForegroundColor)")
    void equalsTest2b() {
        Fill style2 = (Fill) exampleStyle.copy();
        style2.setForegroundColor("AA330000");
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of IndexedColor)")
    void equalsTest2c() {
        Fill style2 = (Fill) exampleStyle.copy();
        style2.setForegroundColor(Color.createIndexed(IndexedColor.Value.OLIVE));
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of PatternFill)")
    void equalsTest2d() {
        Fill style2 = (Fill) exampleStyle.copy();
        style2.setPatternFill(Fill.PatternValue.SOLID);
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
        Fill copy = (Fill) exampleStyle.copy();
        assertFalse(copy.equals(origin));
    }

    @Test
    @DisplayName("Test of the GetHashCode method (equality of two identical objects)")
    void getHashCodeTest() {
        Fill copy = (Fill) exampleStyle.copy();
        copy.setInternalId(Optional.of(99)); // Should not influence
        assertEquals(exampleStyle.hashCode(), copy.hashCode()); // For code coverage
    }

    @Test
    @DisplayName("Test of the GetHashCode method (inequality of two different objects)")
    void getHashCodeTest2() {
        Fill copy = (Fill) exampleStyle.copy();
        copy.setBackgroundColor("778800FF");
        assertNotEquals(exampleStyle.hashCode(), copy.hashCode()); // For code coverage
    }

    @Test
    @DisplayName("Test of the CompareTo method")
    void compareToTest() {
        Fill fill = new Fill();
        Fill other = new Fill();
        fill.setInternalId(Optional.empty());
        other.setInternalId(Optional.empty());
        assertEquals(-1, fill.compareTo(other));
        fill.setInternalId(Optional.of(5));
        assertEquals(1, fill.compareTo(other));
        assertEquals(1, fill.compareTo(null));
        other.setInternalId(Optional.of(5));
        assertEquals(0, fill.compareTo(other));
        other.setInternalId(Optional.of(4));
        assertEquals(1, fill.compareTo(other));
        other.setInternalId(Optional.of(6));
        assertEquals(-1, fill.compareTo(other));
    }

    // For code coverage
    @Test
    @DisplayName("Test of the ToString function")
    void toStringTest() {
        Fill fill = new Fill();
        String s1 = fill.toString();
        fill.setForegroundColor("FFAABBCC");
        assertNotEquals(s1, fill.toString()); // An explicit value comparison is probably not sensible
    }

    private static Stream<Arguments> differentObjects() {
        return Stream.of(Arguments.of((Object) null), Arguments.of("text"), Arguments.of(true));
    }

    private static Stream<Arguments> differentOriginObjects() {
        return Stream.of(Arguments.of((Object) null), Arguments.of(true), Arguments.of("origin"));
    }
}
