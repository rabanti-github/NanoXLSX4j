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

import ch.rabanti.nanoxlsx4j.colors.Color;
import ch.rabanti.nanoxlsx4j.colors.IndexedColor;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.MethodSource;
import org.junit.jupiter.params.provider.ValueSource;

class FontTest {

    private final Font exampleStyle;

    FontTest() {
        exampleStyle = new Font();
        exampleStyle.setBold(true);
        exampleStyle.setItalic(true);
        exampleStyle.setUnderline(Font.UnderlineValue.DOUBLE);
        exampleStyle.setStrike(true);
        exampleStyle.setShadow(true);
        exampleStyle.setExtend(true);
        exampleStyle.setCharset(Font.CharsetValue.ANSI);
        exampleStyle.setSize(15);
        exampleStyle.setName("Arial");
        exampleStyle.setFamily(Font.FontFamilyValue.SCRIPT);
        exampleStyle.setColorValue("FF22AACC");
        exampleStyle.setScheme(Font.SchemeValue.MINOR);
        exampleStyle.setVerticalAlign(Font.VerticalTextAlignValue.SUBSCRIPT);
    }

    @Test
    @DisplayName("Test of the default values")
    void defaultValuesTest() {
        assertEquals(11f, Font.DEFAULT_FONT_SIZE);
        assertEquals(Font.FontFamilyValue.SWISS, Font.DEFAULT_FONT_FAMILY);
        assertEquals(Font.SchemeValue.MINOR, Font.DEFAULT_FONT_SCHEME);
        assertEquals(Font.VerticalTextAlignValue.NONE, Font.DEFAULT_VERTICAL_ALIGN);
        assertEquals("Calibri", Font.DEFAULT_FONT_NAME);
    }

    @Test
    @DisplayName("Test of the constructor")
    void constructorTest() {
        Font font = new Font();
        assertEquals(Font.DEFAULT_FONT_SIZE, font.getSize());
        assertEquals(Font.DEFAULT_FONT_NAME, font.getName());
        assertEquals(Font.DEFAULT_FONT_FAMILY, font.getFamily());
        assertEquals(Font.DEFAULT_FONT_SCHEME, font.getScheme());
        assertEquals(Font.DEFAULT_VERTICAL_ALIGN, font.getVerticalAlign());
        assertFalse(font.getColorValue().isDefined());
        assertNull(font.getColorValue().getValue());
        assertEquals(Font.CharsetValue.DEFAULT, font.getCharset());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Bold property")
    @ValueSource(booleans = {true, false})
    void boldTest(boolean value) {
        Font font = new Font();
        assertFalse(font.isBold());
        font.setBold(value);
        assertEquals(value, font.isBold());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Italic property")
    @ValueSource(booleans = {true, false})
    void italicTest(boolean value) {
        Font font = new Font();
        assertFalse(font.isItalic());
        font.setItalic(value);
        assertEquals(value, font.isItalic());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Underline property")
    @EnumSource(Font.UnderlineValue.class)
    void underlineTest(Font.UnderlineValue value) {
        Font font = new Font();
        assertEquals(Font.UnderlineValue.NONE, font.getUnderline());
        font.setUnderline(value);
        assertEquals(value, font.getUnderline());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Strike property")
    @ValueSource(booleans = {true, false})
    void strikeTest(boolean value) {
        Font font = new Font();
        assertFalse(font.isStrike());
        font.setStrike(value);
        assertEquals(value, font.isStrike());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Outline property")
    @ValueSource(booleans = {true, false})
    void outlineTest(boolean value) {
        Font font = new Font();
        assertFalse(font.isOutline());
        font.setOutline(value);
        assertEquals(value, font.isOutline());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Shadow property")
    @ValueSource(booleans = {true, false})
    void shadowTest(boolean value) {
        Font font = new Font();
        assertFalse(font.isShadow());
        font.setShadow(value);
        assertEquals(value, font.isShadow());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Condense property")
    @ValueSource(booleans = {true, false})
    void condenseTest(boolean value) {
        Font font = new Font();
        assertFalse(font.isCondense());
        font.setCondense(value);
        assertEquals(value, font.isCondense());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Extend property")
    @ValueSource(booleans = {true, false})
    void extendTest(boolean value) {
        Font font = new Font();
        assertFalse(font.isExtend());
        font.setExtend(value);
        assertEquals(value, font.isExtend());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Charset property")
    @EnumSource(Font.CharsetValue.class) // adds all enum values
    void charsetTest(Font.CharsetValue value) {
        Font font = new Font();
        assertEquals(Font.CharsetValue.DEFAULT, font.getCharset());
        font.setCharset(value);
        assertEquals(value, font.getCharset());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Size property")
    @ValueSource(ints = {8, 75, 11})
    void sizeTest(int value) {
        Font font = new Font();
        assertEquals(Font.DEFAULT_FONT_SIZE, font.getSize()); // 11 is default
        font.setSize(value);
        assertEquals(value, font.getSize());
    }

    @ParameterizedTest
    @DisplayName("Test of the auto-adjusting set function of the Size property (invalid values)")
    @CsvSource({
            "0, 1",
            "7, 7",
            "-100, 1",
            "0.5, 1",
            "200, 200",
            "500, 409",
            "409.05, 409"
    })
    void sizeFailTest(float givenValue, float expectedValue) {
        Font font = new Font();
        font.setSize(givenValue);
        assertEquals(expectedValue, font.getSize());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Name property")
    @ValueSource(strings = {"Calibri", "Arial", "---"}) // "---" is not a font but a valid string as font definition
    void nameTest(String value) {
        Font font = new Font();
        assertEquals(Font.DEFAULT_FONT_NAME, font.getName()); // Default is 'Calibri'
        font.setName(value);
        assertEquals(value, font.getName());
    }

    @Test
    @DisplayName("Test of the failing set function of the Name property")
    void nameFailTest() {
        Font font = new Font();
        assertThrows(StyleException.class, () -> font.setName(null));
        assertThrows(StyleException.class, () -> font.setName(""));
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Family property")
    @ValueSource(strings = {
            "NOT_APPLICABLE", "ROMAN", "SWISS", "MODERN", "SCRIPT", "DECORATIVE", "RESERVED_1", "RESERVED_2",
            "RESERVED_3", "RESERVED_4", "RESERVED_5", "RESERVED_6", "RESERVED_7"
    })
    void familyTest(Font.FontFamilyValue value) {
        Font font = new Font();
        assertEquals(Font.DEFAULT_FONT_FAMILY, font.getFamily());
        font.setFamily(value);
        assertEquals(value, font.getFamily());
    }

    @Test
    @DisplayName("Test of the get and set function of the ColorValue property on sRGB (ARGB)")
    void colorValueTest() {
        Font font = new Font();
        assertEquals(Color.ColorType.NONE, font.getColorValue().getType()); // default is none
        font.setColorValue("FFAA3322"); // implicit in C#
        assertEquals(Color.ColorType.RGB, font.getColorValue().getType());
        assertEquals("FFAA3322", font.getColorValue().getRgbColor().getColorValue());

        Font font2 = new Font();
        assertEquals(Color.ColorType.NONE, font2.getColorValue().getType()); // default is none
        font2.setColorValue(Color.createRgb("FFAA33AA"));
        assertEquals(Color.ColorType.RGB, font2.getColorValue().getType());
        assertEquals("FFAA33AA", font2.getColorValue().getRgbColor().getColorValue());
    }

    @Test
    @DisplayName("Test of the get and set function of the ColorValue property on Indexed colors")
    void colorValueTest2() {
        Font font = new Font();
        assertEquals(Color.ColorType.NONE, font.getColorValue().getType()); // default is none
        font.setColorValue(32); // implicit in C#
        assertEquals(Color.ColorType.INDEXED, font.getColorValue().getType());
        assertEquals(IndexedColor.Value.NAVY, font.getColorValue().getIndexedColor().getColorValue());

        Font font2 = new Font();
        assertEquals(Color.ColorType.NONE, font2.getColorValue().getType()); // default is none
        font2.setColorValue(Color.createIndexed(IndexedColor.Value.NAVY));
        assertEquals(Color.ColorType.INDEXED, font.getColorValue().getType());
        assertEquals(IndexedColor.Value.NAVY, font.getColorValue().getIndexedColor().getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing implicit set function of the ColorValue property (invalid string values)")
    @ValueSource(strings = {"77BB0", "0002200000", "XXXXXXXX"})
    void colorValueFailTest(String value) {
        Font font = new Font();
        StyleException exception = assertThrows(StyleException.class, () -> font.setColorValue(value));
        assertEquals(StyleException.class, exception.getClass());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing implicit set function of the ColorValue property (invalid int values)")
    @ValueSource(ints = {-10, 66, 999})
    void colorValueFailTest2(int index) {
        Font font = new Font();
        StyleException exception = assertThrows(StyleException.class, () -> font.setColorValue(index));
        assertEquals(StyleException.class, exception.getClass());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Scheme property")
    @EnumSource(Font.SchemeValue.class)
    void schemeTest(Font.SchemeValue value) {
        Font font = new Font();
        assertEquals(Font.DEFAULT_FONT_SCHEME, font.getScheme()); // default is minor
        font.setScheme(value);
        assertEquals(value, font.getScheme());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the VerticalAlign property")
    @EnumSource(Font.VerticalTextAlignValue.class)
    void verticalAlignTest(Font.VerticalTextAlignValue value) {
        Font font = new Font();
        assertEquals(Font.DEFAULT_VERTICAL_ALIGN, font.getVerticalAlign()); // default is none
        font.setVerticalAlign(value);
        assertEquals(value, font.getVerticalAlign());
    }

    @Test
    @DisplayName("Test of the get function of the IsDefaultFont property")
    void isDefaultFontTest() {
        Font font = new Font();
        assertTrue(font.isDefaultFont());
        font.setItalic(true);
        font.setName("XYZ");
        assertFalse(font.isDefaultFont());
    }

    @ParameterizedTest
    @DisplayName("Test of the automatic assignment of font schemes on font names")
    @CsvSource({
            "Calibri, MINOR",
            "'Calibri Light', MAJOR",
            "Arial, NONE",
            "---, NONE"
    })
    void validateFontSchemeTest(String fontName, Font.SchemeValue scheme) {
        Font font = new Font();
        font.setName(fontName);
        assertEquals(scheme, font.getScheme());
    }

    @Test
    @DisplayName("Test of the CopyFont function")
    void copyFontTest() {
        Font copy = exampleStyle.copyFont();
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @Test
    @DisplayName("Test of the Equals method")
    void equalsTest() {
        Font style2 = (Font) exampleStyle.copy();
        assertTrue(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Bold)")
    void equalsTest2a() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setBold(false);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Italic)")
    void equalsTest2b() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setItalic(false);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Underline)")
    void equalsTest2c() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setUnderline(Font.UnderlineValue.DOUBLE_ACCOUNTING);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Strike)")
    void equalsTest2e() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setStrike(false);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Charset)")
    void equalsTest2f() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setCharset(Font.CharsetValue.BIG_5);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Size)")
    void equalsTest2g() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setSize(33);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Name)")
    void equalsTest2h() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setName("Comic Sans");
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Family)")
    void equalsTest2i() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setFamily(Font.FontFamilyValue.RESERVED_5);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of ColorValue)")
    void equalsTest2k() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setColorValue("FF9988AA");
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Scheme)")
    void equalsTest2l() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setScheme(Font.SchemeValue.NONE);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of VerticalAlign)")
    void equalsTest2m() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setVerticalAlign(Font.VerticalTextAlignValue.NONE);
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
        Font copy = (Font) exampleStyle.copy();
        assertFalse(copy.equals(origin));
    }

    @Test
    @DisplayName("Test of the GetHashCode method (equality of two identical objects)")
    void getHashCodeTest() {
        Font copy = (Font) exampleStyle.copy();
        copy.setInternalId(Optional.of(99)); // Should not influence
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @Test
    @DisplayName("Test of the GetHashCode method (inequality of two different objects)")
    void getHashCodeTest2() {
        Font copy = (Font) exampleStyle.copy();
        copy.setBold(false);
        assertNotEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @Test
    @DisplayName("Test of the CompareTo method")
    void compareToTest() {
        Font font = new Font();
        Font other = new Font();
        font.setInternalId(Optional.empty());
        other.setInternalId(Optional.empty());
        assertEquals(-1, font.compareTo(other));
        font.setInternalId(Optional.of(5));
        assertEquals(1, font.compareTo(other));
        assertEquals(1, font.compareTo(null));
        other.setInternalId(Optional.of(5));
        assertEquals(0, font.compareTo(other));
        other.setInternalId(Optional.of(4));
        assertEquals(1, font.compareTo(other));
        other.setInternalId(Optional.of(6));
        assertEquals(-1, font.compareTo(other));
    }

    // For code coverage
    @Test
    @DisplayName("Test of the ToString function")
    void toStringTest() {
        Font font = new Font();
        String s1 = font.toString();
        font.setName("YXZ");
        assertNotEquals(s1, font.toString()); // An explicit value comparison is probably not sensible
    }

    private static Stream<Arguments> differentObjects() {
        return Stream.of(Arguments.of((Object) null), Arguments.of("text"), Arguments.of(true));
    }

    private static Stream<Arguments> differentOriginObjects() {
        return Stream.of(Arguments.of((Object) null), Arguments.of(true), Arguments.of("origin"));
    }
}
