package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class FontTest {

    private Font exampleStyle;

    public FontTest() {
        exampleStyle = new Font();
        exampleStyle.setBold(true);
        exampleStyle.setItalic(true);
        exampleStyle.setUnderline(Font.UnderlineValue.u_double);
        exampleStyle.setStrike(true);
        exampleStyle.setCharset("ASCII");
        exampleStyle.setSize(15);
        exampleStyle.setName("Arial");
        exampleStyle.setFamily("X");
        exampleStyle.setColorTheme(10);
        exampleStyle.setColorValue("FF22AACC");
        exampleStyle.setScheme(Font.SchemeValue.minor);
        exampleStyle.setVerticalAlign(Font.VerticalAlignValue.subscript);
    }

    @DisplayName("Test of the default values")
    @Test()
    void defaultValuesTest() {
        assertEquals(11f, Font.DEFAULT_FONT_SIZE);
        assertEquals("2", Font.DEFAULT_FONT_FAMILY);
        assertEquals(Font.SchemeValue.minor, Font.DEFAULT_FONT_SCHEME);
        assertEquals(Font.VerticalAlignValue.none, Font.DEFAULT_VERTICAL_ALIGN);
        assertEquals("Calibri", Font.DEFAULT_FONT_NAME);
    }

    @DisplayName("Test of the constructor")
    @Test()
    void constructorTest() {
        Font font = new Font();
        assertEquals(Font.DEFAULT_FONT_SIZE, font.getSize());
        assertEquals(Font.DEFAULT_FONT_NAME, font.getName());
        assertEquals(Font.DEFAULT_FONT_FAMILY, font.getFamily());
        assertEquals(Font.DEFAULT_FONT_SCHEME, font.getScheme());
        assertEquals(Font.DEFAULT_VERTICAL_ALIGN, font.getVerticalAlign());
        assertEquals("", font.getColorValue());
        assertEquals("", font.getCharset());
        assertEquals(1, font.getColorTheme());
    }

    @DisplayName("Test of the get and set function of the bold field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource({
            "true",
            "false",
    })
    void boldTest(boolean value) {
        Font font = new Font();
        assertFalse(font.isBold());
        font.setBold(value);
        assertEquals(value, font.isBold());
    }

    @DisplayName("Test of the get and set function of the italic field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource({
            "true",
            "false",
    })
    void italicTest(boolean value) {
        Font font = new Font();
        assertFalse(font.isItalic());
        font.setItalic(value);
        assertEquals(value, font.isItalic());
    }

    @DisplayName("Test of the get and set function of the underline field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource({
            "none",
            "doubleAccounting",
            "singleAccounting",
            "u_double",
            "u_single",
    })
    void underlineTest(Font.UnderlineValue value) {
        Font font = new Font();
        assertEquals(Font.UnderlineValue.none, font.getUnderline());
        font.setUnderline(value);
        assertEquals(value, font.getUnderline());
    }

    @DisplayName("Test of the get and set function of the strike field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource({
            "true",
            "false",
    })
    void strikeTest(boolean value) {
        Font font = new Font();
        assertFalse(font.isStrike());
        font.setStrike(value);
        assertEquals(value, font.isStrike());
    }

    @DisplayName("Test of the get and set function of the charset field")
    @ParameterizedTest(name = "Given value {1} should lead to the defined field")
    @CsvSource({
            "NULL, ''",
            "STRING, 'ASCII'",
    })
    void charsetTest(String sourceType, String sourceValue) {
        String value = (String) TestUtils.createInstance(sourceType, sourceValue);
        Font font = new Font();
        assertEquals("", font.getCharset());
        font.setCharset(value);
        assertEquals(value, font.getCharset());
    }

    @DisplayName("Test of the get and set function of the size field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource({
            "8",
            "75",
            "11",
    })
    void sizeTest(int value) {
        Font font = new Font();
        assertEquals(Font.DEFAULT_FONT_SIZE, font.getSize()); // 11 is default
        font.setSize(value);
        assertEquals(value, font.getSize());
    }

    @DisplayName("Test of the auto-adjusting set function of the size field (invalid values)")
    @ParameterizedTest(name = "Given value {0} should lead to the adjusted value {1}")
    @CsvSource({
            "0f, 1f",
            "7f, 7f",
            "-100f, 1f",
            "0.5f, 1f",
            "200f, 200f",
            "500f, 409f",
            "409.05f, 409f",
    })
    void sizeFailTest(float givenValue, float expectedValue) {
        Font font = new Font();
        font.setSize(givenValue);
        assertEquals(expectedValue, font.getSize());
    }

    @DisplayName("Test of the get and set function of the name field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource({
            "Calibri",
            "Arial",
            "---", // Not a font but a valid string
    })
    void nameTest(String value) {
        Font font = new Font();
        assertEquals(Font.DEFAULT_FONT_NAME, font.getName()); // Default is 'Calibri'
        font.setName(value);
        assertEquals(value, font.getName());
    }

    @DisplayName("Test of the failing set function of the name field")
    @Test()
    void nameFailTest() {
        Font font = new Font();
        assertThrows(StyleException.class, () -> font.setName(null));
        assertThrows(StyleException.class, () -> font.setName(""));
    }


    @DisplayName("Test of the get and set function of the family field")
    @ParameterizedTest(name = "Given value {1} should lead to the defined field")
    @CsvSource({
            "NULL, ''",
            "STRING, '4'",
    })
    void familyTest(String sourceType, String sourceValue) {
        String value = (String) TestUtils.createInstance(sourceType, sourceValue);
        Font font = new Font();
        assertEquals(Font.DEFAULT_FONT_FAMILY, font.getFamily());
        font.setFamily(value);
        assertEquals(value, font.getFamily());
    }

    @DisplayName("Test of the get and set function of the colorTheme field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource({
            "1",
            "10",
    })
    void colorThemeTest(int value) {
        Font font = new Font();
        assertEquals(1, font.getColorTheme()); // 1 is default
        font.setColorTheme(value);
        assertEquals(value, font.getColorTheme());
    }

    @DisplayName("Test of the failing set function of the colorTheme field (invalid values)")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource({
            "-1",
            "-100",
    })
    void colorThemeFailTest(int value) {
        Font font = new Font();
        assertThrows(Exception.class, () -> font.setColorTheme(value));
    }

    @DisplayName("Test of the get and set function of the colorValue field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined filed")
    @CsvSource({
            "STRING, ''",
            "NULL, ''",
            "STRING, 'FFAA22CC'",
    })
    void colorValueTest(String sourceType, String sourceValue) {
        String value = (String) TestUtils.createInstance(sourceType, sourceValue);
        Font font = new Font();
        assertEquals("", font.getColorValue()); // default is empty
        font.setColorValue(value);
        assertEquals(value, font.getColorValue());
    }

    @DisplayName("Test of the failing set function of the colorValue field (invalid values)")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource({
            "77BB00",
            "0002200000",
            "XXXXXXXX",
    })
    void colorValueFailTest(String value) {
        Font font = new Font();
        assertThrows(Exception.class, () -> font.setColorValue(value));
    }

    @DisplayName("Test of the get and set function of the scheme field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource({
            "major",
            "minor",
            "none",
    })
    void schmeTest(Font.SchemeValue value) {
        Font font = new Font();
        assertEquals(Font.DEFAULT_FONT_SCHEME, font.getScheme()); // default is minor
        font.setScheme(value);
        assertEquals(value, font.getScheme());
    }

    @DisplayName("Test of the get and set function of the verticalAlign field")
    @ParameterizedTest(name = "Given value {0} should lead to the defined field")
    @CsvSource({
            "none",
            "subscript",
            "superscript",
    })
    void verticalAlignTest(Font.VerticalAlignValue value) {
        Font font = new Font();
        assertEquals(Font.DEFAULT_VERTICAL_ALIGN, font.getVerticalAlign()); // default is none
        font.setVerticalAlign(value);
        assertEquals(value, font.getVerticalAlign());
    }

    @DisplayName("Test of the get function of the isDefaultFont field")
    @Test()
    void isDefaultFontTest() {
        Font font = new Font();
        assertTrue(font.isDefaultFont());
        font.setItalic(true);
        font.setName("XYZ");
        assertFalse(font.isDefaultFont());
    }

    @DisplayName("Test of the copy function")
    @Test()
    void copyTest() {
        Font copy = exampleStyle.copy();
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @DisplayName("Test of the equals method")
    @Test()
    void equalsTest() {
        Font style2 = (Font) exampleStyle.copy();
        assertTrue(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of bold)")
    @Test()
    void equalsTest2a() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setBold(false);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of italic)")
    @Test()
    void equalsTest2b() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setItalic(false);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of underline)")
    @Test()
    void equalsTest2c() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setUnderline(Font.UnderlineValue.singleAccounting);
        assertFalse(exampleStyle.equals(style2));
    }

    @DisplayName("Test of the equals method (inequality of strike)")
    @Test()
    void equalsTest2e() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setStrike(false);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of charset)")
    @Test()
    void equalsTest2f() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setCharset("XYZ");
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of size)")
    @Test()
    void equalsTest2g() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setSize(33);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of name)")
    @Test()
    void equalsTest2h() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setName("Comic Sans");
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of family)")
    @Test()
    void equalsTest2i() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setFamily("999");
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of colorTheme)")
    @Test()
    void equalsTest2j() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setColorTheme(22);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of colorValue)")
    @Test()
    void equalsTest2k() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setColorValue("FF9988AA");
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of scheme)")
    @Test()
    void equalsTest2l() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setScheme(Font.SchemeValue.none);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality of verticalAlign)")
    @Test()
    void equalsTest2m() {
        Font style2 = (Font) exampleStyle.copy();
        style2.setVerticalAlign(Font.VerticalAlignValue.none);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the equals method (inequality on null or different objects)")
    @ParameterizedTest(name = "Given value {1} should lead to inequality")
    @CsvSource({
            "NULL, ''",
            "STRING, 'text'",
            "BOOLEAN, 'true'",
    })
    void equalsTest3(String sourceType, String sourceValue) {
        Object obj = TestUtils.createInstance(sourceType, sourceValue);
        assertFalse(exampleStyle.equals(obj));
    }

    @DisplayName("Test of the equals method when the origin object is null or not of the same type")
    @ParameterizedTest(name = "Given value {1} should lead to inequality")
    @CsvSource({
            "NULL, ''",
            "STRING, 'origin'",
            "BOOLEAN, 'true'",
    })
    void equalsTest5(String sourceType, String sourceValue) {
        Object origin = TestUtils.createInstance(sourceType, sourceValue);
        Font copy = (Font) exampleStyle.copy();
        assertFalse(copy.equals(origin));
    }

    @DisplayName("Test of the hashCode method (equality of two identical objects)")
    @Test()
    void hashCodeTest() {
        Font copy = (Font) exampleStyle.copy();
        copy.setInternalID(99);  // Should not influence
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }


    @DisplayName("Test of the hashCode method (inequality of two different objects)")
    @Test()
    void hashCodeTest2() {
        Font copy = (Font) exampleStyle.copy();
        copy.setBold(false);
        assertNotEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @DisplayName("Test of the CompareTo method")
    @Test()
    void compareToTest() {
        Font font = new Font();
        Font other = new Font();
        font.setInternalID(null);
        other.setInternalID(null);
        assertEquals(-1, font.compareTo(other));
        font.setInternalID(5);
        assertEquals(1, font.compareTo(other));
        assertEquals(1, font.compareTo(null));
        other.setInternalID(5);
        assertEquals(0, font.compareTo(other));
        other.setInternalID(4);
        assertEquals(1, font.compareTo(other));
        other.setInternalID(6);
        assertEquals(-1, font.compareTo(other));
    }

    // For code coverage
    @DisplayName("Test of the toString function")
    @Test()
    void toStringTest() {
        Font font = new Font();
        String s1 = font.toString();
        font.setName("YXZ");
        assertNotEquals(s1, font.toString()); // An explicit value comparison is probably not sensible
    }

    @DisplayName("Test of the getValue function of the VerticalAlignValue enum")
    @Test()
    void verticalAlignValueTest() {
        assertEquals(0, Font.VerticalAlignValue.none.getValue());
        assertEquals(1, Font.VerticalAlignValue.subscript.getValue());
        assertEquals(2, Font.VerticalAlignValue.superscript.getValue());
    }


}
