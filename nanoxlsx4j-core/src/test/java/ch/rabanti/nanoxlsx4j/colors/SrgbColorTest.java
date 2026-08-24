package ch.rabanti.nanoxlsx4j.colors;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.internal.interfaces.BaseColor;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;

class SrgbColorTest {

    @ParameterizedTest
    @DisplayName("Test of the getter and setter of the ColorValue property on valid values")
    @CsvSource({
            "FFFFFF, FFFFFFFF",
            "000000, FF000000",
            "ABCDEF, FFABCDEF",
            "123456, FF123456",
            "abcdef, FFABCDEF",
            "ffaabb, FFFFAABB"
    })
    void colorValueTest(String givenSrgbValue, String expectedSrgbValue) {
        SrgbColor color = new SrgbColor();
        assertNull(color.getColorValue());
        color.setColorValue(givenSrgbValue);
        assertEquals(expectedSrgbValue, color.getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing getter and setter of the ColorValue property on invalid values")
    @NullSource
    @ValueSource(strings = {"-1", "0", "", "XABBCC", "AAAAA", "AAAAAAA", "AAAAAAAAA", "#AAAAAAAA",
            "01234", "#001122", "-aabbcc"})
    void colorValueFailTest(String srgbValue) {
        SrgbColor color = new SrgbColor();
        assertNull(color.getColorValue());
        assertThrows(StyleException.class, () -> color.setColorValue(srgbValue));
    }

    @ParameterizedTest
    @DisplayName("Test of the getter of the StringValue property on valid values")
    @CsvSource({
            "FFFFFF, FFFFFFFF",
            "000000, FF000000",
            "ABCDEF, FFABCDEF",
            "123456, FF123456",
            "abcdef, FFABCDEF",
            "ffaabb, FFFFAABB"
    })
    void stringValueTest(String givenSrgbValue, String expectedSrgbValue) {
        SrgbColor color = new SrgbColor();
        assertNull(color.getStringValue());
        color.setColorValue(givenSrgbValue);
        assertEquals(expectedSrgbValue, color.getStringValue());
    }

    @ParameterizedTest
    @DisplayName("Test of Constructor with arguments (ColorValue) on valid values")
    @CsvSource({
            "FFFFFF, FFFFFFFF",
            "000000, FF000000",
            "ABCDEF, FFABCDEF",
            "123456, FF123456",
            "abcdef, FFABCDEF",
            "ffaabb, FFFFAABB"
    })
    void constructorTest(String givenSrgbValue, String expectedSrgbValue) {
        SrgbColor color = new SrgbColor(givenSrgbValue);
        assertEquals(expectedSrgbValue, color.getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing constructor with arguments (ColorValue) on invalid values")
    @NullSource
    @ValueSource(strings = {"-1", "0", "", "XABBCC", "AAAAA", "AAAAAAA", "AAAAAAAAA", "#AAAAAAAA",
            "01234", "#001122", "-aabbcc"})
    void constructorFailTest(String srgbValue) {
        assertThrows(StyleException.class, () -> new SrgbColor(srgbValue));
    }

    @ParameterizedTest
    @DisplayName("Test of the ToArgbColor function")
    @CsvSource({
            "FFFFFF, FFFFFFFF",
            "000000, FF000000",
            "ABCDEF, FFABCDEF",
            "123456, FF123456",
            "abcdef, FFABCDEF",
            "ffaabb, FFFFAABB"
    })
    void toArgbColorTest(String srgbValue, String expectedArgbColor) {
        SrgbColor color = new SrgbColor(srgbValue);
        assertEquals(expectedArgbColor, color.getColorValue());
    }

    @Test
    @DisplayName("Test of the Equals method (multiple cases)")
    void equalsTest() {
        SrgbColor color1 = new SrgbColor("ACADAF");
        SrgbColor color2 = new SrgbColor();
        color2.setColorValue("ACADAF");
        assertTrue(color1.equals(color2));

        SrgbColor color3 = new SrgbColor();
        SrgbColor color4 = new SrgbColor();
        assertTrue(color3.equals(color4));
    }

    @Test
    @DisplayName("Test of the Equals method on inequality (multiple cases)")
    void equalsTest2() {
        SrgbColor color1 = new SrgbColor("ACADAF");
        SrgbColor color2 = new SrgbColor();
        color2.setColorValue("ACADA0");
        assertFalse(color1.equals(color2));

        SrgbColor color3 = new SrgbColor("ACADAF");
        SrgbColor color4 = new SrgbColor();
        assertFalse(color3.equals(color4));

        SrgbColor color5 = new SrgbColor();
        DummyColor color6 = new DummyColor();
        assertFalse(color5.equals(color6));
    }

    @Test
    @DisplayName("Test of the GetHashCode method (multiple cases)")
    void getHashCodeTest() {
        SrgbColor color1 = new SrgbColor("ACADAF");
        SrgbColor color2 = new SrgbColor();
        color2.setColorValue("ACADAF");
        assertEquals(color1.hashCode(), color2.hashCode());

        SrgbColor color3 = new SrgbColor();
        SrgbColor color4 = new SrgbColor();
        assertEquals(color3.hashCode(), color4.hashCode());
    }

    @Test
    @DisplayName("Test of the GetHashCode method on inequality (multiple cases)")
    void getHashCodeTest2() {
        SrgbColor color1 = new SrgbColor("ACADAF");
        SrgbColor color2 = new SrgbColor();
        color2.setColorValue("ACADA0");
        assertNotEquals(color1.hashCode(), color2.hashCode());

        SrgbColor color3 = new SrgbColor("ACADAF");
        SrgbColor color4 = new SrgbColor();
        assertNotEquals(color3.hashCode(), color4.hashCode());

        SrgbColor color5 = new SrgbColor();
        DummyColor color6 = new DummyColor();
        assertNotEquals(color5.hashCode(), color6.hashCode());
    }

    @Test
    @DisplayName("Test of the ToString method")
    void toStringTest() {
        SrgbColor color1 = new SrgbColor("ACADAF");
        assertEquals("FFACADAF", color1.toString());
        SrgbColor color2 = new SrgbColor();
        assertNull(color2.toString());
    }

    private static final class DummyColor implements BaseColor {
        @Override
        public String getStringValue() {
            return null;
        }

        @Override
        public int hashCode() {
            return 800285906;
        }
    }
}
