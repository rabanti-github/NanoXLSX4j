package ch.rabanti.nanoxlsx4j.colors;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.themes.Theme;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.ValueSource;

class ThemeColorTest {

    @Test
    @DisplayName("Default Constructor Test")
    void constructorTest() {
        ThemeColor themeColor = new ThemeColor();
        assertEquals(Theme.ColorSchemeElement.DARK_1, themeColor.getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Default Constructor Test with enum element as argument")
    @EnumSource(Theme.ColorSchemeElement.class)
    void constructorTest2(Theme.ColorSchemeElement colorSchemeElement) {
        ThemeColor themeColor = new ThemeColor(colorSchemeElement);
        assertEquals(colorSchemeElement, themeColor.getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Default Constructor Test with index as argument")
    @CsvSource({
            "0, DARK_1",
            "1, LIGHT_1",
            "2, DARK_2",
            "3, LIGHT_2",
            "4, ACCENT_1",
            "5, ACCENT_2",
            "6, ACCENT_3",
            "7, ACCENT_4",
            "8, ACCENT_5",
            "9, ACCENT_6",
            "10, HYPERLINK",
            "11, FOLLOWED_HYPERLINK"
    })
    void constructorTest3(int givenIndex, Theme.ColorSchemeElement expectedElement) {
        ThemeColor themeColor = new ThemeColor(givenIndex);
        assertEquals(expectedElement, themeColor.getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing Constructor on invalid values")
    @ValueSource(ints = {-1, 12, 255, -100})
    void constructorFailTest(int value) {
        assertThrows(StyleException.class, () -> new ThemeColor(value));
    }

    @ParameterizedTest
    @DisplayName("Test of the StringValue property")
    @CsvSource({
            "DARK_1, 0",
            "LIGHT_1, 1",
            "DARK_2, 2",
            "LIGHT_2, 3",
            "ACCENT_1, 4",
            "ACCENT_2, 5",
            "ACCENT_3, 6",
            "ACCENT_4, 7",
            "ACCENT_5, 8",
            "ACCENT_6, 9",
            "HYPERLINK, 10",
            "FOLLOWED_HYPERLINK, 11"
    })
    void stringValueTest(Theme.ColorSchemeElement colorSchemeElement, String expectedValue) {
        ThemeColor themeColor = new ThemeColor(colorSchemeElement);
        assertEquals(expectedValue, themeColor.getStringValue());
    }

    @Test
    @DisplayName("Test of the Equals method on equality (multiple cases)")
    void equalsTestTrue() {
        ThemeColor color1 = new ThemeColor(Theme.ColorSchemeElement.ACCENT_3);
        ThemeColor color2 = new ThemeColor(Theme.ColorSchemeElement.ACCENT_3);
        assertTrue(color1.equals(color2));

        ThemeColor color3 = new ThemeColor(Theme.ColorSchemeElement.DARK_1);
        ThemeColor color4 = new ThemeColor(Theme.ColorSchemeElement.DARK_1);
        assertTrue(color3.equals(color4));
    }

    @Test
    @DisplayName("Test of the Equals method on inequality (multiple cases)")
    void equalsTestFalse() {
        ThemeColor color1 = new ThemeColor(Theme.ColorSchemeElement.ACCENT_3);
        ThemeColor color2 = new ThemeColor(Theme.ColorSchemeElement.ACCENT_4);
        assertFalse(color1.equals(color2));

        ThemeColor color3 = new ThemeColor(Theme.ColorSchemeElement.DARK_1);
        ThemeColor color4 = new ThemeColor(Theme.ColorSchemeElement.LIGHT_1);
        assertFalse(color3.equals(color4));
    }

    @Test
    @DisplayName("Test of the GetHashCode method on equality (multiple cases)")
    void getHashCodeTestTrue() {
        ThemeColor color1 = new ThemeColor(Theme.ColorSchemeElement.ACCENT_3);
        ThemeColor color2 = new ThemeColor(Theme.ColorSchemeElement.ACCENT_3);
        assertEquals(color1.hashCode(), color2.hashCode());

        ThemeColor color3 = new ThemeColor(Theme.ColorSchemeElement.DARK_1);
        ThemeColor color4 = new ThemeColor(Theme.ColorSchemeElement.DARK_1);
        assertEquals(color3.hashCode(), color4.hashCode());
    }

    @Test
    @DisplayName("Test of the GetHashCode method on inequality (multiple cases)")
    void getHashCodeTestFalse() {
        ThemeColor color1 = new ThemeColor(Theme.ColorSchemeElement.ACCENT_3);
        ThemeColor color2 = new ThemeColor(Theme.ColorSchemeElement.DARK_1);
        assertNotEquals(color1.hashCode(), color2.hashCode());

        ThemeColor color3 = new ThemeColor(Theme.ColorSchemeElement.DARK_1);
        ThemeColor color4 = new ThemeColor(Theme.ColorSchemeElement.LIGHT_1);
        assertNotEquals(color3.hashCode(), color4.hashCode());
    }

    @Test
    @DisplayName("Test of the ToString method")
    void toStringTest() {
        ThemeColor color = new ThemeColor(Theme.ColorSchemeElement.ACCENT_3);
        assertEquals(color.getStringValue(), color.toString());
    }

}
