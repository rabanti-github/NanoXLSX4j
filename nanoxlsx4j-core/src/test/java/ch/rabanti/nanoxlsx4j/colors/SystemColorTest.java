package ch.rabanti.nanoxlsx4j.colors;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.internal.interfaces.BaseColor;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;

class SystemColorTest {

    @ParameterizedTest
    @DisplayName("Test of the getter and setter of the ColorValue property")
    @EnumSource(SystemColor.Value.class)
    void colorValueTest(SystemColor.Value value) {
        SystemColor color = new SystemColor();
        assertEquals(SystemColor.Value.WINDOW_TEXT, color.getColorValue());
        color.setColorValue(value);
        assertEquals(value, color.getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the getter of the StringValue property")
    @CsvSource({
            "ACTIVE_BORDER, activeBorder",
            "ACTIVE_CAPTION, activeCaption",
            "APP_WORKSPACE, appWorkspace",
            "BACKGROUND, background",
            "BUTTON_FACE, btnFace",
            "BUTTON_HIGHLIGHT, btnHighlight",
            "BUTTON_SHADOW, btnShadow",
            "BUTTON_TEXT, btnText",
            "CAPTION_TEXT, captionText",
            "GRADIENT_ACTIVE_CAPTION, gradientActiveCaption",
            "GRADIENT_INACTIVE_CAPTION, gradientInactiveCaption",
            "GRAY_TEXT, grayText",
            "HIGHLIGHT, highlight",
            "HIGHLIGHT_TEXT, highlightText",
            "HOT_LIGHT, hotLight",
            "INACTIVE_BORDER, inactiveBorder",
            "INACTIVE_CAPTION, inactiveCaption",
            "INACTIVE_CAPTION_TEXT, inactiveCaptionText",
            "INFO_BACKGROUND, infoBk",
            "INFO_TEXT, infoText",
            "MENU, menu",
            "MENU_BAR, menuBar",
            "MENU_HIGHLIGHT, menuHighlight",
            "MENU_TEXT, menuText",
            "SCROLL_BAR, scrollBar",
            "THREE_DIMENSIONAL_DARK_SHADOW, 3dDkShadow",
            "THREE_DIMENSIONAL_LIGHT, 3dLight",
            "WINDOW, window",
            "WINDOW_FRAME, windowFrame",
            "WINDOW_TEXT, windowText"
    })
    void stringValueTest(SystemColor.Value givenValue, String expectedValue) {
        SystemColor color = new SystemColor();
        assertEquals(SystemColor.Value.WINDOW_TEXT, color.getColorValue());
        color.setColorValue(givenValue);
        assertEquals(expectedValue, color.getStringValue());
    }

    // C# StringValueFailTest omitted: Java enums cannot represent an undefined numeric enum value.

    @ParameterizedTest
    @DisplayName("Test of the getter and setter of the LastColor property on valid values")
    @ValueSource(strings = {"FFFFFF", "000000", "ABCDEF", "123456", "abcdef", "ffaabb"})
    void lastColorTest(String srgbValue) {
        SystemColor color = new SystemColor();
        assertEquals("000000", color.getLastColor());
        color.setLastColor(srgbValue);
        assertEquals(srgbValue, color.getLastColor());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing getter and setter of the LastColor property on invalid values")
    @NullSource
    @ValueSource(strings = {"-1", "0", "", "XABBCC", "AAAAA", "AAAAAAA", "AAAAAAAA", "01234",
            "#001122", "-aabbcc"})
    void lastColorFailTest(String srgbValue) {
        SystemColor color = new SystemColor();
        assertEquals("000000", color.getLastColor());
        assertThrows(StyleException.class, () -> color.setLastColor(srgbValue));
    }

    @ParameterizedTest
    @DisplayName("Test of the constructor with the color value as argument")
    @CsvSource({
            "ACTIVE_BORDER, AABBCC",
            "ACTIVE_CAPTION, FFFFFF",
            "APP_WORKSPACE, 000000",
            "BACKGROUND, 999999",
            "BUTTON_FACE, A3F4C5",
            "BUTTON_HIGHLIGHT, aaaaaa",
            "BUTTON_SHADOW, ffffff",
            "BUTTON_TEXT, 012345",
            "CAPTION_TEXT, A9A9A9",
            "GRADIENT_ACTIVE_CAPTION, A1c4F9",
            "GRADIENT_INACTIVE_CAPTION, 000001",
            "GRAY_TEXT, 100000",
            "HIGHLIGHT, ABCDEF",
            "HIGHLIGHT_TEXT, aabbcc",
            "HOT_LIGHT, ffffff",
            "INACTIVE_BORDER, 010101",
            "INACTIVE_CAPTION, a4a4a4",
            "INACTIVE_CAPTION_TEXT, CCCCCC",
            "INFO_BACKGROUND, BbBbBb",
            "INFO_TEXT, 898900",
            "MENU, cccccc",
            "MENU_BAR, 0A0B0C",
            "MENU_HIGHLIGHT, 777777",
            "MENU_TEXT, 70A9f7",
            "SCROLL_BAR, 4cff33",
            "THREE_DIMENSIONAL_DARK_SHADOW, 00000A",
            "THREE_DIMENSIONAL_LIGHT, FFFFFE",
            "WINDOW, eeeeef",
            "WINDOW_FRAME, 65CC78",
            "WINDOW_TEXT, AD44FF"
    })
    void constructorTest2(SystemColor.Value value, String lastColor) {
        SystemColor color = new SystemColor(value, lastColor);
        assertEquals(value, color.getColorValue());
        assertEquals(lastColor, color.getLastColor());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing constructor with arguments (color value and last color) on invalid values")
    @CsvSource(value = {
            "INACTIVE_CAPTION_TEXT, -1",
            "WINDOW_TEXT, 0",
            "GRAY_TEXT, ''",
            "ACTIVE_BORDER, NULL",
            "HIGHLIGHT_TEXT, XABBCC",
            "BACKGROUND, AAAAA",
            "BUTTON_SHADOW, AAAAAAA",
            "CAPTION_TEXT, AAAAAAAA",
            "BUTTON_HIGHLIGHT, 01234",
            "ACTIVE_CAPTION, #001122",
            "BUTTON_FACE, -aabbcc"
    }, nullValues = "NULL")
    void constructorFailTest(SystemColor.Value value, String srgbValue) {
        assertThrows(StyleException.class, () -> new SystemColor(value, srgbValue));
    }

    @Test
    @DisplayName("Test of the Equals method (multiple cases)")
    void equalsTest() {
        SystemColor color1 = new SystemColor(SystemColor.Value.BUTTON_HIGHLIGHT);
        color1.setLastColor("112233");
        SystemColor color2 = new SystemColor();
        color2.setColorValue(SystemColor.Value.BUTTON_HIGHLIGHT);
        color2.setLastColor("112233");
        assertTrue(color1.equals(color2));

        SystemColor color3 = new SystemColor();
        SystemColor color4 = new SystemColor();
        assertTrue(color3.equals(color4));
    }

    @Test
    @DisplayName("Test of the Equals method on inequality (multiple cases)")
    void equalsTest2() {
        SystemColor color1 = new SystemColor(SystemColor.Value.CAPTION_TEXT);
        SystemColor color2 = new SystemColor();
        color2.setColorValue(SystemColor.Value.GRADIENT_ACTIVE_CAPTION);
        assertFalse(color1.equals(color2));

        SystemColor color3 = new SystemColor(SystemColor.Value.ACTIVE_CAPTION);
        SystemColor color4 = new SystemColor();
        assertFalse(color3.equals(color4));

        SystemColor color5 = new SystemColor();
        DummyColor color6 = new DummyColor();
        assertFalse(color5.equals(color6));

        SystemColor color7 = new SystemColor(SystemColor.Value.CAPTION_TEXT, "AABBCC");
        SystemColor color8 = new SystemColor(SystemColor.Value.CAPTION_TEXT, "001122");
        assertFalse(color7.equals(color8));
    }

    @Test
    @DisplayName("Test of the GetHashCode method (multiple cases)")
    void getHashCodeTest() {
        SystemColor color1 = new SystemColor(SystemColor.Value.APP_WORKSPACE);
        SystemColor color2 = new SystemColor();
        color2.setColorValue(SystemColor.Value.APP_WORKSPACE);
        assertEquals(color1.hashCode(), color2.hashCode());

        SystemColor color3 = new SystemColor();
        SystemColor color4 = new SystemColor();
        assertEquals(color3.hashCode(), color4.hashCode());

        SystemColor color5 = new SystemColor(SystemColor.Value.APP_WORKSPACE, "CCDDEE");
        SystemColor color6 = new SystemColor();
        color6.setColorValue(SystemColor.Value.APP_WORKSPACE);
        color6.setLastColor("CCDDEE");
        assertEquals(color5.hashCode(), color6.hashCode());
    }

    @Test
    @DisplayName("Test of the GetHashCode method on inequality (multiple cases)")
    void getHashCodeTest2() {
        SystemColor color1 = new SystemColor(SystemColor.Value.BACKGROUND);
        SystemColor color2 = new SystemColor();
        color2.setColorValue(SystemColor.Value.BUTTON_FACE);
        assertNotEquals(color1.hashCode(), color2.hashCode());

        SystemColor color3 = new SystemColor(SystemColor.Value.APP_WORKSPACE);
        SystemColor color4 = new SystemColor();
        assertNotEquals(color3.hashCode(), color4.hashCode());

        SystemColor color5 = new SystemColor();
        DummyColor color6 = new DummyColor();
        assertNotEquals(color5.hashCode(), color6.hashCode());

        SystemColor color7 = new SystemColor(SystemColor.Value.BACKGROUND, "AACCDD");
        SystemColor color8 = new SystemColor();
        color8.setColorValue(SystemColor.Value.BACKGROUND);
        color8.setLastColor("002233");
        assertNotEquals(color7.hashCode(), color8.hashCode());
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
