package ch.rabanti.nanoxlsx4j.colors;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Locale;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.themes.Theme;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;

class ColorTest {

    @Test
    @DisplayName("Test of the CreateNone function")
    void createNoneTest() {
        Color color = Color.createNone();
        assertEquals(Color.ColorType.NONE, color.getType());
        assertFalse(color.isDefined());
        assertNull(color.getValue());
    }

    @Test
    @DisplayName("Test of the CreateAuto function")
    void createAutoTest() {
        Color color = Color.createAuto();
        assertEquals(Color.ColorType.AUTO, color.getType());
        assertTrue(color.isAuto());
        assertTrue(color.isDefined());
        assertNotNull(color.getValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the CreateRgb function")
    @CsvSource({
            "000000, FF000000",
            "FFFFFF, FFFFFFFF",
            "AABBCC, FFAABBCC",
            "FF000000, FF000000",
            "FFFFFFFF, FFFFFFFF",
            "FFAABBCC, FFAABBCC"
    })
    void createRgbFromStringTest(String givenRgb, String expectedRgb) {
        Color color = Color.createRgb(givenRgb);
        assertEquals(Color.ColorType.RGB, color.getType());
        assertEquals(expectedRgb.toUpperCase(Locale.ROOT), color.getArgbValue().toUpperCase(Locale.ROOT));
    }

    @ParameterizedTest
    @DisplayName("Test of the CreateRgb function, using a SrgbColor instance")
    @CsvSource({
            "000000, FF000000",
            "FFFFFF, FFFFFFFF",
            "AABBCC, FFAABBCC",
            "FF000000, FF000000",
            "FFFFFFFF, FFFFFFFF",
            "FFAABBCC, FFAABBCC"
    })
    void createRgbFromStringTest2(String givenRgb, String expectedRgb) {
        SrgbColor color = new SrgbColor(givenRgb);
        Color compoundColor = Color.createRgb(color);
        assertEquals(Color.ColorType.RGB, compoundColor.getType());
        assertEquals(expectedRgb.toUpperCase(Locale.ROOT), compoundColor.getArgbValue().toUpperCase(Locale.ROOT));
    }

    @ParameterizedTest
    @DisplayName("Test of the failing CreateRgb function")
    @NullSource
    @ValueSource(strings = {"", "XYZ", "FFAABBCCDD", "FFAAB"})
    void createRgbFromStringFailureTest(String rgb) {
        assertThrows(StyleException.class, () -> Color.createRgb(rgb));
    }

    @ParameterizedTest
    @DisplayName("Test of the CreateIndexed function")
    @ValueSource(ints = {0, 8, 64})
    void createIndexedTest(int index) {
        Color color = Color.createIndexed(index);
        assertEquals(Color.ColorType.INDEXED, color.getType());
        assertNotNull(color.getIndexedColor());
        assertNotNull(color.getArgbValue());
        assertEquals(index, color.getIndexedColor().getColorValue().value);
    }

    @ParameterizedTest
    @DisplayName("Test of the CreateIndexed function, using a IndexedColor instance")
    @ValueSource(ints = {0, 8, 64})
    void createIndexedTest2(int index) {
        IndexedColor indexedColor = new IndexedColor(index);
        Color color = Color.createIndexed(indexedColor);
        assertEquals(Color.ColorType.INDEXED, color.getType());
        assertNotNull(color.getIndexedColor());
        assertNotNull(color.getArgbValue());
        assertEquals(index, color.getIndexedColor().getColorValue().value);
    }

    @ParameterizedTest
    @DisplayName("Test of the CreateIndexed function, using a IndexedColor enum value")
    @EnumSource(value = IndexedColor.Value.class,
            names = {"BLACK", "STRONG_MAGENTA", "SYSTEM_BACKGROUND", "SYSTEM_FOREGROUND"})
    void createIndexedTest3(IndexedColor.Value value) {
        IndexedColor indexedColor = new IndexedColor(value);
        Color color = Color.createIndexed(indexedColor);
        assertEquals(Color.ColorType.INDEXED, color.getType());
        assertNotNull(color.getIndexedColor());
        assertNotNull(color.getArgbValue());
        assertEquals(value, color.getIndexedColor().getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing CreateIndexed function")
    @ValueSource(ints = {-1, 66})
    void createIndexedFailureTest(int index) {
        assertThrows(StyleException.class, () -> Color.createIndexed(index));
    }

    @Test
    @DisplayName("Test of the failing CreateIndexed function when passing null")
    void createIndexedFailureTest2() {
        assertThrows(StyleException.class, () -> Color.createIndexed((IndexedColor) null));
    }

    @ParameterizedTest
    @DisplayName("Test of the CreateTheme function")
    @EnumSource(value = Theme.ColorSchemeElement.class,
            names = {"ACCENT_1", "DARK_1", "FOLLOWED_HYPERLINK", "LIGHT_1"})
    void createThemeTest(Theme.ColorSchemeElement value) {
        Color color = Color.createTheme(value, 0.25);
        assertEquals(Color.ColorType.THEME, color.getType());
        assertEquals(0.25, color.getTint());
        assertNull(color.getArgbValue());
        assertEquals(value, color.getThemeColor().getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the CreateTheme function, using a ThemeColor instance")
    @EnumSource(value = Theme.ColorSchemeElement.class,
            names = {"ACCENT_1", "DARK_1", "FOLLOWED_HYPERLINK", "LIGHT_1"})
    void createThemeTest2(Theme.ColorSchemeElement value) {
        ThemeColor themeColor = new ThemeColor(value);
        Color color = Color.createTheme(themeColor, -0.25);
        assertEquals(Color.ColorType.THEME, color.getType());
        assertEquals(-0.25, color.getTint());
        assertNull(color.getArgbValue());
        assertEquals(value, color.getThemeColor().getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the CreateTheme function, using an index")
    @CsvSource({
            "4, ACCENT_1",
            "0, DARK_1",
            "11, FOLLOWED_HYPERLINK",
            "1, LIGHT_1"
    })
    void createThemeTest3(int givenIndex, Theme.ColorSchemeElement expectedValue) {
        Color color = Color.createTheme(givenIndex, -0.25);
        assertEquals(Color.ColorType.THEME, color.getType());
        assertEquals(-0.25, color.getTint());
        assertNull(color.getArgbValue());
        assertEquals(expectedValue, color.getThemeColor().getColorValue());
    }

    @Test
    @DisplayName("Test of the failing CreateTheme function")
    void createThemeFailureTest() {
        assertThrows(StyleException.class, () -> Color.createTheme((ThemeColor) null));
    }

    @ParameterizedTest
    @DisplayName("Test of the CreateSystem function")
    @EnumSource(value = SystemColor.Value.class, names = {"ACTIVE_BORDER", "BUTTON_TEXT", "HIGHLIGHT", "WINDOW"})
    void createSystemTest(SystemColor.Value value) {
        Color color = Color.createSystem(value);
        assertEquals(Color.ColorType.SYSTEM, color.getType());
        assertNull(color.getTint());
        assertNull(color.getArgbValue());
        assertEquals(value, color.getSystemColor().getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the CreateSystem function, using a SystemColor instance")
    @EnumSource(value = SystemColor.Value.class, names = {"ACTIVE_BORDER", "BUTTON_TEXT", "HIGHLIGHT", "WINDOW"})
    void createSystemTest2(SystemColor.Value value) {
        SystemColor systemColor = new SystemColor(value);
        Color color = Color.createSystem(systemColor);
        assertEquals(Color.ColorType.SYSTEM, color.getType());
        assertNull(color.getTint());
        assertNull(color.getArgbValue());
        assertEquals(value, color.getSystemColor().getColorValue());
    }

    @Test
    @DisplayName("Test of the failing CreateSystem function")
    void createSystemFailureTest() {
        assertThrows(StyleException.class, () -> Color.createSystem((SystemColor) null));
    }

    // C# implicit conversion tests omitted: Java does not support operator overloading.

    @Test
    @DisplayName("Test of the Value property on None")
    void valueNoneTest() {
        Color color = Color.createNone();
        assertNull(color.getValue());
    }

    @Test
    @DisplayName("Test of the Value property on Auto")
    void valueAutoTest() {
        Color color = Color.createAuto();
        assertInstanceOf(AutoColor.class, color.getValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the Value property on sRGB")
    @CsvSource({
            "000000, FF000000",
            "FFFFFF, FFFFFFFF",
            "123456, FF123456",
            "FF000000, FF000000",
            "FFFFFFFF, FFFFFFFF",
            "FF234567, FF234567"
    })
    void valueSrgbTest(String givenRgbValue, String expectedRgbValue) {
        Color color = Color.createRgb(givenRgbValue);
        assertInstanceOf(SrgbColor.class, color.getValue());
        assertEquals(expectedRgbValue, color.getRgbColor().getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the Value property on indexed colors")
    @EnumSource(value = IndexedColor.Value.class,
            names = {"BLACK", "STRONG_MAGENTA", "SYSTEM_BACKGROUND", "SYSTEM_FOREGROUND"})
    void valueIndexedTest(IndexedColor.Value indexedValue) {
        Color color = Color.createIndexed(indexedValue);
        assertInstanceOf(IndexedColor.class, color.getValue());
        assertEquals(indexedValue, color.getIndexedColor().getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the Value property on system colors")
    @EnumSource(value = SystemColor.Value.class, names = {"ACTIVE_BORDER", "BUTTON_TEXT", "HIGHLIGHT", "WINDOW"})
    void valueSystemTest(SystemColor.Value systemColor) {
        Color color = Color.createSystem(systemColor);
        assertInstanceOf(SystemColor.class, color.getValue());
        assertEquals(systemColor, color.getSystemColor().getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the Value property on theme colors")
    @EnumSource(value = Theme.ColorSchemeElement.class,
            names = {"ACCENT_1", "DARK_1", "FOLLOWED_HYPERLINK", "LIGHT_1"})
    void valueThemeTest(Theme.ColorSchemeElement themeElement) {
        Color color = Color.createTheme(themeElement);
        ThemeColor value = assertInstanceOf(ThemeColor.class, color.getValue());
        assertEquals(themeElement, value.getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the GetArgbValue function on a sRGB color")
    @CsvSource({
            "000000, FF000000",
            "FFFFFF, FFFFFFFF",
            "123456, FF123456",
            "FF000000, FF000000",
            "FFFFFFFF, FFFFFFFF",
            "FF234567, FF234567"
    })
    void getArgbValueSrgbTest(String givenRgb, String expectedRgb) {
        Color color = Color.createRgb(givenRgb);
        assertEquals(expectedRgb, color.getArgbValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the GetArgbValue function on a sRGB color")
    @CsvSource({
            "BLACK_0, FF000000",
            "BLACK, FF000000",
            "WHITE, FFFFFFFF",
            "STRONG_CYAN, FF00FFFF",
            "DARK_MAROON, FF800000",
            "LAVENDER, FFCC99FF"
    })
    void getArgbValueIndexedTest(IndexedColor.Value givenIndex, String expectedRgb) {
        Color color = Color.createIndexed(givenIndex);
        assertEquals(expectedRgb, color.getArgbValue());
    }

    @Test
    @DisplayName("Test of the GetArgbValue function on a theme color")
    void getArgbValueReturnsNullForThemeTest() {
        Color color = Color.createTheme(Theme.ColorSchemeElement.DARK_1);
        assertNull(color.getArgbValue());
    }

    @Test
    @DisplayName("Test of the GetArgbValue function on a system color")
    void getArgbValueReturnsNullForSystemTest() {
        Color color = Color.createSystem(SystemColor.Value.ACTIVE_BORDER);
        assertNull(color.getArgbValue());
    }

    @Test
    @DisplayName("Test of the GetArgbValue function on a auto color")
    void getArgbValueReturnsNullForAutoTest() {
        Color color = Color.createAuto();
        assertNull(color.getArgbValue());
    }

    @Test
    @DisplayName("Test of the Equals method on equality")
    void equalsSameRgbValueTest() {
        Color a = Color.createRgb("FFABCDEF");
        Color b = Color.createRgb("FFABCDEF");
        assertEquals(a, b);
        assertTrue(a.equals(b));
    }

    @Test
    @DisplayName("Test of the Equals method on inequality")
    void equalsDifferentRgbValueTest() {
        Color a = Color.createRgb("FFABCDEF");
        Color b = Color.createRgb("FFABCDEE");
        assertNotEquals(a, b);
    }

    @Test
    @DisplayName("Test of the Equals method on inequality on different types")
    void equalsDifferentTypeTest() {
        Color a = Color.createRgb("FF000000");
        Color b = Color.createIndexed(0);
        assertNotEquals(a, b);
    }

    @Test
    @DisplayName("Test of the GetHasCode method on equality")
    void getHashCodeEqualObjectsTest() {
        Color a = Color.createRgb("FF112233");
        Color b = Color.createRgb("FF112233");
        assertEquals(a.hashCode(), b.hashCode());
    }

    @Test
    @DisplayName("Test of the GetHasCode method on inequality")
    void getHashCodeDifferentObjectsTest() {
        Color a = Color.createRgb("FF112233");
        Color b = Color.createRgb("FF332211");
        assertNotEquals(a.hashCode(), b.hashCode());
    }

    @Test
    @DisplayName("Test of the CompareTo method on null values")
    void compareToNullTest() {
        Color color = Color.createRgb("FF000000");
        assertTrue(color.compareTo(null) > 0);
    }

    // C# CompareToWrongTypeTest omitted: Comparable<Color> prevents wrong-type calls in Java.

    @Test
    @DisplayName("Test of the CompareTo method on two none color types")
    void compareNoneColorTypeTest() {
        Color a = Color.createNone();
        Color b = Color.createNone();
        assertEquals(0, a.compareTo(b));
    }

    @Test
    @DisplayName("Test of the CompareTo method on two auto color types")
    void compareAutoColorTypeTest() {
        Color a = Color.createAuto();
        Color b = Color.createAuto();
        assertEquals(0, a.compareTo(b));
    }

    @ParameterizedTest
    @DisplayName("Test of the CompareTo method on identical RGB/ARGB values")
    @ValueSource(strings = {"000000", "FFFFFF", "AABBCC", "FF000000", "FFFFFFFF", "FFAABBCC"})
    void compareToSameRgbTest(String rgbValue) {
        Color a = Color.createRgb(rgbValue);
        Color b = Color.createRgb(rgbValue);
        assertEquals(0, a.compareTo(b));
    }

    @Test
    @DisplayName("Test of the CompareTo method on different sRGB values")
    void compareToRgbOrderingTest() {
        Color a = Color.createRgb("FF000000");
        Color b = Color.createRgb("FFFFFFFF");
        assertTrue(a.compareTo(b) < 0);
    }

    @Test
    @DisplayName("Test of the CompareTo method on different color values if sRGB and indexes are compared")
    void compareToDifferentTypeOrderingTest() {
        Color rgb = Color.createRgb("FF000000");
        Color indexed = Color.createIndexed(0);
        assertNotEquals(0, rgb.compareTo(indexed));
    }

    @Test
    @DisplayName("Test of the CompareTo method on different tint values")
    void compareToThemeTintTest() {
        Color a = Color.createTheme(Theme.ColorSchemeElement.ACCENT_1, 0.1);
        Color b = Color.createTheme(Theme.ColorSchemeElement.ACCENT_1, 0.2);
        assertTrue(a.compareTo(b) < 0);
    }

    @Test
    @DisplayName("Test of the CompareTo method on colors with different theme slots")
    void compareToThemeDifferentThemeSlots() {
        Color color1 = Color.createTheme(Theme.ColorSchemeElement.DARK_1);
        Color color2 = Color.createTheme(Theme.ColorSchemeElement.ACCENT_1);
        assertTrue(color1.compareTo(color2) < 0);
    }

    @Test
    @DisplayName("Test of the CompareTo method on colors with same slot but different tint")
    void compareToThemeSameSlotDifferentTint() {
        Color color1 = Color.createTheme(Theme.ColorSchemeElement.ACCENT_1, -0.2);
        Color color2 = Color.createTheme(Theme.ColorSchemeElement.ACCENT_1, 0.2);
        assertTrue(color1.compareTo(color2) < 0);
    }

    @Test
    @DisplayName("Test of the CompareTo method on System colors")
    void compareToSystemColors() {
        Color color1 = Color.createSystem(new SystemColor(SystemColor.Value.APP_WORKSPACE));
        Color color2 = Color.createSystem(new SystemColor(SystemColor.Value.MENU));
        assertNotEquals(0, color1.compareTo(color2));
    }

    // C# CompareToDefensiveFallback omitted: Java enums cannot represent an undefined numeric enum value.

    @Test
    @DisplayName("Test of the CompareTo method on indexed colors uses numeric index")
    void compareToIndexedNumericComparison() {
        Color color1 = Color.createIndexed(IndexedColor.Value.BLACK);
        Color color2 = Color.createIndexed(IndexedColor.Value.WHITE);
        assertTrue(color1.compareTo(color2) < 0);
    }

    @ParameterizedTest
    @DisplayName("Test of the ToStringFunction (for code coverage)")
    @EnumSource(Color.ColorType.class)
    void toStringTest(Color.ColorType type) {
        Color color;
        String expectedToken;
        switch (type) {
            case RGB -> {
                color = Color.createRgb("FFAABB");
                expectedToken = "FFAABB";
            }
            case INDEXED -> {
                color = Color.createIndexed(IndexedColor.Value.ROSE);
                expectedToken = Integer.toString(IndexedColor.Value.ROSE.value);
            }
            case THEME -> {
                color = Color.createTheme(Theme.ColorSchemeElement.ACCENT_5);
                expectedToken = Integer.toString(Theme.ColorSchemeElement.ACCENT_5.value);
            }
            case SYSTEM -> {
                color = Color.createSystem(SystemColor.Value.BACKGROUND);
                expectedToken = "Background";
            }
            case AUTO -> {
                color = Color.createAuto();
                expectedToken = "Auto";
            }
            default -> {
                color = Color.createNone();
                expectedToken = "Undefined";
            }
        }
        String given = color.toString().toLowerCase(Locale.ROOT);
        assertTrue(given.contains(expectedToken.toLowerCase(Locale.ROOT)));
    }

}
