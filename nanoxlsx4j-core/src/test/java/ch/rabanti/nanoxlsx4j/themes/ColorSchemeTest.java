package ch.rabanti.nanoxlsx4j.themes;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import ch.rabanti.nanoxlsx4j.colors.SrgbColor;
import ch.rabanti.nanoxlsx4j.colors.SystemColor;
import ch.rabanti.nanoxlsx4j.internal.interfaces.BaseColor;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;

class ColorSchemeTest {

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Name property")
    @NullSource
    @ValueSource(strings = {"XYZ", " ", ""})
    void nameTest(String value) {
        ColorScheme scheme = new ColorScheme();
        assertNull(scheme.getName());
        scheme.setName(value);
        assertEquals(value, scheme.getName());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the color properties for SRGB values")
    @ValueSource(strings = {"FFFFFF", "000000", "AACC3D"})
    void colorPropertiesSrgbTest(String value) {
        SrgbColor color = new SrgbColor(value);
        assertColorProperties(color);
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the color properties for SystemColor values")
    @EnumSource(SystemColor.Value.class)
    void colorPropertiesSystemColorTest(SystemColor.Value value) {
        SystemColor color = new SystemColor(value);
        assertColorProperties(color);
    }

    @Test
    @DisplayName("Test of the get and set function of the color properties for null values")
    void propertiesNullTest() {
        assertColorProperties(null);
    }

    @Test
    @DisplayName("Test of Equals() and HashCode() implementations for equality")
    void equalsTest() {
        ColorScheme scheme1 = new ColorScheme();
        scheme1.setName("scheme1");
        scheme1.setDark1(new SystemColor(SystemColor.Value.ACTIVE_BORDER));
        scheme1.setLight1(new SystemColor(SystemColor.Value.MENU));
        scheme1.setDark2(new SystemColor(SystemColor.Value.BACKGROUND));
        scheme1.setLight2(new SystemColor(SystemColor.Value.BACKGROUND));
        scheme1.setAccent1(new SystemColor(SystemColor.Value.APP_WORKSPACE));
        scheme1.setAccent2(new SystemColor(SystemColor.Value.BUTTON_SHADOW));
        scheme1.setAccent3(new SrgbColor("FFAABB"));
        scheme1.setAccent4(null);
        scheme1.setAccent5(new SrgbColor("FFAABB"));
        scheme1.setAccent6(new SrgbColor("FFAABB"));
        scheme1.setHyperlink(new SrgbColor("FFAABB"));
        scheme1.setFollowedHyperlink(new SrgbColor("FFAABB"));

        ColorScheme scheme2 = new ColorScheme();
        scheme2.setName("scheme1");
        scheme2.setDark1(new SystemColor(SystemColor.Value.ACTIVE_BORDER));
        scheme2.setLight1(new SystemColor(SystemColor.Value.MENU));
        scheme2.setDark2(new SystemColor(SystemColor.Value.BACKGROUND));
        scheme2.setLight2(new SystemColor(SystemColor.Value.BACKGROUND));
        scheme2.setAccent1(new SystemColor(SystemColor.Value.APP_WORKSPACE));
        scheme2.setAccent2(new SystemColor(SystemColor.Value.BUTTON_SHADOW));
        scheme2.setAccent3(new SrgbColor("FFAABB"));
        scheme2.setAccent4(null);
        scheme2.setAccent5(new SrgbColor("FFAABB"));
        scheme2.setAccent6(new SrgbColor("FFAABB"));
        scheme2.setHyperlink(new SrgbColor("FFAABB"));
        scheme2.setFollowedHyperlink(new SrgbColor("FFAABB"));

        assertTrue(scheme1.equals(scheme2));
        assertEquals(scheme1.hashCode(), scheme2.hashCode());
    }

    @Test
    @DisplayName("Test Equals method for Theme")
    void themeEqualsTest() {
        Theme theme1 = new Theme("TestTheme");
        Theme theme2 = new Theme("TestTheme");

        BaseColor newDark1 = new SrgbColor("123456");
        BaseColor newAccent1 = new SrgbColor("654321");
        BaseColor newHyperlink = new SrgbColor("ABC123");

        theme1.getColors().setDark1(newDark1);
        theme1.getColors().setAccent1(newAccent1);
        theme1.getColors().setHyperlink(newHyperlink);

        theme2.getColors().setDark1(newDark1);
        theme2.getColors().setAccent1(newAccent1);
        theme2.getColors().setHyperlink(newHyperlink);

        assertTrue(theme1.equals(theme2));
    }

    @Test
    @DisplayName("Test GetHashCode method for Theme")
    void themeGetHashCodeTest() {
        Theme theme1 = new Theme("TestTheme");
        Theme theme2 = new Theme("TestTheme");

        BaseColor newDark1 = new SrgbColor("123456");
        BaseColor newAccent1 = new SrgbColor("654321");
        BaseColor newHyperlink = new SrgbColor("ABC123");

        theme1.getColors().setDark1(newDark1);
        theme1.getColors().setAccent1(newAccent1);
        theme1.getColors().setHyperlink(newHyperlink);

        theme2.getColors().setDark1(newDark1);
        theme2.getColors().setAccent1(newAccent1);
        theme2.getColors().setHyperlink(newHyperlink);

        assertEquals(theme1.hashCode(), theme2.hashCode());
    }

    private static void assertColorProperties(BaseColor color) {
        ColorScheme scheme = new ColorScheme();
        assertNull(scheme.getDark1());
        assertNull(scheme.getLight1());
        assertNull(scheme.getDark2());
        assertNull(scheme.getLight2());
        assertNull(scheme.getAccent1());
        assertNull(scheme.getAccent2());
        assertNull(scheme.getAccent3());
        assertNull(scheme.getAccent4());
        assertNull(scheme.getAccent5());
        assertNull(scheme.getAccent6());
        assertNull(scheme.getHyperlink());
        assertNull(scheme.getFollowedHyperlink());

        scheme.setDark1(color);
        scheme.setLight1(color);
        scheme.setDark2(color);
        scheme.setLight2(color);
        scheme.setAccent1(color);
        scheme.setAccent2(color);
        scheme.setAccent3(color);
        scheme.setAccent4(color);
        scheme.setAccent5(color);
        scheme.setAccent6(color);
        scheme.setHyperlink(color);
        scheme.setFollowedHyperlink(color);

        assertEquals(color, scheme.getDark1());
        assertEquals(color, scheme.getLight1());
        assertEquals(color, scheme.getDark2());
        assertEquals(color, scheme.getLight2());
        assertEquals(color, scheme.getAccent1());
        assertEquals(color, scheme.getAccent2());
        assertEquals(color, scheme.getAccent3());
        assertEquals(color, scheme.getAccent4());
        assertEquals(color, scheme.getAccent5());
        assertEquals(color, scheme.getAccent6());
        assertEquals(color, scheme.getHyperlink());
        assertEquals(color, scheme.getFollowedHyperlink());
    }
}
