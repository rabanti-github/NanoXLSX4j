/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.themes;

import java.util.Objects;

import ch.rabanti.nanoxlsx4j.colors.SrgbColor;
import ch.rabanti.nanoxlsx4j.colors.SystemColor;

public class Theme {

// constants
    /**
     * Default theme ID, stated in the workbook document
     *
     * <p>
     * Remarks: According to the official OOXML documentation (part 1, chapter 18.2.28) the version consists of the
     * application version and build where the Excel file was created. The value was extracted from a valid Excel file,
     * created with Excel 2019. However, although '16' can be assumed to be the Version of Excel 2019, the build part
     * '6925' cannot be originated, is not reflecting the retrieved application build version, and seems not to be
     * listed publicly
     * </p>
     */
    public static final String DEFAULT_THEME_VERSION = "166925";

// enums

    public enum ColorSchemeElement {
        /// <summary>Dark 1</summary>
        DARK_1(0),
        /// <summary>Light 1</summary>
        LIGHT_1(1),
        /// <summary>Dark 2</summary>
        DARK_2(2),
        /// <summary>Light 2</summary>
        LIGHT_2(3),
        /// <summary>Accent 1</summary>
        ACCENT_1(4),
        /// <summary>Accent 2</summary>
        ACCENT_2(5),
        /// <summary>Accent 3</summary>
        ACCENT_3(6),
        /// <summary>Accent 4</summary>
        ACCENT_4(7),
        /// <summary>Accent 5</summary>
        ACCENT_5(8),
        /// <summary>Accent 6</summary>
        ACCENT_6(9),
        /// <summary>Hyperlink</summary>
        HYPERLINK(10),
        /// <summary>Followed Hyperlink</summary>
        FOLLOWED_HYPERLINK(11);

        public final int value;

        private ColorSchemeElement(int value) {
            this.value = value;
        }
    }

    private String name;
    private ColorScheme colors;
    private final boolean defaultTheme;

    /**
     * Gets or sets the name of the theme
     *
     * @return Name of the theme
     */
    public String getName() {
        return name;
    }

    /**
     * Sets or sets the name of the theme
     *
     * @param name Name of the theme
     */
    public void setName(String name) {
        this.name = name;
    }

    /**
     * Gets the {@link ColorScheme} of the theme
     *
     * @return Color scheme
     */
    public ColorScheme getColors() {
        return colors;
    }

    /**
     * Sets the {@link ColorScheme} of the theme
     *
     * @param colors Color scheme
     */
    public void setColors(ColorScheme colors) {
        this.colors = colors;
    }

    /**
     * Gets whether the theme is defined as copy or reference to the application default theme.
     * <p>
     * Remarks: This indication and the default theme ({@link Theme#getDefaultTheme()}) may still deviate from the
     * actual default
     * </p>
     *
     * @return True if the current instance is a default theme
     */
    public boolean isDefaultTheme() {
        return defaultTheme;
    }

    /**
     * Constructor with name. Using this constructor initialized the {@link Theme#getColors()} property with valid
     * default values
     *
     * @param name Name of the theme
     */
    public Theme(String name) {
        this(name, false);
    }

    /**
     * Internal constructor with parameters. Using this constructor initialized the {@link Theme#getColors()} property
     * with valid default values
     *
     * @param name         Name of the theme
     * @param defaultTheme If true, the theme is a default theme
     */
    Theme(String name, boolean defaultTheme) {
        this.name = name;
        this.defaultTheme = defaultTheme;
        this.colors = getDefaultColorScheme();
    }

    // TODO getDefaultTheme() was originally internal in C# -> Check accessibility

    /**
     * Gets the default theme if no theme was explicitly defined. This theme will be stored into an XLSX file if not
     * otherwise defined
     *
     * @return Theme with default values according to the default theme of Office 2019 (may be deviating)
     */
    private static Theme getDefaultTheme() {
        Theme theme = new Theme("default", true);
        ColorScheme colors = getDefaultColorScheme();
        theme.setColors(colors);
        return theme;
    }

    // TODO getDefaultTheme() was originally internal in C# -> Check accessibility

    /**
     * Gets the default color scheme if no scheme was explicitly defined. This theme will be incorporated into the
     * default theme of an XLSX file if not otherwise defined
     *
     * @return Color scheme with default values according to the default color scheme of Office 2019 (may be deviating)
     */
    private static ColorScheme getDefaultColorScheme() {
        ColorScheme colors = new ColorScheme();
        colors.setName("default");
        colors.setDark1(new SystemColor(SystemColor.Value.WINDOW_TEXT));
        colors.setDark1(new SystemColor(SystemColor.Value.WINDOW_TEXT));
        colors.setLight1(new SystemColor(SystemColor.Value.WINDOW, "FFFFFF"));
        colors.setDark2(new SrgbColor("44546A"));
        colors.setLight2(new SrgbColor("E7E6E6"));
        colors.setAccent1(new SrgbColor("4472C4"));
        colors.setAccent2(new SrgbColor("ED7D31"));
        colors.setAccent3(new SrgbColor("A5A5A5"));
        colors.setAccent4(new SrgbColor("FFC000"));
        colors.setAccent5(new SrgbColor("5B9BD5"));
        colors.setAccent6(new SrgbColor("70AD47"));
        colors.setHyperlink(new SrgbColor("0563C1"));
        colors.setFollowedHyperlink(new SrgbColor("954F72"));
        return colors;
    }

    /**
     * Returns whether two instances are the same
     *
     * @param o the reference object with which to compare.
     * @return True if this instance and the other are the same
     */
    @Override
    public final boolean equals(Object o) {
        if (!(o instanceof Theme theme)) {
            return false;
        }

        return Objects.equals(name, theme.name) && Objects.equals(colors, theme.colors);
    }

    /**
     * Returns a hash code for this instance.
     *
     * @return A hash code for this instance, suitable to be used in hashing algorithms and data structures like a hash
     * table.
     */
    @Override
    public int hashCode() {
        int result = Objects.hashCode(name);
        result = 31 * result + Objects.hashCode(colors);
        return result;
    }
}
