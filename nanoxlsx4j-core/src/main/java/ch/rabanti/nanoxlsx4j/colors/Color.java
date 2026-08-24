/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.colors;

import java.util.Comparator;
import java.util.Objects;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.internal.interfaces.BaseColor;
import ch.rabanti.nanoxlsx4j.themes.Theme;

/**
 * Compound class representing a color in various representations (RGB, indexed, theme, system or automatic)
 */
public final class Color implements Comparable<Color> {

    /**
     * Enum defining the type of color representation.
     */
    public enum ColorType {
        /** No color defined. */
        NONE,
        /** Automatic color (determined by application). */
        AUTO,
        /** RGB/ARGB color value. */
        RGB,
        /** Legacy indexed color (0-56+). */
        INDEXED,
        /** Theme color reference (0-11: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink). */
        THEME,
        /** System color (used in themes). */
        SYSTEM
    }

    private ColorType type;
    private boolean auto;
    private SrgbColor rgbColor;
    private IndexedColor indexedColor;
    private ThemeColor themeColor;
    private SystemColor systemColor;
    private Double tint;

    /**
     * Gets the type of color this value represents
     *
     * @return Color type
     */
    public ColorType getType() {
        return type;
    }

    /**
     * Gets the auto attribute - if true, color is automatically determined
     *
     * @return Automatic color if true
     */
    public boolean isAuto() {
        return auto;
    }

    /**
     * Gets the RGB/ARGB value when Type is RGB
     *
     * @return Color value
     */
    public SrgbColor getRgbColor() {
        return rgbColor;
    }

    /**
     * Gets the indexed color when Type is Indexed (See {@link IndexedColor.Value})
     *
     * @return Color value
     */
    public IndexedColor getIndexedColor() {
        return indexedColor;
    }

    /**
     * Gets the theme-based color when Type is Theme (See {@link Theme.ColorSchemeElement})
     *
     * @return Color value
     */
    public ThemeColor getThemeColor() {
        return themeColor;
    }

    /**
     * Gets the system color when Type is System (See {@link SystemColor.Value})
     *
     * @return Color value
     */
    public SystemColor getSystemColor() {
        return systemColor;
    }

    /**
     * Gets the optional tint value for colors (-1.0 to 1.0), mainly for theme colors
     *
     * @return Tint value
     */
    public Double getTint() {
        return tint;
    }

    /**
     * Sets the optional tint value for colors (-1.0 to 1.0), mainly for theme colors
     *
     * @param tint Tint value
     */
    public void setTint(Double tint) {
        this.tint = tint;
    }

    /**
     * Checks if this Color is defined (not None)
     *
     * @return
     */
    public boolean isDefined() {
        return type != ColorType.NONE;
    }

    /**
     * Gets the color value as BaseColor interface. If no color was defined ({@link ColorType#NONE} or getter
     * {@link Color#isDefined()} is false), null is returned.
     *
     * @return Color value as BaseColor
     */
    public BaseColor getValue() {
        return switch (type) {
            case RGB -> rgbColor;
            case INDEXED -> indexedColor;
            case THEME -> themeColor;
            case SYSTEM -> systemColor;
            case AUTO -> AutoColor.INSTANCE;
            case NONE -> null;
        };
    }

    /**
     * Private constructor to enforce factory methods
     */
    private Color() {
    }

    /**
     * Gets the ARGB string value of the color, if applicable
     * <p>
     * Remarks: This method only works for colors of Type Rgb or Indexed. For Theme, System or auto colors, the RGB
     * value depends on the actual theme or system settings and cannot be determined here.
     * </p>
     *
     * @return ARGB value or null, if not applicable
     */
    public String getArgbValue() {
        return switch (type) {
            case RGB -> rgbColor.getColorValue();
            case INDEXED -> indexedColor.getSrgbColor().getColorValue();
            default -> null;
        };
    }

    /**
     * Creates a Color with no color (empty element)
     *
     * @return Returns a dummy instance of the Color class, where {@link ColorType} is set to NONE
     */
    public static Color createNone() {
        Color color = new Color();
        color.type = ColorType.NONE;
        return color;
    }

    /**
     * Creates a Color with auto=true {@link Color#isAuto()}
     *
     * @return Returns a dummy instance of the Color class, where {@link ColorType} is set to AUTO
     */
    public static Color createAuto() {
        Color color = new Color();
        color.type = ColorType.AUTO;
        color.auto = true;
        return color;
    }

    /**
     * Creates a Color from an RGB/ARGB color
     *
     * @param rgbColor Instance of the type {@link SrgbColor}
     * @return Color instance with the value type {@link SrgbColor}
     */
    public static Color createRgb(SrgbColor rgbColor) {
        Color color = new Color();
        color.type = ColorType.RGB;
        color.rgbColor = rgbColor;
        return color;
    }

    /**
     * Creates a Color from an RGB string (e.g., "FFAABBCC")
     *
     * @param rgbValue RGB or ARGB value as string
     * @return Color instance with the value type {@link SrgbColor}
     * @throws StyleException Thrown if the passed RGB/ARGB value is invalid
     */
    public static Color createRgb(String rgbValue) {
        Color color = new Color();
        color.type = ColorType.RGB;
        color.rgbColor = new SrgbColor(rgbValue);
        return color;
    }

    /**
     * Creates a Color from an indexed color
     *
     * @param indexedColor Instance of the type {@link IndexedColor}
     * @return Color instance with the value type
     * @throws StyleException Thrown if the passed value was null
     */
    public static Color createIndexed(IndexedColor indexedColor) throws StyleException {
        if (indexedColor == null) {
            throw new StyleException("An indexed color cannot be null");
        }
        Color color = new Color();
        color.type = ColorType.INDEXED;
        color.indexedColor = indexedColor;
        return color;
    }

    /**
     * Creates a Color from an indexed color value (see {@link IndexedColor.Value})
     *
     * @param indexValue Color index enum value
     * @return Color instance with the value type {@link IndexedColor.Value}
     */
    public static Color createIndexed(IndexedColor.Value indexValue) {
        return createIndexed(new IndexedColor(indexValue));
    }

    /**
     * Creates a Color from a color index (0 to 65)
     *
     * @param index Color index (0 to 65)
     * @return Color instance with the value type {@link IndexedColor}
     * @throws StyleException Thrown if the passed index is invalid
     */
    public static Color createIndexed(int index) {
        return createIndexed(new IndexedColor(index));
    }

    /**
     * Creates a Color from a theme color instance
     *
     * @param themeColor Instance of the typ {@link ThemeColor}
     * @return Color instance with the value type {@link ThemeColor}
     */
    public static Color createTheme(ThemeColor themeColor) {
        return createTheme(themeColor, null);
    }

    /**
     * Creates a Color from a theme color instance
     *
     * @param themeColor Instance of the typ {@link ThemeColor}
     * @param tint       Tint value (from -1 to 1)
     * @return Color instance with the value type {@link ThemeColor}
     */
    public static Color createTheme(ThemeColor themeColor, Double tint) throws StyleException {
        if (themeColor == null) {
            throw new StyleException("A theme color cannot be null");
        }
        Color color = new Color();
        color.type = ColorType.THEME;
        color.themeColor = themeColor;
        color.tint = tint;
        return color;
    }

    /**
     * Creates a Color from a theme color scheme element
     *
     * @param element Color scheme element
     * @return Color instance with the value type {@link ThemeColor}
     */
    public static Color createTheme(Theme.ColorSchemeElement element) {
        return createTheme(element, null);
    }

    /**
     * Creates a Color from a theme color scheme element
     *
     * @param element Color scheme element
     * @param tint    Tint value (from -1 to 1)
     * @return Color instance with the value type {@link ThemeColor}
     */
    public static Color createTheme(Theme.ColorSchemeElement element, Double tint) {
        return createTheme(new ThemeColor(element), tint);
    }

    /**
     * Creates a Color from the index of a theme color scheme element
     *
     * @param index Color scheme element index
     * @return Color instance with the value type {@link ThemeColor}
     */
    public static Color createTheme(int index) {
        return createTheme(index, null);
    }

    /**
     * Creates a Color from the index of a theme color scheme element
     *
     * @param index Color scheme element index
     * @param tint  Tint value (from -1 to 1)
     * @return Color instance with the value type {@link ThemeColor}
     */
    public static Color createTheme(int index, Double tint) {
        return createTheme(new ThemeColor(index), tint);
    }

    /**
     * Creates a Color from a system color
     *
     * @param systemColor Instance of the type {@link SystemColor}
     * @return Color instance with the value type {@link SystemColor}
     */
    public static Color createSystem(SystemColor systemColor) throws StyleException {
        if (systemColor == null) {
            throw new StyleException("A system color cannot be null");
        }
        Color color = new Color();
        color.type = ColorType.SYSTEM;
        color.systemColor = systemColor;
        return color;
    }

    /**
     * Creates a Color from a system color instance
     *
     * @param systemColorValue System color value
     * @return Color instance with the value type {@link SystemColor}
     */
    public static Color createSystem(SystemColor.Value systemColorValue) {
        return createSystem(new SystemColor(systemColorValue));
    }

    /**
     * String representation of a Color instance
     *
     * @return String value
     */
    @Override
    public String toString() {
        return switch (type) {
            case RGB -> "RGBColor:" + rgbColor.getStringValue();
            case INDEXED -> "IndexedColor:" + indexedColor.getStringValue();
            case THEME -> "ThemeColor:" + themeColor.getStringValue();
            case SYSTEM -> "SystemColor:" + systemColor.getStringValue();
            case AUTO -> "Auto-Color";
            case NONE -> "Undefined Color";
        };
    }

    /**
     * Determines whether the specified object is equal to the current object
     *
     * @param o the reference object with which to compare.
     * @return True if both objects are equal
     */
    @Override
    public boolean equals(Object o) {
        if (!(o instanceof Color color)) {
            return false;
        }
        return type == color.type
                && auto == color.auto
                && Objects.equals(rgbColor, color.rgbColor)
                && Objects.equals(indexedColor, color.indexedColor)
                && Objects.equals(themeColor, color.themeColor)
                && Objects.equals(systemColor, color.systemColor)
                && Objects.equals(tint, color.tint);
    }

    /**
     * Gets the hash code of the instance
     *
     * @return Hash code
     */
    @Override
    public int hashCode() {
        return Objects.hash(type, auto, rgbColor, indexedColor, themeColor, systemColor, tint);
    }

    /**
     * Compares two instances for sorting purpose
     *
     * @param other the object to be compared.
     * @return Negative integer, zero, or positive integer as this object is less than, equal to, or greater than the
     * specified object
     * @throws StyleException thrown if the compared object is not from the type Color
     */
    @Override
    public int compareTo(Color other) {
        if (other == null) {
            return 1;
        }
        // 1) Compare by color type first
        int typeComparison = type.compareTo(other.type);
        if (typeComparison != 0) {
            return typeComparison;
        }

        // 2) Same type -> compare internal representation
        return switch (type) {
            case NONE, AUTO -> 0;
            case RGB -> Comparator.nullsFirst(String.CASE_INSENSITIVE_ORDER)
                    .compare(
                            rgbColor == null ? null : rgbColor.getStringValue(),
                            other.rgbColor == null ? null : other.rgbColor.getStringValue()
                    );
            // Numeric comparison of palette index
            case INDEXED -> indexedColor.getColorValue().compareTo(other.indexedColor.getColorValue());
            // Numeric comparison of theme slot
            case THEME -> {
                int themeComparison = themeColor.getColorValue().compareTo(other.themeColor.getColorValue());
                if (themeComparison != 0) {
                    yield themeComparison;
                }
                // Same theme slot -> compare tint
                yield Comparator.nullsFirst(Double::compareTo).compare(tint, other.tint);
            }
            // Enum-based comparison -> not string-based
            case SYSTEM -> systemColor.getColorValue().compareTo(other.systemColor.getColorValue());
        };
    }
}
