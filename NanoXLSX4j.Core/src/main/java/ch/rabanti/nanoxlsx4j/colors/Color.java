//STATUS: REVISE ENUM
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.colors;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.interfaces.IColor;
import ch.rabanti.nanoxlsx4j.themes.Theme;

import java.util.Objects;

/**
 * Compound class representing a color in various representations (RGB, indexed, theme, system or automatic).
 * <p>
 * Java equivalent of {@code NanoXLSX.Colors.Color} in the .NET reference.
 * </p>
 *
 * @see ch.rabanti.nanoxlsx4j.interfaces.writer.IColorWriter
 */
public class Color implements Comparable<Object> {

    // ### E N U M ###

    /**
     * Enum defining the type of color representation
     */
    public enum ColorType {
        /** No color defined */
        None,
        /** Automatic color (determined by the application) */
        Auto,
        /** RGB/ARGB color value */
        Rgb,
        /** Legacy indexed color (0–65) */
        Indexed,
        /** Theme color reference (0–11) */
        Theme,
        /** System color */
        System
    }

    // ### F I E L D S ###

    private ColorType type = ColorType.None;
    private boolean auto;
    private SrgbColor rgbColor;
    private IndexedColor indexedColor;
    private ThemeColor themeColor;
    private SystemColor systemColor;
    private Double tint;

    // ### C O N S T R U C T O R ###

    private Color() {
    }

    // ### G E T T E R S ###

    /**
     * Gets the color type of this instance
     *
     * @return Color type
     */
    public ColorType getType() {
        return type;
    }

    /**
     * Gets whether the auto attribute is set - if true, color is automatically determined
     *
     * @return True if auto
     */
    public boolean isAuto() {
        return auto;
    }

    /**
     * Gets the RGB/ARGB color value (only valid when {@link #getType()} is {@link ColorType#Rgb})
     *
     * @return SrgbColor instance or null
     */
    public SrgbColor getRgbColor() {
        return rgbColor;
    }

    /**
     * Gets the indexed color value (only valid when {@link #getType()} is {@link ColorType#Indexed})
     *
     * @return IndexedColor instance or null
     */
    public IndexedColor getIndexedColor() {
        return indexedColor;
    }

    /**
     * Gets the theme color value (only valid when {@link #getType()} is {@link ColorType#Theme})
     *
     * @return ThemeColor instance or null
     */
    public ThemeColor getThemeColor() {
        return themeColor;
    }

    /**
     * Gets the system color value (only valid when {@link #getType()} is {@link ColorType#System})
     *
     * @return SystemColor instance or null
     */
    public SystemColor getSystemColor() {
        return systemColor;
    }

    /**
     * Gets the optional tint value (-1.0 to 1.0). Mainly meaningful for theme colors. Positive values lighten, negative
     * values darken.
     *
     * @return Tint value or null if not set
     */
    public Double getTint() {
        return tint;
    }

    /**
     * Sets the optional tint value (-1.0 to 1.0)
     *
     * @param tint Tint value or null to clear
     */
    public void setTint(Double tint) {
        this.tint = tint;
    }

    /**
     * Returns whether this Color is defined (type is not {@link ColorType#None})
     *
     * @return True if defined
     */
    public boolean isDefined() {
        return type != ColorType.None;
    }

    /**
     * Gets the underlying color value as {@link IColor}. Returns null if {@link #getType()} is {@link ColorType#None}.
     *
     * @return IColor instance or null
     */
    public IColor getValue() {
        switch (type) {
            case Rgb:
                return rgbColor;
            case Indexed:
                return indexedColor;
            case Theme:
                return themeColor;
            case System:
                return systemColor;
            case Auto:
                return AutoColor.getInstance();
            default:
                return null;
        }
    }

    /**
     * Gets the ARGB string value of this color, if applicable. Works only for {@link ColorType#Rgb} and
     * {@link ColorType#Indexed}. For theme, system, or auto colors, null is returned.
     *
     * @return ARGB hex string or null
     */
    public String getArgbValue() {
        if (type == ColorType.Rgb) {
            return rgbColor != null ? rgbColor.getColorValue() : null;
        }
        else if (type == ColorType.Indexed) {
            return indexedColor != null ? indexedColor.getArgbValue() : null;
        }
        return null;
    }

    // ### F A C T O R Y   M E T H O D S ###

    /**
     * Creates a {@link Color} with no color (empty element, {@link ColorType#None})
     *
     * @return Color instance with type None
     */
    public static Color createNone() {
        Color c = new Color();
        c.type = ColorType.None;
        return c;
    }

    /**
     * Creates a {@link Color} with auto=true ({@link ColorType#Auto})
     *
     * @return Color instance with type Auto
     */
    public static Color createAuto() {
        Color c = new Color();
        c.type = ColorType.Auto;
        c.auto = true;
        return c;
    }

    /**
     * Creates a {@link Color} from an {@link SrgbColor} instance
     *
     * @param color SrgbColor instance
     * @return Color instance with type Rgb
     */
    public static Color createRgb(SrgbColor color) {
        Color c = new Color();
        c.type = ColorType.Rgb;
        c.rgbColor = color;
        return c;
    }

    /**
     * Creates a {@link Color} from an RGB or ARGB hex string
     *
     * @param rgbValue RGB (6-char) or ARGB (8-char) hex string
     * @return Color instance with type Rgb
     * @throws StyleException Thrown if the hex string is invalid
     */
    public static Color createRgb(String rgbValue) {
        Color c = new Color();
        c.type = ColorType.Rgb;
        c.rgbColor = new SrgbColor(rgbValue);
        return c;
    }

    /**
     * Creates a {@link Color} from an {@link IndexedColor} instance
     *
     * @param color IndexedColor instance
     * @return Color instance with type Indexed
     * @throws StyleException Thrown if the color is null
     */
    public static Color createIndexed(IndexedColor color) {
        if (color == null) {
            throw new StyleException("An indexed color cannot be null");
        }
        Color c = new Color();
        c.type = ColorType.Indexed;
        c.indexedColor = color;
        return c;
    }

    /**
     * Creates a {@link Color} from an {@link IndexedColor.Value} enum
     *
     * @param indexValue Color index enum value
     * @return Color instance with type Indexed
     */
    public static Color createIndexed(IndexedColor.Value indexValue) {
        Color c = new Color();
        c.type = ColorType.Indexed;
        c.indexedColor = new IndexedColor(indexValue);
        return c;
    }

    /**
     * Creates a {@link Color} from a numeric palette index (0–65)
     *
     * @param index Color palette index
     * @return Color instance with type Indexed
     * @throws StyleException Thrown if the index is out of range
     */
    public static Color createIndexed(int index) {
        Color c = new Color();
        c.type = ColorType.Indexed;
        c.indexedColor = new IndexedColor(index);
        return c;
    }

    /**
     * Creates a {@link Color} from a {@link ThemeColor} instance
     *
     * @param color ThemeColor instance
     * @return Color instance with type Theme
     * @throws StyleException Thrown if the color is null
     */
    public static Color createTheme(ThemeColor color) {
        return createTheme(color, null);
    }

    /**
     * Creates a {@link Color} from a {@link ThemeColor} instance with an optional tint
     *
     * @param color ThemeColor instance
     * @param tint  Optional tint value (-1.0 to 1.0), or null
     * @return Color instance with type Theme
     * @throws StyleException Thrown if the color is null
     */
    public static Color createTheme(ThemeColor color, Double tint) {
        if (color == null) {
            throw new StyleException("A theme color cannot be null");
        }
        Color c = new Color();
        c.type = ColorType.Theme;
        c.themeColor = color;
        c.tint = tint;
        return c;
    }

    /**
     * Creates a {@link Color} from a {@link Theme.ColorSchemeElement}
     *
     * @param element Color scheme element
     * @return Color instance with type Theme
     */
    public static Color createTheme(Theme.ColorSchemeElement element) {
        return createTheme(element, null);
    }

    /**
     * Creates a {@link Color} from a {@link Theme.ColorSchemeElement} with an optional tint
     *
     * @param element Color scheme element
     * @param tint    Optional tint value (-1.0 to 1.0), or null
     * @return Color instance with type Theme
     */
    public static Color createTheme(Theme.ColorSchemeElement element, Double tint) {
        Color c = new Color();
        c.type = ColorType.Theme;
        c.themeColor = new ThemeColor(element);
        c.tint = tint;
        return c;
    }

    /**
     * Creates a {@link Color} from a theme color scheme index (0–11)
     *
     * @param index Theme color index
     * @return Color instance with type Theme
     * @throws StyleException Thrown if the index is out of range
     */
    public static Color createTheme(int index) {
        return createTheme(index, null);
    }

    /**
     * Creates a {@link Color} from a theme color scheme index with an optional tint
     *
     * @param index Theme color index
     * @param tint  Optional tint value (-1.0 to 1.0), or null
     * @return Color instance with type Theme
     * @throws StyleException Thrown if the index is out of range
     */
    public static Color createTheme(int index, Double tint) {
        Color c = new Color();
        c.type = ColorType.Theme;
        c.themeColor = new ThemeColor(index);
        c.tint = tint;
        return c;
    }

    /**
     * Creates a {@link Color} from a {@link SystemColor} instance
     *
     * @param color SystemColor instance
     * @return Color instance with type System
     * @throws StyleException Thrown if the color is null
     */
    public static Color createSystem(SystemColor color) {
        if (color == null) {
            throw new StyleException("A system color cannot be null");
        }
        Color c = new Color();
        c.type = ColorType.System;
        c.systemColor = color;
        return c;
    }

    /**
     * Creates a {@link Color} from a {@link SystemColor.Value} enum
     *
     * @param systemColorValue System color enum value
     * @return Color instance with type System
     */
    public static Color createSystem(SystemColor.Value systemColorValue) {
        Color c = new Color();
        c.type = ColorType.System;
        c.systemColor = new SystemColor(systemColorValue);
        return c;
    }

    // ### O V E R R I D E S ###

    /**
     * String representation of a Color instance
     *
     * @return String value
     */
    @Override
    public String toString() {
        switch (type) {
            case Rgb:
                return "RGBColor:" + (rgbColor != null ? rgbColor.getStringValue() : "null");
            case Indexed:
                return "IndexedColor:" + (indexedColor != null ? indexedColor.getStringValue() : "null");
            case Theme:
                return "ThemeColor:" + (themeColor != null ? themeColor.getStringValue() : "null");
            case System:
                return "SystemColor:" + (systemColor != null ? systemColor.getStringValue() : "null");
            case Auto:
                return "Auto-Color";
            default:
                return "Undefined Color";
        }
    }

    /**
     * Determines whether the specified object is equal to the current object
     *
     * @param obj Other object to compare
     * @return True if both objects are equal
     */
    @Override
    public boolean equals(Object obj) {
        if (this == obj) return true;
        if (!(obj instanceof Color)) return false;
        Color other = (Color) obj;
        return type == other.type && auto == other.auto && Objects.equals(rgbColor, other.rgbColor) && Objects.equals(indexedColor, other.indexedColor) && Objects.equals(themeColor, other.themeColor) && Objects.equals(systemColor, other.systemColor) && Objects.equals(tint, other.tint);
    }

    /**
     * Gets the hash code of the instance
     *
     * @return Hash code
     */
    @Override
    public int hashCode() {
        int hashCode = -1729664991;
        hashCode = hashCode * -1521134295 + type.hashCode();
        hashCode = hashCode * -1521134295 + Boolean.hashCode(auto);
        hashCode = hashCode * -1521134295 + Objects.hashCode(rgbColor);
        hashCode = hashCode * -1521134295 + Objects.hashCode(indexedColor);
        hashCode = hashCode * -1521134295 + Objects.hashCode(themeColor);
        hashCode = hashCode * -1521134295 + Objects.hashCode(systemColor);
        hashCode = hashCode * -1521134295 + Objects.hashCode(tint);
        return hashCode;
    }

    /**
     * Compares two instances for sorting purpose
     *
     * @param obj Object to compare
     * @return Negative integer, zero, or positive integer as this object is less than, equal to, or greater than the
     * specified object
     */
    @Override
    public int compareTo(Object obj) {
        if (obj == null) {
            return 1;
        }
        if (!(obj instanceof Color)) {
            throw new StyleException("The provided object to compare is not a Color");
        }
        Color other = (Color) obj;

        int typeCompare = type.compareTo(other.type);
        if (typeCompare != 0) {
            return typeCompare;
        }

        switch (type) {
            case None:
            case Auto:
                return 0;
            case Rgb: {
                String a = rgbColor != null ? rgbColor.getStringValue() : null;
                String b = other.rgbColor != null ? other.rgbColor.getStringValue() : null;
                return compareStrings(a, b);
            }
            case Indexed: {
                int ai = indexedColor != null ? indexedColor.getColorValue().ordinal() : -1;
                int bi = other.indexedColor != null ? other.indexedColor.getColorValue().ordinal() : -1;
                return Integer.compare(ai, bi);
            }
            case Theme: {
                int at = themeColor != null ? themeColor.getColorValue().ordinal() : -1;
                int bt = other.themeColor != null ? other.themeColor.getColorValue().ordinal() : -1;
                int themeCompare = Integer.compare(at, bt);
                if (themeCompare != 0) return themeCompare;
                return compareNullableDoubles(tint, other.tint);
            }
            case System: {
                int as = systemColor != null ? systemColor.getColorValue().ordinal() : -1;
                int bs = other.systemColor != null ? other.systemColor.getColorValue().ordinal() : -1;
                return Integer.compare(as, bs);
            }
            default: {
                IColor va = getValue();
                IColor vb = other.getValue();
                String sa = va != null ? va.getStringValue() : null;
                String sb = vb != null ? vb.getStringValue() : null;
                return compareStrings(sa, sb);
            }
        }
    }

    private static int compareStrings(String a, String b) {
        if (a == null && b == null) return 0;
        if (a == null) return -1;
        if (b == null) return 1;
        return a.compareToIgnoreCase(b);
    }

    private static int compareNullableDoubles(Double a, Double b) {
        if (a == null && b == null) return 0;
        if (a == null) return -1;
        if (b == null) return 1;
        return Double.compare(a, b);
    }
}
