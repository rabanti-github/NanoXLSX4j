//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.colors;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.interfaces.ITypedColor;
import ch.rabanti.nanoxlsx4j.themes.Theme;

/**
 * Class representing a color defined by a theme color scheme element.
 * <p>
 * Java equivalent of {@code NanoXLSX.Colors.ThemeColor} in the .NET reference.
 * </p>
 *
 * @see Theme.ColorSchemeElement
 */
public class ThemeColor implements ITypedColor<Theme.ColorSchemeElement> {

    private Theme.ColorSchemeElement colorValue;

    /**
     * Gets the color scheme element
     *
     * @return Color scheme element
     */
    @Override
    public Theme.ColorSchemeElement getColorValue() {
        return colorValue;
    }

    /**
     * Sets the color scheme element
     *
     * @param value Color scheme element
     */
    @Override
    public void setColorValue(Theme.ColorSchemeElement value) {
        this.colorValue = value;
    }

    /**
     * Gets the numeric index of the color scheme element as a string
     *
     * @return Numeric index string (0–11)
     */
    @Override
    public String getStringValue() {
        return String.valueOf(colorValue.ordinal());
    }

    /**
     * Default constructor. Color value defaults to the first element ({@link Theme.ColorSchemeElement#dk1})
     */
    public ThemeColor() {
        this.colorValue = Theme.ColorSchemeElement.dk1;
    }

    /**
     * Constructor with a color scheme element
     *
     * @param color Color scheme element
     */
    public ThemeColor(Theme.ColorSchemeElement color) {
        this.colorValue = color;
    }

    /**
     * Constructor with a numeric index (0–11)
     *
     * @param index Theme color index
     * @throws StyleException Thrown if the index is out of range (0–11)
     */
    public ThemeColor(int index) {
        if (index < 0 || index > 11) {
            throw new StyleException("Theme color index must be between 0 and 11.");
        }
        this.colorValue = Theme.ColorSchemeElement.values()[index];
    }

    /**
     * Determines whether the specified object is equal to the current system color instance
     *
     * @param obj Other object to compare
     * @return True if both objects are equal
     */
    @Override
    public boolean equals(Object obj) {
        if (this == obj) return true;
        if (!(obj instanceof ThemeColor)) return false;
        return colorValue == ((ThemeColor) obj).colorValue;
    }

    /**
     * Gets the hash code of the instance
     *
     * @return Hash code
     */
    @Override
    public int hashCode() {
        return 800285905 + colorValue.hashCode();
    }

    /**
     * Gets the string representation of the theme color, which is the numeric index of the color scheme element
     *
     * @return String value
     */
    @Override
    public String toString() {
        return getStringValue();
    }
}
