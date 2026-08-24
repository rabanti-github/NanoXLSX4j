/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.colors;

import java.util.Objects;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.internal.interfaces.TypedColor;
import ch.rabanti.nanoxlsx4j.themes.Theme;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;

public class ThemeColor implements TypedColor<Theme.ColorSchemeElement> {

    private Theme.ColorSchemeElement colorValue = Theme.ColorSchemeElement.DARK_1; // Default

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
     * @param colorValue Color scheme element
     */
    @Override
    public void setColorValue(Theme.ColorSchemeElement colorValue) {
        this.colorValue = colorValue;
    }

    /**
     * Gets the internal, numeric OOXML string value of the enum, defined in {@link ThemeColor#getColorValue()}
     *
     * @return The string value
     */
    @Override
    public String getStringValue() {
        return ParserUtils.toString(colorValue.value);
    }

    /**
     * Default constructor
     */
    public ThemeColor() {
    }

    /**
     * Constructor with color scheme element as parameter
     *
     * @param colorValue Color value
     */
    public ThemeColor(Theme.ColorSchemeElement colorValue) {
        this();
        this.colorValue = colorValue;
    }

    /**
     * Constructor with index as parameter
     *
     * @param index Theme color index
     * @throws StyleException Thrown if the color scheme element index is out of range
     */
    public ThemeColor(int index) throws StyleException {
        this();
        if (index < 0 || index > 11) {
            throw new StyleException("Indexed color value must be between 0 and 65.");
        }
        colorValue = Theme.ColorSchemeElement.values()[index];
    }

    /**
     * Determines whether the specified object is equal to the current system color instance
     *
     * @param o the reference object with which to compare.
     * @return True if both objects are equal
     */
    @Override
    public final boolean equals(Object o) {
        if (!(o instanceof ThemeColor that)) {
            return false;
        }

        return colorValue == that.colorValue;
    }

    /**
     * Gets the hash code of the instance
     *
     * @return hash code
     */
    @Override
    public int hashCode() {
        return Objects.hashCode(colorValue);
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
