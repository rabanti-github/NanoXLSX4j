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
import ch.rabanti.nanoxlsx4j.utils.Validators;

import java.util.Objects;

/**
 * Class representing a generic sRGB color (with or without alpha channel).
 * <p>
 * Java equivalent of {@code NanoXLSX.Colors.SrgbColor} in the .NET reference.
 * </p>
 */
public class SrgbColor implements ITypedColor<String> {

    /**
     * Default sRGB color value (opaque black: #000000)
     */
    public static final String DEFAULT_SRGB_COLOR = "FF000000";

    private String colorValue;

    /**
     * Gets the sRGB value as an ARGB hex string. If the original input was a 6-character RGB string, 'FF' was
     * automatically prepended as the alpha channel.
     *
     * @return ARGB hex string (8 characters, upper-case)
     */
    @Override
    public String getColorValue() {
        return colorValue;
    }

    /**
     * Sets the sRGB value. The value is converted to upper-case. If a 6-character RGB value is provided, 'FF' is
     * automatically prepended as the alpha channel.
     *
     * @param value RGB (6-char) or ARGB (8-char) hex string (case-insensitive)
     * @throws StyleException Thrown if the value is null, empty, or not a valid hex string of length 6 or 8
     */
    @Override
    public void setColorValue(String value) {
        Validators.validateColor(value, false);
        String upper = value.toUpperCase();
        if (upper.length() == 6) {
            this.colorValue = "FF" + upper;
        }
        else {
            this.colorValue = upper;
        }
    }

    /**
     * Gets the string value of the color, identical to {@link #getColorValue()}
     *
     * @return ARGB hex string
     */
    @Override
    public String getStringValue() {
        return colorValue;
    }

    /**
     * Default constructor. Color value is not set (null)
     */
    public SrgbColor() {
    }

    /**
     * Constructor with an RGB or ARGB hex string
     *
     * @param rgb RGB (6-char) or ARGB (8-char) hex string (case-insensitive)
     * @throws StyleException Thrown if the color value is null, empty, or not a valid hex color string
     */
    public SrgbColor(String rgb) {
        setColorValue(rgb);
    }

    /**
     * Determines whether the specified object is equal to the current object
     *
     * @param obj Other object to compar
     * @return True if the specified object is equal to the current object; otherwise, false
     */
    @Override
    public boolean equals(Object obj) {
        if (this == obj) return true;
        if (!(obj instanceof SrgbColor)) return false;
        SrgbColor other = (SrgbColor) obj;
        return Objects.equals(colorValue, other.colorValue);
    }

    /**
     * Gets the hash code of the instance
     *
     * @return Hash code
     */
    @Override
    public int hashCode() {
        return Objects.hashCode(colorValue);
    }

    /**
     * Returns a string that represents the current object, which is the color value
     *
     * @return String value
     */
    @Override
    public String toString() {
        return colorValue;
    }
}
