/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.colors;

import java.util.Objects;

import ch.rabanti.nanoxlsx4j.internal.interfaces.TypedColor;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;
import ch.rabanti.nanoxlsx4j.utils.Validators;

public class SrgbColor implements TypedColor<String> {

    /**
     * Default color value (opaque black: #000000)
     */
    public static final String DEFAULT_SRGB_COLOR = "#FF000000";

    private String colorValue;

    /**
     * Gets the sRGB value (Hex code of RGB/ARGB)
     *
     * @return sRGB value as string
     */
    @Override
    public String getColorValue() {
        return colorValue;
    }

    /**
     * Sets the sRGB value (Hex code of RGB/ARGB). The value will be cast to upper case. If a 6-character RGB value is
     * provided, 'FF' is automatically prepended as alpha channel.
     *
     * @param colorValue sRGB value as string
     */
    @Override
    public void setColorValue(String colorValue) {
        Validators.validateGenericColor(colorValue);
        if (colorValue.length() == 6) {
            this.colorValue = "FF" + ParserUtils.toUpper(colorValue);
        } else {
            this.colorValue = ParserUtils.toUpper(colorValue);
        }
    }

    /**
     * Gets the string value of the color. The value is identical to {@link SrgbColor#getColorValue()} and defined as
     * interface implementation of {@link TypedColor}
     *
     * @return sRGB value as string
     */
    @Override
    public String getStringValue() {
        return colorValue;
    }

    /**
     * Default constructor
     */
    public SrgbColor() {
    }

    /**
     * Constructor with an RGB or ARGB hex string
     *
     * @param rgbValue RGB (6-char) or ARGB (8-char) hex string (case-insensitive)
     */
    public SrgbColor(String rgbValue) {
        this();
        setColorValue(rgbValue); // Validate
    }

    /**
     * Determines whether the specified object is equal to the current object
     *
     * @param o the reference object with which to compare.
     * @return True if the specified object is equal to the current object; otherwise, false
     */
    @Override
    public final boolean equals(Object o) {
        if (!(o instanceof SrgbColor srgbColor)) {
            return false;
        }

        return Objects.equals(colorValue, srgbColor.colorValue);
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
