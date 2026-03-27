//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces;

/**
 * Interface to represent a typed color with a specific value, extending {@link IColor}.
 *
 * @param <T> the concrete type of the color value (e.g. {@code String} for ARGB,
 *            {@code Integer} for indexed colors)
 * @see IColor
 * @see IColorScheme
 */
public interface ITypedColor<T> extends IColor {

    /**
     * Gets the color value in its native type.
     *
     * @return the color value
     */
    T getColorValue();

    /**
     * Sets the color value.
     *
     * @param value the color value to set
     */
    void setColorValue(T value);
}
