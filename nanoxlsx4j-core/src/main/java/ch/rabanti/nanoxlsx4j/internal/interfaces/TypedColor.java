/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces;

/**
 * Interface to represent typed color with a specific value, based on a generic type {@code T}
 * @param <T> The concrete color value type
 */
public interface TypedColor<T> extends BaseColor {

    /**
     * Gets the color value of the type {@code  T}
     * @return Value of the type {@code  T}
     */
    T getColorValue();

    /**
     * Sets the color value of the type {@code T}
     * @param colorValue Value of the type {@code T}
     */
    void setColorValue(T colorValue);
}
