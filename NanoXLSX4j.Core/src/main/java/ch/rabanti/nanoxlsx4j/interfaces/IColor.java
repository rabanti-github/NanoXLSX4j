//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces;

/**
 * Interface to represent a non-typed color, either defined by the system or the user.
 *
 * @see IColorScheme
 * @see ITypedColor
 */
public interface IColor {

    /**
     * Color value. This may be a system-defined name or a formatted value such as an ARGB hex string.
     *
     * @return the string representation of this color
     */
    String getStringValue();
}
