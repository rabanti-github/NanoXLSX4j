/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces;

/**
 * Represents a color value defined by the system or the user.
 */
public interface BaseColor {

    /**
     * Gets the string representation used for the color value.
     *
     * @return string representation of the color value
     */
    String getStringValue();
}
