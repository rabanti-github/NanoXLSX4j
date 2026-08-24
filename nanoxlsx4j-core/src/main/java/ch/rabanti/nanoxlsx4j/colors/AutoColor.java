/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.colors;

import ch.rabanti.nanoxlsx4j.internal.interfaces.BaseColor;

/**
 * Class representing an automatic color.
 */
public class AutoColor implements BaseColor {

    /**
     * Static instance of the AutoColor class to avoid multiple instances (instances does not deviate)
     */
    public static final AutoColor INSTANCE = new AutoColor();

    /**
     * The string value of an auto color is always null
     * @return Null in any case
     */
    @Override
    public String getStringValue() {
        return null;
    }
}
