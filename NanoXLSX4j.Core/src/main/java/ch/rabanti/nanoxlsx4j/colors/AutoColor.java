//STATUS : CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.colors;

import ch.rabanti.nanoxlsx4j.interfaces.IColor;

/**
 * Class representing an automatic color.
 * <p>
 * This class does not carry any value. It is only for the purpose of identification. Java equivalent of
 * {@code NanoXLSX.Colors.AutoColor} in the .NET reference.
 * </p>
 */
public class AutoColor implements IColor {

    /**
     * Static singleton instance of {@link AutoColor} to avoid multiple instances (all instances are equivalent since
     * there is no value)
     */
    private static AutoColor instance;

    /**
     * Gets the static singleton instance of {@link AutoColor} to avoid multiple instances (all instances are equivalent
     * since there is no value)
     *
     * @return the singleton instance of {@link AutoColor}
     */
    public static synchronized AutoColor getInstance() {
        if (instance == null) {
            return new AutoColor();
        }
        return instance;
    }

    private AutoColor() {
    }

    /**
     * The string value of an auto color is always null
     *
     * @return null
     */
    @Override
    public String getStringValue() {
        return null;
    }
}
