/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.colors;

/**
 * Represents a color in one of the forms supported by NanoXLSX4j.
 *
 * <p>The complete color model and its factory methods will be added when the corresponding Core domain feature is
 * ported.</p>
 */
public final class Color implements ColorValue, Comparable<Color> {
    @Override
    public int compareTo(Color o) {
        return 0; // TODO: Not yet implemented
    }

    @Override
    public String getStringValue() {
      return null; // TODO: Not yet implemented
    }

    private Color() {
    }
}
