/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.themes;

/**
 * Class representing the theme of a workbook.
 * <p>
 * Java equivalent of {@code NanoXLSX.Themes.Theme} in the .NET reference.
 * Full implementation will be added during the NanoXLSX4j v3 source migration.
 * </p>
 */
public class Theme {

    /**
     * Enum defining the available theme color scheme elements (indices 0–11).
     * <p>
     * Java equivalent of {@code NanoXLSX.Themes.Theme.ColorSchemeElement} in the .NET reference.
     * </p>
     */
    public enum ColorSchemeElement {
        /** Dark 1 (index 0) */
        dk1,
        /** Light 1 (index 1) */
        lt1,
        /** Dark 2 (index 2) */
        dk2,
        /** Light 2 (index 3) */
        lt2,
        /** Accent 1 (index 4) */
        accent1,
        /** Accent 2 (index 5) */
        accent2,
        /** Accent 3 (index 6) */
        accent3,
        /** Accent 4 (index 7) */
        accent4,
        /** Accent 5 (index 8) */
        accent5,
        /** Accent 6 (index 9) */
        accent6,
        /** Hyperlink color (index 10) */
        hlink,
        /** Followed hyperlink color (index 11) */
        folHlink
    }

    /**
     * Gets the default workbook theme
     *
     * @return Default theme instance
     */
    public static Theme getDefaultTheme() {
        return new Theme();
    }

    // Preparation stub — to be fully implemented during v3 source migration.
}
