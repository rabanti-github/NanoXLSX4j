//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces;

/**
 * Interface used by text-related partitions of option classes (e.g. {@code ImportOptions}).
 *
 * @see IOptions
 */
public interface ITextOptions {

    /**
     * Returns whether phonetic characters (e.g. ruby / Furigana / Zhuyin fuhao) in strings
     * are appended in brackets after the transcribed symbols.
     * <p>
     * By default, phonetic characters are removed from strings.
     * This option is applied globally and is not limited to specific rows or columns.
     * </p>
     *
     * @return {@code true} if phonetic characters are preserved
     */
    boolean isEnforcePhoneticCharacterImport();

    /**
     * Sets whether phonetic characters are preserved during import.
     *
     * @param enforce {@code true} to preserve phonetic characters
     */
    void setEnforcePhoneticCharacterImport(boolean enforce);

    /**
     * Returns whether empty cells are interpreted as strings with an empty value.
     * <p>
     * When {@code false} (default), empty cells have type {@code Empty} and value {@code null}.
     * </p>
     *
     * @return {@code true} if empty cells are treated as empty strings
     */
    boolean isEnforceEmptyValuesAsString();

    /**
     * Sets whether empty cells are treated as empty strings.
     *
     * @param enforce {@code true} to treat empty cells as strings
     */
    void setEnforceEmptyValuesAsString(boolean enforce);
}
