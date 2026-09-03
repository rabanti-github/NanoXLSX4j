/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces;

/**
 * Interface used by text partitions of option classes (e.g. ReaderOptions)
 */
public interface TextOptions {
    /**
     * Gets whether phonetic characters are enforced during the import of a workbook
     * @return If true, phonetic characters (like ruby characters / Furigana / Zhuyin fuhao) in strings are added in brackets after the transcribed symbols. By default, phonetic characters are removed from strings.
     * <p>Remarks: This option is not applicable to specific rows or a start column (applied globally)</p>
     */
    boolean getEnforcePhoneticCharacterImport();

    /**
     * Sets whether phonetic characters are enforced during the import of a workbook
     * @param value If true, phonetic characters (like ruby characters / Furigana / Zhuyin fuhao) in strings are added in brackets after the transcribed symbols. By default, phonetic characters are removed from strings.
     *              <p>Remarks: This option is not applicable to specific rows or a start column (applied globally)</p>
     */
    void setEnforcePhoneticCharacterImport(boolean value);

    /**
     * Gets whether empty values are enforced as strings during the import of a workbook
     * @return If true, empty cells will be interpreted as type of string with an empty value. If false, the type will be Empty and the value null
     */
    boolean getEnforceEmptyValuesAsString();

    /**
     * Sets whether empty values are enforced as strings during the import of a workbook
     * @param value If true, empty cells will be interpreted as type of string with an empty value. If false, the type will be Empty and the value null
     */
    void setEnforceEmptyValuesAsString(boolean value);
}
