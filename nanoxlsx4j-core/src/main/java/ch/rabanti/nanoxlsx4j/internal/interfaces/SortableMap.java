/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.internal.interfaces;

/**
 * Interface to represent a sorted map with FormattableText as key and string as value
 */
public interface SortableMap {

    /// <summary>
    /// Number of map entries
    /// </summary>
    int size();

    /**
     * Method to add a key value pair (IFormattableText as key and its index in the worksheet as value)
     *
     * @param text           Text (Key) as string
     * @param referenceIndex Reference index as string
     * @return Returns the resolved string (either added or returned from an existing entry) of the reference index
     */
    String add(FormattableText text, String referenceIndex);
}
