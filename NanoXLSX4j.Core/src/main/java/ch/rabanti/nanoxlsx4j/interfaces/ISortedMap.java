//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces;

/**
 * Interface to represent a sorted map with {@link IFormattableText} as key and
 * a string index as value.
 * <p>
 * Java equivalent of the {@code internal} .NET interface {@code NanoXLSX.Interfaces.ISortedMap}.
 * </p>
 *
 * @implNote Framework-internal interface. Not intended for use by external plugin implementors.
 * @see ch.rabanti.nanoxlsx4j.interfaces.writer.ISharedStringWriter
 */
public interface ISortedMap {

    /**
     * Returns the number of entries currently in the map.
     *
     * @return the entry count
     */
    int getCount();

    /**
     * Adds a key–value pair to the map, or returns the existing value if the key
     * is already present.
     *
     * @param text           the formattable text used as key
     * @param referenceIndex the index string to associate with the text
     * @return the resolved reference index (either newly added or pre-existing)
     */
    String add(IFormattableText text, String referenceIndex);
}
