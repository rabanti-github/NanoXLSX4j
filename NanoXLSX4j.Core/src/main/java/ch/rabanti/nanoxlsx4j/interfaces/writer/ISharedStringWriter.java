//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces.writer;

import ch.rabanti.nanoxlsx4j.interfaces.ISortedMap;

/**
 * Interface used by shared-string writer plug-ins.
 * <p>
 * Java equivalent of the {@code internal} .NET interface
 * {@code NanoXLSX.Interfaces.Writer.ISharedStringWriter}.
 * </p>
 *
 * @implNote Framework-internal interface. Not intended for use by external plugin implementors.
 * @see ISortedMap
 */
public interface ISharedStringWriter extends IPluginWriter {

    /**
     * Gets the sorted map that stores the shared strings.
     * The indices in the map correspond to the {@code si} element positions in
     * {@code sharedStrings.xml}.
     *
     * @return the shared-strings {@link ISortedMap}
     */
    ISortedMap getSharedStrings();

    /**
     * Gets the total count of shared-string references across all worksheets.
     *
     * @return the total shared-string reference count
     */
    int getSharedStringsTotalCount();

    /**
     * Sets the total count of shared-string references across all worksheets.
     *
     * @param count the total count
     */
    void setSharedStringsTotalCount(int count);
}
