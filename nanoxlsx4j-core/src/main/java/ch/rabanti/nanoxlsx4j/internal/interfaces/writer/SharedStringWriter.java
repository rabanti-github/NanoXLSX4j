/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.writer;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.internal.interfaces.SortableMap;

/**
 * Interface, used by shared string writers
 */
@InternalApi
public interface SharedStringWriter extends PluginWriter {

    /**
     * Gets the sorted map that contains the shared strings
     *
     * @return Sortable map of shared strings
     */
    SortableMap getSharedStrings();

    /**
     * Gets the total number of shared strings
     *
     * @return Total count of shared strings
     */
    int getSharedStringsTotalCount();

    /**
     * Sets the total number of shared strings
     *
     * @param sharedStringsTotalCount Total count of shared strings
     */
    void setSharedStringsTotalCount(int sharedStringsTotalCount);
}
