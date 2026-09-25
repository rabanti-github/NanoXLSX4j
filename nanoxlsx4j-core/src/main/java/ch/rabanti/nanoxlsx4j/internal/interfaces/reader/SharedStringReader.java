/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.reader;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;

import java.util.List;

/**
 * Interface, used by shared string readers
 */
@InternalApi
public interface SharedStringReader extends DocumentReader {

    /**
     * Gets the resolved list of shared strings. The indices of the shared strings are defined by the order of the
     * strings in the list.
     *
     * @return List of shared strings
     */
    List<String> getSharedStrings();
}
