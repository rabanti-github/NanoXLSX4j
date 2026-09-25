/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.writer;

import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.annotations.InternalApi;

/**
 * Interface, used by worksheet writers
 */
@InternalApi
public interface WorksheetWriter {
    /**
     * Gets the currently active worksheet
     *
     * @return Currently active worksheet
     */
    Worksheet getCurrentWorksheet();

    /**
     * Sets the currently active worksheet
     *
     * @param worksheet Currently active worksheet
     */
    void setCurrentWorksheet(Worksheet worksheet);

    /**
     * Method to initiate freeing memory used by the XML element
     */
    void releaseXmlElement();
}
