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
 * Interface, used by worksheet readers
 */
@InternalApi
public interface WorksheetReader extends DocumentReader {

    /**
     * Gets the (r)ID (1-based) of the currently processed worksheet.
     *
     * @return Worksheet rID
     */
    int getCurrentWorksheetId();

    /**
     * Sets the (r)ID (1-based) of the currently processed worksheet.
     *
     * @param worksheetId Worksheet rID
     */
    void setCurrentWorksheetId(int worksheetId);

    /**
     * Gets the list of shared strings. The index of the list corresponds to the index defined in cell values. A null
     * value indicates that the workbook has no usable shared strings relationship.
     *
     * @return Shared strings
     */
    List<String> getSharedStrings();

    /**
     * Sets the list of shared strings. The index of the list corresponds to the index defined in cell values. A null
     * value indicates that the workbook has no usable shared strings relationship.
     *
     * @param sharedStrings Shared strings
     */
    void setSharedStrings(List<String> sharedStrings);
}
