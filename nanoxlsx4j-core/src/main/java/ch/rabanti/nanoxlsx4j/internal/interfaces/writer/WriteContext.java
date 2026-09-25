/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.writer;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.annotations.InternalApi;

@InternalApi
public interface WriteContext {

    /**
     * Gets the workbook handled by the current writing operation.
     *
     * @return Workbook instance
     */
    Workbook getWorkbook();

    /**
     * Gets the processing data that contains data outside of the workbook (used for preparation etc.)
     *
     * @return Writer processing data
     */
    WriterProcessingData getWriterProcessingData();

    /**
     * Marks a writing feature as successfully prepared for the current writing operation.
     *
     * @param featureUuid UUID of the prepared writing feature.
     */
    void markFeatureAsPrepared(String featureUuid);

    /**
     * Determines whether a writing feature was successfully prepared for the current writing operation.
     *
     * @param featureUuid UUID of the writing feature.
     * @return True if the feature was prepared; otherwise false.
     */
    boolean ssFeaturePrepared(String featureUuid);
}
