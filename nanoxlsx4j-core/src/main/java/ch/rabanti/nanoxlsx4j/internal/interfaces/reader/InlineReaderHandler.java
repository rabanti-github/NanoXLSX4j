/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.reader;

import java.io.InputStream;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.internal.interfaces.Options;

/**
 * Handles in-line reader plug-ins after a reader has processed its XML content.
 */
@FunctionalInterface
@InternalApi
public interface InlineReaderHandler {

    /**
     * Handles the plug-ins registered for an in-line reader queue.
     *
     * @param stream        XML stream positioned by the invoking reader
     * @param workbook      workbook receiving the read data
     * @param queueUuid     UUID of the in-line plug-in queue
     * @param readerOptions reader options for the current operation
     * @param index         optional index, such as a worksheet index; {@code null} when not applicable
     */
    void handle(InputStream stream, Workbook workbook, String queueUuid, Options readerOptions, Integer index);
}
