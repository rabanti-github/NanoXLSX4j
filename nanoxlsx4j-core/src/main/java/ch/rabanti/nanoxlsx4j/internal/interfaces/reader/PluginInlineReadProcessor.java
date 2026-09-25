/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.reader;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.internal.interfaces.Options;

/**
 * Interface, used by in-line (queue) reader plug-ins that are not handling a stream, but data in
 * {@link PluginReader#getWorkbook()}
 */
@InternalApi
public interface PluginInlineReadProcessor extends PluginBaseReader {

    /**
     * Initialization method
     *
     * @param workbook      Workbook instance where read data is placed
     * @param readerOptions Optional reader options
     * @param index         Optional index, e.g. for worksheet identification
     */
    void init(Workbook workbook, Options readerOptions, Integer index);
}
