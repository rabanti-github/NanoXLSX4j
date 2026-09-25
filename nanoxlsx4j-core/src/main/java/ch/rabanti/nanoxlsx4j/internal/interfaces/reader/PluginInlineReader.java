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
 * Interface, used by in-line (queue) plug-ins in XML reader classes
 */
@InternalApi
public interface PluginInlineReader extends PluginReader {

    /**
     * Gets the reference to the handler action, to be used for post operations in reader methods. Only relevant for
     * in-line plug-ins, therefore null for queue plug-ins.
     *
     * @return in-line plug-in handler, or {@code null} when no handler is configured
     */
    InlineReaderHandler getInlinePluginHandler();

    /**
     * Sets the reference to the handler action, to be used for post operations in reader methods. Only relevant for
     * in-line plug-ins, therefore null for queue plug-ins.
     *
     * @param inlinePluginHandler in-line plug-in handler, or {@code null} when no handler is configured
     */
    void setInlinePluginHandler(InlineReaderHandler inlinePluginHandler);

    /**
     * Initialization method
     *
     * @param stream        Stream containing the XML file to read
     * @param workbook      Workbook instance where read data is placed
     * @param readerOptions Optional reader options
     * @param index         Optional index
     */
    void init(InputStream stream, Workbook workbook, Options readerOptions, Integer index);

}
