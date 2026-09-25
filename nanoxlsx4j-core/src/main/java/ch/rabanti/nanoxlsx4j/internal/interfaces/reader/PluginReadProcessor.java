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
 * Interface used by reader plug-ins that process workbook data without handling a stream.
 */
@InternalApi
public interface PluginReadProcessor extends PluginReader {

    /**
     * Gets the optional reader options
     *
     * @return Reader options
     */
    Options getOptions();

    /**
     * Sets the optional reader options
     *
     * @param options Reader options
     */
    void setOptions(Options options);

    /**
     * Gets the handler of in-line plug-ins used for post-processing in {@link #execute()}.
     *
     * @return in-line plug-in handler, or {@code null} when no handler is configured
     */
    InlineReadProcessorHandler getInlinePluginHandler();

    /**
     * Sets the handler of in-line plug-ins used for post-processing in {@link #execute()}.
     *
     * @param inlinePluginHandler in-line plug-in handler, or {@code null} to disable post-processing
     */
    void setInlinePluginHandler(InlineReadProcessorHandler inlinePluginHandler);

    /**
     * Initializes the processor.
     *
     * @param workbook            Workbook instance where read data is placed
     * @param readerOptions       Optional reader options
     * @param inlinePluginHandler Handler used for post operations in processor methods
     */
    void init(Workbook workbook, Options readerOptions, InlineReadProcessorHandler inlinePluginHandler);

}
