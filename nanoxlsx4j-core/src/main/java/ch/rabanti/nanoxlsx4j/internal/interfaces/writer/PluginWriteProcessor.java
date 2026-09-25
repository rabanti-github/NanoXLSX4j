/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.writer;

import java.util.function.BiConsumer;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.internal.interfaces.Plugin;

/**
 * Interface used by write processors that do not perform actual XML writing tasks, but perform actions on a
 * {@link Workbook}.
 */
@InternalApi
public interface PluginWriteProcessor extends Plugin {

    /**
     * Gets the write context, defined by the constructor
     *
     * @return Write context
     */
    WriteContext getWriteContext();

    /**
     * Replaces the write context, defined by the constructor
     *
     * @param writeContext Write context
     */
    void setWriteContext(WriteContext writeContext);

    /**
     * Gets the handler of in-line plug-ins used during processor execution.
     *
     * @return in-line plug-in handler, or {@code null} when no handler is configured
     */
    BiConsumer<WriteContext, String> getInlinePluginHandler();

    /**
     * Sets the handler of in-line plug-ins used during processor execution.
     *
     * @param inlinePluginHandler in-line plug-in handler, or {@code null} to disable preparation operations
     */
    void setInlinePluginHandler(BiConsumer<WriteContext, String> inlinePluginHandler);

    /**
     * Initializes the processor.
     *
     * @param context             Context of the current writing operation
     * @param inlinePluginHandler Handler used for preparation operations in processor methods
     */
    void init(WriteContext context, BiConsumer<WriteContext, String> inlinePluginHandler);

}
