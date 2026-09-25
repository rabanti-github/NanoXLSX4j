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
import ch.rabanti.nanoxlsx4j.internal.interfaces.Plugin;

/**
 * Interface, used by in-line (queued) write processors that do not perform actual XML writing tasks, but performing
 * actions on {@link Workbook}
 */
@InternalApi
public interface PluginInlineWriteProcessor extends Plugin {

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
     * Initializing method for the processor
     *
     * @param context Context of the current writing operation.
     */
    void init(WriteContext context);
}
