/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.writer;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;

/**
 * Interface, used by writers
 */
@InternalApi
public interface BaseWriter extends WriteContext {

    /**
     * Gets the writer to write shared strings
     *
     * @return SharedString writer
     */
    SharedStringWriter getSharedStringWriter();

    /**
     * Sets the writer to write shared strings
     *
     * @param sharedStringWriter SharedString writer
     */
    void setSharedStringWriter(SharedStringWriter sharedStringWriter);
}
