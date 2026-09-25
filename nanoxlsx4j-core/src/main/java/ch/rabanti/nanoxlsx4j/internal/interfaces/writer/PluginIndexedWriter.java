/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.writer;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.internal.interfaces.Plugin;

/**
 * Interface, used by classes to write XML content iteratively, using {@link #getCurrentIndex()}, at the end of the XLSX
 * creation process
 */
@InternalApi
public interface PluginIndexedWriter extends PluginWriter {

    /**
     * Gets the current index that should be used in {@link Plugin#execute()}, to identify the current action to
     * execute. The index will automatically be set during the iterative execution of the plug-in, from 0 to the
     * {@link #getMaxIndex()}
     *
     * @return Current index
     */
    int getCurrentIndex();

    /**
     * Sets the current index that should be used in {@link Plugin#execute()}, to identify the current action to
     * execute. The index will automatically be set during the iterative execution of the plug-in, from 0 to the
     * {@link #getMaxIndex()}
     *
     * @param index Current index
     */
    void setCurrentIndex(int index);

    /**
     * Current unique index to determine the package part to write. The index must correlate with a prior executed
     * {@link PluginPackageRegistry}. If null, no package part will be written
     *
     * @return Current package part index
     */
    String gerCurrentUniquePackagePartIndex();

    /**
     * Maximum inclusive (0-based) index for the iterator. A value of -1 indicates that no iterations are required
     *
     * @return Max index number
     */
    int getMaxIndex();

}
