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
import ch.rabanti.nanoxlsx4j.internal.interfaces.Plugin;

/**
 * Interface, used by XML reader classes
 */
@InternalApi
public interface PluginReader extends Plugin {

    /**
     * Gets the workbook instance, defined by the constructor
     *
     * @return Workbook instance
     */
    Workbook getWorkbook();

    /**
     * Replaces the workbook instance, defined by the constructor
     *
     * @param workbook Workbook instance
     */
    void setWorkbook(Workbook workbook);

}
