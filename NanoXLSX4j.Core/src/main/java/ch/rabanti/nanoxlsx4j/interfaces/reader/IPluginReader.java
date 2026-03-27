//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces.reader;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.interfaces.IPlugin;

/**
 * Base interface used by XML reader plug-in classes.
 * <p>
 * Java equivalent of the {@code internal} .NET interface
 * {@code NanoXLSX.Interfaces.Reader.IPluginReader}.
 * </p>
 */
public interface IPluginReader extends IPlugin {

    /**
     * Gets the workbook instance into which read data is placed.
     *
     * @return the workbook
     */
    Workbook getWorkbook();

    /**
     * Sets the workbook instance for this reader.
     *
     * @param workbook the workbook to use
     */
    void setWorkbook(Workbook workbook);
}
