//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces.writer;

import ch.rabanti.nanoxlsx4j.Worksheet;

/**
 * Interface used by worksheet writer plug-ins.
 * <p>
 * Java equivalent of the {@code internal} .NET interface
 * {@code NanoXLSX.Interfaces.Writer.IWorksheetWriter}.
 * </p>
 *
 * @implNote Framework-internal interface. Not intended for use by external plugin implementors.
 */
public interface IWorksheetWriter extends IPluginWriter {

    /**
     * Gets the currently active worksheet being written.
     *
     * @return the current worksheet
     */
    Worksheet getCurrentWorksheet();

    /**
     * Sets the currently active worksheet to write.
     *
     * @param worksheet the worksheet to activate
     */
    void setCurrentWorksheet(Worksheet worksheet);
}
