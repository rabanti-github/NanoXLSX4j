//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces.writer;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.styles.StyleManager;

/**
 * Interface used by the central writer orchestrator.
 * <p>
 * Java equivalent of the {@code internal} .NET interface
 * {@code NanoXLSX.Interfaces.Writer.IBaseWriter}.
 * </p>
 */
public interface IBaseWriter {

    /**
     * Gets the workbook instance used by this writer.
     *
     * @return the workbook
     */
    Workbook getWorkbook();

    /**
     * Gets the style manager instance used by this writer.
     *
     * @return the style manager ({@code ch.rabanti.nanoxlsx4j.styles.StyleManager} once migrated)
     */
    StyleManager getStyles();

    /**
     * Gets the shared-string writer delegate.
     *
     * @return the active {@link ISharedStringWriter}
     */
    ISharedStringWriter getSharedStringWriter();

    /**
     * Sets the shared-string writer delegate.
     *
     * @param writer the {@link ISharedStringWriter} to use
     */
    void setSharedStringWriter(ISharedStringWriter writer);
}
