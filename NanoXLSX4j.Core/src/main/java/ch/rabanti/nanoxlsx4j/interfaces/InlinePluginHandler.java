//STATUS: USAGE TO BE CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces;

import ch.rabanti.nanoxlsx4j.Workbook;

import java.io.InputStream;

/**
 * Functional interface for inline plugin callbacks during read operations.
 * <p>
 * Java equivalent of the .NET delegate type {@code Action<MemoryStream, Workbook, string, IOptions, int?>} used by
 * {@link ch.rabanti.nanoxlsx4j.interfaces.reader.IPluginBaseReader} and
 * {@link ch.rabanti.nanoxlsx4j.interfaces.reader.IPluginInlineReader}.
 * </p>
 */
@FunctionalInterface
public interface InlinePluginHandler {

    /**
     * Invoked by the framework to let an inline reader process a specific OOXML package part.
     *
     * @param stream    the raw XML data stream for the OOXML part
     * @param workbook  the active workbook instance
     * @param entryName the OOXML package entry name / relative path
     * @param options   optional reader options, or {@code null}
     * @param index     optional index (e.g. worksheet index), or {@code null}
     */
    void handle(
            InputStream stream, Workbook workbook, String entryName,
            IOptions options, Integer index
    );
}
