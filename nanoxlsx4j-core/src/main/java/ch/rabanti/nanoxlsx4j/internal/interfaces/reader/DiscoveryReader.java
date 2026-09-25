/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.reader;

import java.util.zip.ZipFile;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.internal.interfaces.Options;

/**
 * Interface for readers that discover package-level information before document readers execute.
 */
@InternalApi
public interface DiscoveryReader extends PluginReader {

    /**
     * Gets the reader options used for discovery validation.
     *
     * @return Reader options
     */
    Options getOptions();

    /**
     * Sets the reader options used for discovery validation.
     *
     * @param options Reader options
     */
    void setOptions(Options options);

    /**
     * Initializes discovery with a caller-owned ZIP archive.
     *
     * @param archive       Zip archive
     * @param readerOptions Reader options
     * @param workbook      Workbook instance
     */
    void Init(ZipFile archive, Workbook workbook, Options readerOptions);
}
