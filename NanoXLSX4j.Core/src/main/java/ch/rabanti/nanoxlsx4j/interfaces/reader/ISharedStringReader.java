//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces.reader;

import java.util.List;

/**
 * Interface used by shared-string reader plug-ins.
 * <p>
 * Java equivalent of the {@code internal} .NET interface
 * {@code NanoXLSX.Interfaces.Reader.ISharedStringReader}.
 * </p>
 *
 * @implNote Framework-internal interface. Not intended for use by external plugin implementors.
 */
public interface ISharedStringReader extends IPluginBaseReader {

    /**
     * Gets the resolved list of shared strings.
     * <p>
     * The position of each string in the list corresponds to its index as referenced by
     * cell values in the worksheet XML.
     * </p>
     *
     * @return the ordered list of shared strings
     */
    List<String> getSharedStrings();
}
