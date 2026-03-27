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
 * Interface used by worksheet reader plug-ins.
 * <p>
 * Java equivalent of the {@code internal} .NET interface
 * {@code NanoXLSX.Interfaces.Reader.IWorksheetReader}.
 * </p>
 *
 * @implNote Framework-internal interface. Not intended for use by external plugin implementors.
 */
public interface IWorksheetReader extends IPluginBaseReader {

    /**
     * Gets the 1-based relationship ID (rId) of the worksheet currently being processed.
     *
     * @return the current worksheet rId
     */
    int getCurrentWorksheetID();

    /**
     * Sets the 1-based relationship ID (rId) of the worksheet to be processed.
     *
     * @param id the worksheet rId
     */
    void setCurrentWorksheetID(int id);

    /**
     * Gets the list of shared strings. The index of each entry corresponds to the
     * {@code si} index referenced by cell values in the worksheet XML.
     *
     * @return the shared-string list
     */
    List<String> getSharedStrings();

    /**
     * Sets the shared-string list used for resolving cell string references.
     *
     * @param sharedStrings the shared-string list
     */
    void setSharedStrings(List<String> sharedStrings);
}
