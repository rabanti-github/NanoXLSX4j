//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces;

/**
 * Base interface for classes that can be handled by extension packages (plug-ins).
 * <p>
 * Java equivalent of the {@code internal} .NET interface {@code NanoXLSX.Interfaces.IPlugin}.
 * </p>
 *
 * @implNote Framework-internal interface. Not intended for use by external plugin implementors.
 * @see ch.rabanti.nanoxlsx4j.interfaces.writer.IPluginWriter
 * @see ch.rabanti.nanoxlsx4j.interfaces.reader.IPluginReader
 */
public interface IPlugin {

    /**
     * Executes the plugin's primary operation.
     */
    void execute();
}
