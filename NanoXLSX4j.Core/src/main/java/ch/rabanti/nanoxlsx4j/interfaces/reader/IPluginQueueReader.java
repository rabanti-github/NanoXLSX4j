//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces.reader;

/**
 * Interface used by XML queue reader plug-in classes.
 * <p>
 * Java equivalent of the {@code internal} .NET interface
 * {@code NanoXLSX.Interfaces.Reader.IPluginQueueReader}.
 * </p>
 *
 * @implNote Framework-internal marker interface. Not intended for use by external plugin implementors.
 */
public interface IPluginQueueReader extends IPluginBaseReader {
    // Marker extension — no additional methods.
}
