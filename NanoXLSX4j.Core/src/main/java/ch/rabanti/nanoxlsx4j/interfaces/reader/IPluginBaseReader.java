//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces.reader;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.interfaces.IOptions;
import ch.rabanti.nanoxlsx4j.interfaces.InlinePluginHandler;

import java.io.InputStream;

/**
 * Interface used by base plug-ins in XML reader classes.
 * <p>
 * Java equivalent of the {@code internal} .NET interface {@code NanoXLSX.Interfaces.Reader.IPluginBaseReader}.
 * </p>
 */
public interface IPluginBaseReader extends IPluginReader {

    /**
     * Gets the optional reader options for this reader.
     *
     * @return the {@link IOptions} instance, or {@code null}
     */
    IOptions getOptions();

    /**
     * Sets the optional reader options.
     *
     * @param options the {@link IOptions} to apply
     */
    void setOptions(IOptions options);

    /**
     * Gets the inline plugin handler used for post-processing after reading a part.
     *
     * @return the {@link InlinePluginHandler}, or {@code null}
     */
    InlinePluginHandler getInlinePluginHandler();

    /**
     * Sets the inline plugin handler.
     *
     * @param handler the {@link InlinePluginHandler} to use
     */
    void setInlinePluginHandler(InlinePluginHandler handler);

    /**
     * Initializes this reader.
     *
     * @param stream              the XML data stream for the OOXML part to read
     * @param workbook            the workbook instance where read data is placed
     *                            ({@code ch.rabanti.nanoxlsx4j.Workbook} once migrated)
     * @param readerOptions       optional reader options, or {@code null}
     * @param inlinePluginHandler reference to the inline plugin handler for post-operations
     */
    void init(
            InputStream stream, Workbook workbook, IOptions readerOptions,
            InlinePluginHandler inlinePluginHandler
    );
}
