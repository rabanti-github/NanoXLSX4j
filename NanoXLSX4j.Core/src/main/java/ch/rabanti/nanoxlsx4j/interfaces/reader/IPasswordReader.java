//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces.reader;

import ch.rabanti.nanoxlsx4j.enums.PasswordType;
import ch.rabanti.nanoxlsx4j.interfaces.IOptions;
import ch.rabanti.nanoxlsx4j.interfaces.IPassword;
import ch.rabanti.nanoxlsx4j.reader.ReaderOptions;
import org.w3c.dom.Node;

/**
 * Interface used by password reader plug-ins.
 * <p>
 * Java equivalent of the {@code public} .NET interface {@code NanoXLSX.Interfaces.Reader.IPasswordReader}.
 * </p>
 *
 * @see IPassword
 * @see PasswordType
 */
public interface IPasswordReader extends IPassword {

    /**
     * Initializes this password reader.
     *
     * @param type          the protection target (workbook or worksheet)
     * @param readerOptions optional reader options
     */
    void init(PasswordType type, ReaderOptions readerOptions);

    /**
     * Reads the password-related attributes from the supplied XML node.
     *
     * @param node the XML node containing password / protection attributes
     */
    void readXmlAttributes(Node node);
}
