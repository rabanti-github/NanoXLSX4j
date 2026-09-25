/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.reader;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.enums.PasswordType;
import ch.rabanti.nanoxlsx4j.internal.interfaces.Password;
import ch.rabanti.nanoxlsx4j.options.ReaderOptions;
import org.w3c.dom.Element;

/**
 * Interface, used by password readers
 */
@InternalApi
public interface PasswordReader extends Password {

    /**
     * Method to initialize the password reader
     *
     * @param type Target type of the password writer
     * @param readerOptions Reader options
     */
    void init(PasswordType type, ReaderOptions readerOptions);

    /**
     * Reads the attributes of the passed XML element that contains password information
     *
     * @param element XML element
     */
    void readXmlAttributes(Element element);
}
