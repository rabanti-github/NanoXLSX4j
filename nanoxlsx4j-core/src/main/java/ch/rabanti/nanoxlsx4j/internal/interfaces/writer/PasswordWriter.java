/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.writer;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.enums.PasswordType;
import ch.rabanti.nanoxlsx4j.internal.interfaces.Password;
import ch.rabanti.nanoxlsx4j.utils.internal.xml.XmlAttribute;

/**
 * Interface, used by specific writers that provides password handling
 */
@InternalApi
public interface PasswordWriter extends Password {

    /**
     * Gets the target type of the password
     */
    PasswordType getType();

    /**
     * Method to initialize the password writer
     *
     * @param type         Target type of the password writer
     * @param passwordHash Hash that will be written
     */
    void init(PasswordType type, String passwordHash);

    /**
     * Gets an Iterable of XML attributes
     *
     * @return Iterable of XML attributes
     */
    Iterable<XmlAttribute> getAttributes();
}
