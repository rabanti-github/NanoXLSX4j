//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces.writer;

import ch.rabanti.nanoxlsx4j.enums.PasswordType;
import ch.rabanti.nanoxlsx4j.interfaces.IPassword;
import ch.rabanti.nanoxlsx4j.utils.xml.XmlAttribute;

/**
 * Interface used by specific writers that serialize password / protection information.
 * <p>
 * Java equivalent of the {@code public} .NET interface
 * {@code NanoXLSX.Interfaces.Writer.IPasswordWriter}.
 * </p>
 *
 * @see IPassword
 * @see PasswordType
 */
public interface IPasswordWriter extends IPassword {

    /**
     * Gets the protection target this password writer operates on.
     *
     * @return the {@link PasswordType}
     */
    PasswordType getType();

    /**
     * Initializes this password writer.
     *
     * @param type         the protection target
     * @param passwordHash the hash that will be written
     */
    void init(PasswordType type, String passwordHash);

    /**
     * Returns an iterable of XML attributes representing the password / protection element.
     *
     * @return the XML attributes to emit
     */
    Iterable<XmlAttribute> getAttributes();
}
