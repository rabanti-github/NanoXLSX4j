//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces;

import ch.rabanti.nanoxlsx4j.utils.xml.XmlElement;

/**
 * Interface to represent complex text data that can be formatted.
 * <p>
 * Java equivalent of the {@code internal} .NET interface {@code NanoXLSX.Interfaces.IFormattableText}.
 * </p>
 *
 * @implNote Framework-internal interface. Not intended for use by external plugin implementors.
 * @see ISortedMap
 */
public interface IFormattableText {

    /**
     * Gets the main XML element of the formattable text.
     *
     * @return the {@link XmlElement} representing this text's XML structure
     */
    XmlElement getXmlElement();
}
