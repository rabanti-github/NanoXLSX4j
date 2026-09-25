/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.internal.interfaces.writer;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.colors.Color;
import ch.rabanti.nanoxlsx4j.utils.internal.xml.XmlAttribute;

/**
 * Interface, used by specific writers that provides color handling
 */
@InternalApi
public interface ColorWriter {

    /**
     * Gets the attribute name for the given color instance
     *
     * @param color Color instance
     * @return Attribute name
     */
    String getAttributeName(Color color);

    /**
     * Gets the attribute value for the given color instance
     *
     * @param color Color instance
     * @return Attribute value
     */
    String getAttributeValue(Color color);

    /**
     * Gets whether a tint value is used for the given color instance
     *
     * @param color Color instance
     * @return True if tint is used
     */
    boolean useTintAttribute(Color color);

    /**
     * Gets the tint value as string of the given color instance, if applicable (see {@link #useTintAttribute(Color)}})
     *
     * @param color Color instance
     * @return Tint value (-1 to 1) as string
     */
    String getTintAttributeValue(Color color);

    /**
     * Gets all applicable attributes of the given color instance
     *
     * @param color Color instance
     * @return Iterable of the applicable XmlAttribute values of the color
     */
    Iterable<XmlAttribute> getAttributes(Color color);
}
