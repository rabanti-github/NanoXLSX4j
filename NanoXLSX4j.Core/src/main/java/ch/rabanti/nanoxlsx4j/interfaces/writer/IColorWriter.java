//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces.writer;

import ch.rabanti.nanoxlsx4j.colors.Color;
import ch.rabanti.nanoxlsx4j.utils.xml.XmlAttribute;

/**
 * Interface used by specific writers that provide color serialization.
 * <p>
 * Java equivalent of the {@code public} .NET interface
 * {@code NanoXLSX.Interfaces.Writer.IColorWriter}.
 * </p>
 */
public interface IColorWriter {

    /**
     * Returns the XML attribute name for the given color instance.
     *
     * @param color the color instance
     * @return the attribute name string
     */
    String getAttributeName(Color color);

    /**
     * Returns the XML attribute value for the given color instance.
     *
     * @param color the color instance
     * @return the attribute value string
     */
    String getAttributeValue(Color color);

    /**
     * Returns whether a tint attribute is applicable for the given color instance.
     *
     * @param color the color instance
     * @return {@code true} if a tint value should be written
     */
    boolean useTintAttribute(Color color);

    /**
     * Returns the tint value (in the range {@code -1} to {@code 1}) as a string,
     * if applicable (see {@link #useTintAttribute(Color)}).
     *
     * @param color the color instance
     * @return the tint value string
     */
    String getTintAttributeValue(Color color);

    /**
     * Returns all applicable XML attributes for the given color instance.
     *
     * @param color the color instance
     * @return an iterable of {@link XmlAttribute} values describing this color
     */
    Iterable<XmlAttribute> getAttributes(Color color);
}
