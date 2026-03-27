/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.internal.structures;

import ch.rabanti.nanoxlsx4j.interfaces.IFormattableText;
import ch.rabanti.nanoxlsx4j.utils.xml.XmlElement;

/**
 * Class to wrap unformatted strings as formattable text for the shared string table.
 * <p>
 * Java equivalent of the .NET internal class {@code NanoXLSX.Internal.Structures.PlainText}.
 * </p>
 *
 * @implNote Framework-internal class. Not intended for use outside the writer module.
 * @author Raphael Stoeckli
 */
public class PlainText implements IFormattableText {

    // ### C O N S T A N T S ###

    private static final String ITEM_TAG_NAME = "si";
    private static final String TEXT_TAG_NAME = "t";
    private static final String PRESERVE_ATTRIBUTE_NAME = "space";
    private static final String PRESERVE_ATTRIBUTE_PREFIX_NAME = "xml";
    private static final String PRESERVE_ATTRIBUTE_VALUE = "preserve";

    // ### P R I V A T E  F I E L D S ###

    private String value;

    // ### G E T T E R S  &  S E T T E R S ###

    /**
     * Sets the unformatted value (plain text)
     *
     * @param value plain text value
     */
    public void setValue(String value) {
        this.value = value;
    }

    // ### C O N S T R U C T O R S ###

    /**
     * Constructor with value assignment
     *
     * @param value Value to assign
     */
    public PlainText(String value) {
        this.value = value;
    }

    // ### M E T H O D S ###

    /**
     * Gets the main XML element of this plain text entry (interface implementation).
     *
     * @return {@link XmlElement} representing the shared string item ({@code <si>})
     * @implNote Full implementation depends on {@link XmlElement} being complete.
     * Currently returns a stub until the XML utility layer is migrated.
     */
    @Override
    public XmlElement getXmlElement() {
        // TODO: implement fully once XmlElement supports child elements and attributes
        return new XmlElement();
    }

    /**
     * Determines whether the specified object is equal to this instance.
     * Equality is based solely on the plain text {@link #value}.
     *
     * @param obj Object to compare
     * @return {@code true} if both objects represent the same text
     */
    @Override
    public boolean equals(Object obj) {
        if (this.value == null && obj == null) {
            return true;
        }
        if (!(obj instanceof PlainText)) {
            return false;
        }
        PlainText other = (PlainText) obj;
        if (this.value == null) {
            return other.value == null;
        }
        if (other.value == null) {
            return false;
        }
        return this.value.equals(other.value);
    }

    /**
     * Returns a hash code based on the plain text value
     *
     * @return hash code, or {@code 0} if the value is {@code null}
     */
    @Override
    public int hashCode() {
        if (this.value == null) {
            return 0;
        }
        return this.value.hashCode();
    }

}
