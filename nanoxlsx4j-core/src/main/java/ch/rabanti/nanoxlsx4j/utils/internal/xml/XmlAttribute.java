/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils.internal.xml;

import ch.rabanti.nanoxlsx4j.utils.ParserUtils;

import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/**
 * Class representing an internally used XML attribute
 */
final class XmlAttribute {

    private final String name;
    private final String value;
    private final boolean hasPrefix;
    private final String prefix;

    /**
     * Gets the name of the attribute (without prefix)
     * @return Name of the attribute
     */
    public String getName() {
        return name;
    }

    /**
     * Gets the attribute value as string
     * @return Value of the attribute
     */
    public String getValue() {
        return value;
    }

    /**
     * Gets whether a prefix for the attribute was defined
     * @return True if defined, otherwise false
     */
    public boolean isHasPrefix() {
        return hasPrefix;
    }

    /**
     *Gets the prefix of the attribute. If not defined, the prefix will be an empty string
     * @return Prefix of the attribute
     */
    public String getPrefix() {
        return prefix;
    }

    /**
     * Constructor with all parameters
     * @param name Attribute name
     * @param value Attribute value
     * @param prefix Attribute prefix (do not pass null)
     */
    XmlAttribute(String name, String value, String prefix) {
        this.name = name;
        this.value = value;
        this.prefix = prefix;
        this.hasPrefix = !ParserUtils.isNullOrEmpty(prefix);
    }

    /**
     * Constructor with parameters (without prefix)
     * @param name Attribute name
     * @param value Attribute value
     */
    XmlAttribute(String name, String value) {
        this(name, value, "");
    }

    /**
     *  Method to create an attribute instance
     * @param name Attribute name
     * @param value Attribute value
     * @return Attribute instance
     */
    public static XmlAttribute createAttribute(String name, String value)
    {
        return new XmlAttribute(name, value);
    }

    /**
     *  Method to create an attribute instance with prefix
     * @param name Attribute name
     * @param value Attribute value
     * @param prefix Attribute prefix (do no pass null)
     * @return Attribute instance
     */
    public static XmlAttribute createAttribute(String name, String value, String prefix)
    {
        return new XmlAttribute(name, value, prefix);
    }

    /**
     *  Method to create an empty attribute instance
     * @param name Attribute name
     * @return Attribute instance
     */
    public static XmlAttribute createEmptyAttribute(String name)
    {
        return new XmlAttribute(name, "");
    }

    /**
     *  Method to create an attribute instance with prefix
     * @param name Attribute name
     * @param prefix Attribute prefix (do no pass null)
     * @return Attribute instance
     */
    public static XmlAttribute createEmptyAttribute(String name, String prefix)
    {
        return new XmlAttribute(name, "", prefix);
    }

    /**
     * Method to find an attribute in a given list by attribute name. It is assumed that there are no duplicates (attribute name)
     * @param name Attribute name
     * @param attributes List of attributes
     * @return Attribute that matches the name, or {@link Optional#empty()} if no attribute was found
     */
    public static Optional<XmlAttribute> findAttribute(String name, Set<XmlAttribute> attributes)
    {
        if (attributes == null || attributes.isEmpty())
        {
            return Optional.empty();
        }
        if (attributes.stream().noneMatch(a -> a.getName().equals(name)))
        {
            return Optional.empty();
        }
        return attributes.stream().filter(a -> a.getName().equals(name)).findFirst();
    }

    /**
     * Returns whether two instances are the same
     * @param o   the reference object with which to compare.
     * @return True if this instance and the other are the same
     */
    @Override
    public boolean equals(Object o) {
        if (!(o instanceof XmlAttribute that))
            return false;

        return Objects.equals(name, that.name) && Objects.equals(value, that.value) &&
                Objects.equals(prefix, that.prefix);
    }

    /**
     * Gets the hash code of the attribute
     * @return Hash Code
     */
    @Override
    public int hashCode() {
        int result = Objects.hashCode(name);
        result = 31 * result + Objects.hashCode(value);
        result = 31 * result + Objects.hashCode(prefix);
        return result;
    }
}
