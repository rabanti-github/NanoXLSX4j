package ch.rabanti.nanoxlsx4j.utils.xml;

/**
 * Represents an internally used XML attribute (name–value pair with an optional prefix).
 * <p>
 * Java equivalent of the {@code NanoXLSX.Utils.Xml.XmlAttribute} struct in the .NET reference.
 * </p>
 *
 * <h3>Factory methods</h3>
 * <ul>
 *   <li>{@link #of(String, String)} — attribute without prefix</li>
 *   <li>{@link #of(String, String, String)} — attribute with prefix</li>
 * </ul>
 */
public final class XmlAttribute {

    private final String name;
    private final String value;
    private final String prefix;

    // -------------------------------------------------------------------------
    // Constructor
    // -------------------------------------------------------------------------

    private XmlAttribute(String name, String value, String prefix) {
        this.name   = name   != null ? name   : "";
        this.value  = value  != null ? value  : "";
        this.prefix = prefix != null ? prefix : "";
    }

    // -------------------------------------------------------------------------
    // Accessors
    // -------------------------------------------------------------------------

    /** Returns the attribute name (without prefix). */
    public String getName()   { return name; }

    /** Returns the attribute value as a string. */
    public String getValue()  { return value; }

    /** Returns the attribute prefix, or an empty string if none was defined. */
    public String getPrefix() { return prefix; }

    /** Returns {@code true} if a non-empty prefix was defined. */
    public boolean hasPrefix() { return !prefix.isEmpty(); }

    // -------------------------------------------------------------------------
    // Factory methods
    // -------------------------------------------------------------------------

    /** Creates an attribute without a prefix. */
    public static XmlAttribute of(String name, String value) {
        return new XmlAttribute(name, value, "");
    }

    /** Creates an attribute with the given prefix. */
    public static XmlAttribute of(String name, String value, String prefix) {
        return new XmlAttribute(name, value, prefix);
    }

    // -------------------------------------------------------------------------
    // Object overrides
    // -------------------------------------------------------------------------

    @Override
    public boolean equals(Object obj) {
        if (this == obj) return true;
        if (!(obj instanceof XmlAttribute)) return false;
        XmlAttribute other = (XmlAttribute) obj;
        return name.equals(other.name)
            && value.equals(other.value)
            && prefix.equals(other.prefix);
    }

    @Override
    public int hashCode() {
        int result = name.hashCode();
        result = 31 * result + value.hashCode();
        result = 31 * result + prefix.hashCode();
        return result;
    }

    @Override
    public String toString() {
        return hasPrefix() ? prefix + ":" + name + "=\"" + value + "\""
                           : name + "=\"" + value + "\"";
    }
}
