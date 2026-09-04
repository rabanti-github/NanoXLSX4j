/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils.internal.xml;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;

/**
 * Class representing an internally used XML element / node
 */
final class XmlElement {

    private final boolean hasPrefix;
    private boolean hasNameSpaces;
    private boolean hasDefaultNameSpace;
    private boolean hasAttributes;
    private boolean hasInnerValue;
    private boolean hasChildren;
    private String innerValue;
    private String defaultXmlNsUri;
    private String prefix;
    private String name;
    private List<XmlElement> children;
    private Set<XmlAttribute> attributes;
    private Map<String, String> prefixNameSpaceMap;

    /**
     * Gets the prefix of the element. If not defined, the prefix will be an empty string
     *
     * @return Element prefix
     */
    public String getPrefix() {
        return prefix;
    }

    /**
     * Sets the prefix of the element. If not defined, the prefix will be an empty string
     *
     * @param prefix Element prefix
     */
    public void setPrefix(String prefix) {
        this.prefix = prefix;
    }

    /**
     * Gets the name of the element (without prefix)
     *
     * @return Element name
     */
    public String getName() {
        return name;
    }

    /**
     * Gets the list of child elements. If none, the list is null
     *
     * @return Child elements (can be null)
     */
    public List<XmlElement> getChildren() {
        return children;
    }

    /**
     * Gets the list of attributes of this element. If none, the list is null
     *
     * @return Attributes (can be null)
     */
    public Set<XmlAttribute> getAttributes() {
        return attributes;
    }

    /**
     * Gets the map of prefixes and corresponding name space URIs of this element
     *
     * @return Prefix map
     */
    public Map<String, String> getPrefixNameSpaceMap() {
        return prefixNameSpaceMap;
    }

    /**
     * Gets the inner value of the element
     *
     * @return Inner value as string
     */
    public String getInnerValue() {
        return innerValue;
    }

    /**
     * Set the inner value of the element
     *
     * @param innerValue Inner value as string
     */
    public void setInnerValue(String innerValue) {
        if (ParserUtils.isNullOrEmpty(innerValue)) {
            this.innerValue = null;
            this.hasInnerValue = false;
        } else {
            this.innerValue = innerValue;
            this.hasInnerValue = true;
        }
    }

    /**
     * Prefix of the element
     *
     * @param name   Name of the element
     * @param prefix Prefix of the element
     */
    XmlElement(String name, String prefix) {
        this.name = name;
        this.prefix = prefix;
        this.hasPrefix = !ParserUtils.isNullOrEmpty(prefix);
    }

    /**
     * Method to add a name space as element attribute. Make sure not to add 'xmlns' as prefix since this is usually
     * only the default name space and will be added implicitly when defined by
     * {@link XmlElement#addDefaultXmlNameSpace(String)}
     *
     * @param prefix        Prefix of the name space
     * @param rootNameSpace Root name space (usually 'xmlns'). This value can also be empty
     * @param uri           URI of the name space
     */
    void addNameSpaceAttribute(String prefix, String rootNameSpace, String uri) {
        if (ParserUtils.isNullOrEmpty(prefix) || ParserUtils.isNullOrEmpty(uri)) {
            return;
        }
        if (prefixNameSpaceMap == null) {
            prefixNameSpaceMap = new HashMap<>();
        }
        if (!prefixNameSpaceMap.containsKey(prefix)) {
            prefixNameSpaceMap.put(prefix, uri);
        }
        hasNameSpaces = true;
        addAttribute(prefix, uri, rootNameSpace);
    }

    /**
     * Method to add the default name space  URI of the current element.
     *
     * @param defaultXmlNsUri URI to be defined as default name space
     */
    void addDefaultXmlNameSpace(String defaultXmlNsUri) {
        this.defaultXmlNsUri = defaultXmlNsUri;
        hasDefaultNameSpace = true;
    }

    /**
     * Method to add an attribute to the element
     *
     * @param name  Attribute name
     * @param value Attribute value
     */
    void addAttribute(String name, String value) {
        addAttribute(name, value, "");
    }

    /**
     * Method to add an attribute to the element
     *
     * @param name   Attribute name
     * @param value  Attribute value
     * @param prefix Attribute prefix
     */
    void addAttribute(String name, String value, String prefix) {
        if (!hasAttributes) {
            attributes = new HashSet<>();
            hasAttributes = true;
        }
        attributes.add(XmlAttribute.createAttribute(name, value, prefix));
    }

    /**
     * Method to add an attribute to the element
     *
     * @param nullableAttribute Nullable attribute instance. If not defined (null), nothing will be added
     */
    void addAttribute(XmlAttribute nullableAttribute) {
        if (nullableAttribute == null) {
            return;
        }
        if (!hasAttributes) {
            attributes = new HashSet<>();
            hasAttributes = true;
        }
        attributes.add(nullableAttribute);
    }

    /**
     * Method to add an enumeration of attributes to the element
     *
     * @param attributes Iterable of Attributes to add. If null or empty, nothing will be added
     */
    void addAttributes(Iterable<XmlAttribute> attributes) {
        if (attributes == null || !attributes.iterator().hasNext()) {
            return;
        }
        if (!hasAttributes) {
            this.attributes = new HashSet<>();
            hasAttributes = true;
        }
        for (XmlAttribute attribute : attributes) {
            this.attributes.add(attribute);
        }
    }

    /**
     * Method to add A child element with one attribute to the current element
     *
     * @param name           Name of the child element
     * @param attributeName  Attribute name, added to the child element
     * @param attributeValue Attribute value, added to the child element
     * @return Instance of the added child element
     */
    XmlElement addChildElementWithAttribute(String name, String attributeName, String attributeValue) {
        return addChildElementWithAttribute(name, attributeName, attributeValue, "", "");
    }

    /**
     * Method to add A child element with one attribute to the current element
     *
     * @param name           Name of the child element
     * @param attributeName  Attribute name, added to the child element
     * @param attributeValue Attribute value, added to the child element
     * @param namePrefix     Prefix of the child element
     * @return Instance of the added child element
     */
    XmlElement addChildElementWithAttribute(
            String name, String attributeName, String attributeValue, String namePrefix) {
        return addChildElementWithAttribute(name, attributeName, attributeValue, namePrefix, "");
    }

    /**
     * Method to add A child element with one attribute to the current element
     *
     * @param name            Name of the child element
     * @param attributeName   Attribute name, added to the child element
     * @param attributeValue  Attribute value, added to the child element
     * @param namePrefix      Prefix of the child element
     * @param attributePrefix Prefix of the attribute, added to the child element
     * @return Instance of the added child element
     */
    XmlElement addChildElementWithAttribute(
            String name, String attributeName, String attributeValue, String namePrefix, String attributePrefix) {
        XmlElement childElement = createElementWithAttribute(
                name, attributeName, attributeValue, namePrefix, attributePrefix);
        addChildElement(childElement);
        return childElement;
    }

    /**
     * Method to add A child element with an inner value
     *
     * @param name       Name of the child element
     * @param innerValue Inner (text) value of the child element
     * @return Instance of the added child element
     */
    XmlElement addChildElementWithValue(String name, String innerValue) {
        return addChildElementWithValue(name, innerValue, "");
    }

    /**
     * Method to add A child element with an inner value
     *
     * @param name       Name of the child element
     * @param innerValue Inner (text) value of the child element
     * @param prefix     Prefix of the child element
     * @return Instance of the added child element
     */
    XmlElement addChildElementWithValue(String name, String innerValue, String prefix) {
        if (ParserUtils.isNullOrEmpty(innerValue)) {
            return null; // Omit empty nodes
        }
        XmlElement childElement = createElement(name, prefix);
        childElement.setInnerValue(innerValue);
        addChildElement(childElement);
        return childElement;
    }

    /**
     * Method to add a child element to the current one
     *
     * @param name Name of the child element
     * @return Instance of the added child element
     */
    XmlElement addChildElement(String name) {
        return addChildElement(name, "");
    }

    /**
     * Method to add a child element to the current one
     *
     * @param name   Name of the child element
     * @param prefix Prefix of the child element
     * @return Instance of the added child element
     */
    XmlElement addChildElement(String name, String prefix) {
        XmlElement childElement = createElement(name, prefix);
        addChildElement(childElement);
        return childElement;
    }

    /**
     * Method to add a child element to the current one
     *
     * @param xmlElement Nullable child element instance. If null, nothing will be added
     */
    void addChildElement(XmlElement xmlElement) {
        if (xmlElement == null) {
            return;
        }
        if (!hasChildren) {
            this.children = new ArrayList<>();
            hasChildren = true;
        }
        this.children.add(xmlElement);
    }

    /**
     * Method to add an enumeration of child element to the current one
     *
     * @param xmlElements Iterable of child elements to be added. If null or empty, nothing will be added
     */
    void addChildElements(Iterable<XmlElement> xmlElements) {
        if (xmlElements == null || !xmlElements.iterator().hasNext()) {
            return;
        }
        if (!hasChildren) {
            this.children = new ArrayList<>();
            hasChildren = true;
        }
        xmlElements.forEach(this.children::add);
    }

    /**
     * Adds a child element before the first occurrence of the first matching ancestor name.
     *
     * @param xmlElement Nullable child element instance. If null, nothing will be added
     * @param ancestors  Names of possible ancestors, ordered by priority
     * @throws IOException Thrown if none of the specified ancestors exists
     */
    void addChildElementBefore(XmlElement xmlElement, String... ancestors) throws IOException {
        if (xmlElement == null) {
            return;
        }
        if (hasChildren && ancestors != null) {
            for (int ancestorIndex = 0; ancestorIndex < ancestors.length; ancestorIndex++) {
                for (int childIndex = 0; childIndex < this.children.size(); childIndex++) {
                    if (this.children.get(childIndex).getName().equals(ancestors[ancestorIndex])) {
                        children.add(childIndex, xmlElement);
                        return;
                    }
                }
            }
        }
        throw new IOException("None of the specified ancestor elements were found");
    }

    /**
     * Adds a child element after the last occurrence of the first matching successor name.
     *
     * @param xmlElement Nullable child element instance. If null, nothing will be added
     * @param successors Names of possible successors, ordered by priority
     * @throws IOException Thrown if none of the specified successors exists
     */
    void addChildElementAfter(XmlElement xmlElement, String... successors) throws IOException {
        if (xmlElement == null) {
            return;
        }
        if (hasChildren && successors != null) {
            for (int successorIndex = 0; successorIndex < successors.length; successorIndex++) {
                for (int childIndex = this.children.size() - 1; childIndex >= 0; childIndex--) {
                    if (this.children.get(childIndex).getName().equals(successors[successorIndex])) {
                        this.children.add(childIndex + 1, xmlElement);
                        return;
                    }
                }
            }
        }
        throw new IOException("None of the specified successor elements were found");
    }

    /**
     * Transforms this custom XML element and its children into a standard namespace-aware DOM document.
     *
     * @return New DOM document representing the hierarchical XML structure
     * @throws IllegalStateException If a DOM document builder cannot be created
     */
    public Document transformToDocument() {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newDefaultInstance();
        factory.setNamespaceAware(true);
        try {
            Document document = factory.newDocumentBuilder().newDocument();
            Element rootElement;
            if (hasDefaultNameSpace) {
                rootElement = createXmlElement(document, this, Map.of(), defaultXmlNsUri);
            } else {
                rootElement = createXmlElement(document, this, Map.of());
            }
            document.appendChild(rootElement);
            return document;
        } catch (ParserConfigurationException exception) {
            throw new IllegalStateException("Could not create an XML document", exception);
        }
    }

    /**
     * Streams this custom XML element and its children directly to an XML stream writer.
     *
     * @param writer Target XML stream writer
     * @throws XMLStreamException If the writer cannot write the XML structure
     */
    void writeTo(XMLStreamWriter writer) throws XMLStreamException {
        writeTo(writer, null);
    }

    /**
     * Streams this custom XML element and its children directly to an XML stream writer without creating an
     * intermediate DOM tree or XML string.
     *
     * @param writer    Target XML stream writer
     * @param defaultNs Default namespace URI inherited from the parent element, or {@code null} at the root
     * @throws XMLStreamException If the writer cannot write the XML structure
     */
    private void writeTo(XMLStreamWriter writer, String defaultNs) throws XMLStreamException {
        String elementDefaultNs = hasDefaultNameSpace ? defaultXmlNsUri : defaultNs;

        if (hasPrefix) {
            writer.writeStartElement(prefix, name, resolveNamespace(prefix, prefixNameSpaceMap, writer));
        } else if (ParserUtils.isNullOrEmpty(elementDefaultNs)) {
            writer.writeStartElement(name);
        } else {
            writer.writeStartElement(XMLConstants.DEFAULT_NS_PREFIX, name, elementDefaultNs);
        }

        if (hasDefaultNameSpace) {
            writer.writeDefaultNamespace(namespaceOrEmpty(defaultXmlNsUri));
        }

        if (hasNameSpaces && prefixNameSpaceMap != null) {
            for (Map.Entry<String, String> namespace : prefixNameSpaceMap.entrySet()) {
                if (XMLConstants.XMLNS_ATTRIBUTE.equals(namespace.getKey())) {
                    continue;
                }
                writer.writeNamespace(namespace.getKey(), namespace.getValue());
            }
        }

        if (hasAttributes) {
            for (XmlAttribute attribute : attributes) {
                writeAttribute(writer, attribute);
            }
        }

        if (hasInnerValue) {
            writer.writeCharacters(innerValue);
        }

        if (hasChildren) {
            for (XmlElement child : children) {
                child.writeTo(writer, elementDefaultNs);
            }
        }

        writer.writeEndElement();
    }

    /**
     * Method to find XML child elements, based of its name. Name space and hierarchy is not considered as exclusion
     * parameters
     *
     * @param name Name of the target element or elements
     * @return Iterable of XML element. If no element was found, an empty IEnumerable will be returned
     */
    public Iterable<XmlElement> findChildElementsByName(String name) {
        if (!hasChildren) {
            return new ArrayList<>();
        }
        List<XmlElement> result = new ArrayList<>();
        for (XmlElement child : this.children) {
            if (child.getName().equals(name)) {
                result.add(child);
            }
            child.findChildElementsByName(name).forEach(result::add);
        }
        return result;
    }

    /**
     * Method to find XML child elements, based of its name, an attribute name. Name space and hierarchy is not
     * considered as exclusion parameters
     *
     * @param elementName   Name of the target element or elements
     * @param attributeName Name of the target attribute, present in the XML element
     * @return Iterable of XML element. If no element was found, an empty IEnumerable will be returned
     */
    public Iterable<XmlElement> findChildElementsByNameAndAttribute(String elementName, String attributeName) {
        return findChildElementsByNameAndAttribute(elementName, attributeName, null, false);
    }

    /**
     * Method to find XML child elements, based of its name, an attribute name and value. Name space and hierarchy is
     * not considered as exclusion parameters
     *
     * @param elementName    Name of the target element or elements
     * @param attributeName  Name of the target attribute, present in the XML element
     * @param attributeValue Value of the XML attribute, present in the XML element
     * @return Iterable of XML element. If no element was found, an empty IEnumerable will be returned
     */
    public Iterable<XmlElement> findChildElementsByNameAndAttribute(
            String elementName, String attributeName, String attributeValue) {
        return findChildElementsByNameAndAttribute(elementName, attributeName, attributeValue, true);
    }

    /**
     * Method to find XML child elements, based of its name, an attribute name and optional value. Name space and
     * hierarchy is not considered as exclusion parameters
     *
     * @param elementName    Name of the target element or elements
     * @param attributeName  Name of the target attribute, present in the XML element
     * @param attributeValue Value of the XML attribute, present in the XML element
     * @param useValue       If true, the attribute name and value will be considered, otherwise only the attribute
     *                       name
     * @return Iterable of XML element. If no element was found, an empty IEnumerable will be returned
     */
    private Iterable<XmlElement> findChildElementsByNameAndAttribute(
            String elementName, String attributeName, String attributeValue, boolean useValue) {
        if (!hasChildren) {
            return new ArrayList<>();
        }
        List<XmlElement> result = new ArrayList<>();
        for (XmlElement child : this.children) {
            if (child.getName().equals(elementName) && child.hasAttributes) {
                Optional<XmlAttribute> attribute = XmlAttribute.findAttribute(attributeName, child.attributes);
                if (attribute.isPresent()) {
                    if (!useValue || attribute.get().getValue().equals(attributeValue)) {
                        result.add(child);
                    }
                }
            }
            child.findChildElementsByNameAndAttribute(elementName, attributeName, attributeValue, useValue)
                    .forEach(result::add);
        }
        return result;
    }

    /**
     * Method to create an XML element
     *
     * @param name Name of the element
     * @return Element instance
     */
    static XmlElement createElement(String name) {
        return createElement(name, "");
    }

    /**
     * Method to create an XML element
     *
     * @param name   Name of the element
     * @param prefix Prefix of the element
     * @return Element instance
     */
    static XmlElement createElement(String name, String prefix) {
        return new XmlElement(name, prefix);
    }

    /**
     * Method to create an XML element with one attribute
     *
     * @param name           Name of the element
     * @param attributeName  Attribute name
     * @param attributeValue Attribute value
     * @return Element instance
     */
    static XmlElement createElementWithAttribute(String name, String attributeName, String attributeValue) {
        return createElementWithAttribute(name, attributeName, attributeValue, "", "");
    }

    /**
     * Method to create an XML element with one attribute
     *
     * @param name           Name of the element
     * @param attributeName  Attribute name
     * @param attributeValue Attribute value
     * @param namePrefix     Prefix of the name
     * @return Element instance
     */
    static XmlElement createElementWithAttribute(
            String name, String attributeName, String attributeValue, String namePrefix) {
        return createElementWithAttribute(name, attributeName, attributeValue, namePrefix, "");
    }

    /**
     * Method to create an XML element with one attribute
     *
     * @param name            Name of the element
     * @param attributeName   Attribute name
     * @param attributeValue  Attribute value
     * @param namePrefix      Prefix of the name
     * @param attributePrefix Prefix of the attribute
     * @return Element instance
     */
    static XmlElement createElementWithAttribute(
            String name, String attributeName, String attributeValue, String namePrefix, String attributePrefix) {
        XmlElement element = new XmlElement(name, namePrefix);
        element.attributes = new HashSet<>();
        element.attributes.add(XmlAttribute.createAttribute(attributeName, attributeValue, attributePrefix));
        element.hasAttributes = true;
        return element;
    }

    /**
     * Recursively creates a standard DOM element from a custom XML element without an inherited default namespace.
     *
     * @param document         DOM document to which the element belongs
     * @param customElement    Custom XML element to convert
     * @param parentNamespaces Namespace prefixes inherited from the parent element
     * @return DOM element representing the custom element
     */
    private static Element createXmlElement(
            Document document, XmlElement customElement, Map<String, String> parentNamespaces) {
        return createXmlElement(document, customElement, parentNamespaces, null);
    }

    /**
     * Recursively creates a standard namespace-aware DOM element from a custom XML element.
     *
     * @param document         DOM document to which the element belongs
     * @param customElement    Custom XML element to convert
     * @param parentNamespaces Namespace prefixes inherited from the parent element
     * @param defaultXmlNsUri  Default namespace URI inherited from the parent element, or {@code null}
     * @return DOM element representing the custom element
     */
    private static Element createXmlElement(
            Document document,
            XmlElement customElement,
            Map<String, String> parentNamespaces,
            String defaultXmlNsUri
    ) {
        Map<String, String> namespaces = parentNamespaces;
        if (customElement.hasNameSpaces && customElement.prefixNameSpaceMap != null) {
            namespaces = new HashMap<>(parentNamespaces);
            namespaces.putAll(customElement.prefixNameSpaceMap);
        }

        String elementDefaultNs = customElement.hasDefaultNameSpace
                ? customElement.defaultXmlNsUri
                : defaultXmlNsUri;
        String qualifiedName = customElement.hasPrefix
                ? customElement.prefix + ':' + customElement.name
                : customElement.name;
        String elementNamespace = customElement.hasPrefix
                ? namespaces.get(customElement.prefix)
                : elementDefaultNs;
        Element xmlElement = document.createElementNS(elementNamespace, qualifiedName);

        if (customElement.hasDefaultNameSpace) {
            xmlElement.setAttributeNS(
                    XMLConstants.XMLNS_ATTRIBUTE_NS_URI,
                    XMLConstants.XMLNS_ATTRIBUTE,
                    namespaceOrEmpty(customElement.defaultXmlNsUri)
            );
        }
        if (customElement.hasNameSpaces && customElement.prefixNameSpaceMap != null) {
            for (Map.Entry<String, String> namespace : customElement.prefixNameSpaceMap.entrySet()) {
                if (XMLConstants.XMLNS_ATTRIBUTE.equals(namespace.getKey())) {
                    continue;
                }
                xmlElement.setAttributeNS(
                        XMLConstants.XMLNS_ATTRIBUTE_NS_URI,
                        XMLConstants.XMLNS_ATTRIBUTE + ':' + namespace.getKey(),
                        namespace.getValue()
                );
            }
        }

        if (customElement.hasAttributes) {
            for (XmlAttribute attribute : customElement.attributes) {
                setAttribute(xmlElement, attribute, namespaces);
            }
        }

        if (customElement.hasInnerValue) {
            xmlElement.setTextContent(customElement.innerValue);
        }

        if (customElement.hasChildren) {
            for (XmlElement child : customElement.children) {
                xmlElement.appendChild(createXmlElement(document, child, namespaces, elementDefaultNs));
            }
        }
        return xmlElement;
    }

    /**
     * Adds a custom attribute to a namespace-aware DOM element.
     *
     * @param element    Target DOM element
     * @param attribute  Custom attribute to add
     * @param namespaces Namespace prefixes visible on the element
     */
    private static void setAttribute(
            Element element, XmlAttribute attribute, Map<String, String> namespaces) {
        String attributeName = attribute.getName();
        String attributePrefix = attribute.getPrefix();
        if (attribute.isHasPrefix()) {
            if (XMLConstants.XMLNS_ATTRIBUTE.equals(attributePrefix)) {
                element.setAttributeNS(
                        XMLConstants.XMLNS_ATTRIBUTE_NS_URI,
                        XMLConstants.XMLNS_ATTRIBUTE + ':' + attributeName,
                        attribute.getValue()
                );
                return;
            }
            element.setAttributeNS(
                    namespaces.get(attributePrefix),
                    attributePrefix + ':' + attributeName,
                    attribute.getValue()
            );
            return;
        }

        int colonIndex = attributeName.indexOf(':');
        if (colonIndex > 0) {
            String implicitPrefix = attributeName.substring(0, colonIndex);
            String localName = attributeName.substring(colonIndex + 1);
            if (XMLConstants.XMLNS_ATTRIBUTE.equals(implicitPrefix)) {
                element.setAttributeNS(
                        XMLConstants.XMLNS_ATTRIBUTE_NS_URI, attributeName, attribute.getValue());
            } else {
                element.setAttributeNS(
                        namespaces.get(implicitPrefix),
                        implicitPrefix + ':' + localName,
                        attribute.getValue()
                );
            }
        } else if (XMLConstants.XMLNS_ATTRIBUTE.equals(attributeName)) {
            element.setAttributeNS(
                    XMLConstants.XMLNS_ATTRIBUTE_NS_URI,
                    XMLConstants.XMLNS_ATTRIBUTE,
                    attribute.getValue()
            );
        } else {
            element.setAttribute(attributeName, attribute.getValue());
        }
    }

    /**
     * Writes a custom attribute to a streaming XML writer.
     *
     * @param writer    Target XML stream writer
     * @param attribute Custom attribute to write
     * @throws XMLStreamException If the writer cannot write the attribute
     */
    private void writeAttribute(XMLStreamWriter writer, XmlAttribute attribute) throws XMLStreamException {
        String attributeName = attribute.getName();
        if (attribute.isHasPrefix()) {
            if (XMLConstants.XMLNS_ATTRIBUTE.equals(attribute.getPrefix())) {
                return;
            }
            writer.writeAttribute(
                    attribute.getPrefix(),
                    resolveNamespace(attribute.getPrefix(), prefixNameSpaceMap, writer),
                    attributeName,
                    attribute.getValue()
            );
            return;
        }

        int colonIndex = attributeName.indexOf(':');
        if (colonIndex > 0) {
            String implicitPrefix = attributeName.substring(0, colonIndex);
            if (XMLConstants.XMLNS_ATTRIBUTE.equals(implicitPrefix)) {
                return;
            }
            writer.writeAttribute(
                    implicitPrefix,
                    resolveNamespace(implicitPrefix, prefixNameSpaceMap, writer),
                    attributeName.substring(colonIndex + 1),
                    attribute.getValue()
            );
        } else if (!XMLConstants.XMLNS_ATTRIBUTE.equals(attributeName)) {
            writer.writeAttribute(attributeName, attribute.getValue());
        }
    }

    /**
     * Resolves a namespace prefix from declarations on the current element or the writer's inherited context.
     *
     * @param prefix     Namespace prefix to resolve
     * @param namespaces Namespace declarations on the current element
     * @param writer     XML writer providing the inherited namespace context
     * @return Namespace URI, or the empty namespace URI when the prefix is not mapped
     */
    private static String resolveNamespace(
            String prefix, Map<String, String> namespaces, XMLStreamWriter writer) {
        if (namespaces != null && namespaces.containsKey(prefix)) {
            return namespaces.get(prefix);
        }
        if (writer.getNamespaceContext() != null) {
            String namespaceUri = writer.getNamespaceContext().getNamespaceURI(prefix);
            if (namespaceUri != null) {
                return namespaceUri;
            }
        }
        return XMLConstants.NULL_NS_URI;
    }

    /**
     * Converts a nullable namespace URI to the empty namespace URI used by the JDK XML APIs.
     *
     * @param namespaceUri Namespace URI
     * @return Original URI, or an empty string when it is {@code null}
     */
    private static String namespaceOrEmpty(String namespaceUri) {
        return namespaceUri == null ? XMLConstants.NULL_NS_URI : namespaceUri;
    }

}
