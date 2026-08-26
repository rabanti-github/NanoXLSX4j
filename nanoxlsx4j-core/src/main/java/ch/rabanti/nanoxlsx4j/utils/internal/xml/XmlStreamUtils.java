/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils.internal.xml;

import javax.xml.XMLConstants;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

/**
 * Provides helpers for secure, forward-only StAX parsing used by XML readers.
 */
public final class XmlStreamUtils {

    private XmlStreamUtils() {
    }

    /**
     * Creates a namespace-aware XML input factory with DTD and external-entity processing disabled. Comments,
     * processing instructions, and whitespace outside simple element text remain stream events and are ignored by
     * the element-oriented helpers in this class.
     *
     * @return Configured XML input factory
     */
    public static XMLInputFactory createInputFactory() {
        XMLInputFactory factory = XMLInputFactory.newDefaultFactory();
        factory.setProperty(XMLInputFactory.IS_NAMESPACE_AWARE, true);
        factory.setProperty(XMLInputFactory.IS_COALESCING, true);
        factory.setProperty(XMLInputFactory.IS_REPLACING_ENTITY_REFERENCES, true);
        factory.setProperty(XMLInputFactory.IS_SUPPORTING_EXTERNAL_ENTITIES, false);
        factory.setProperty(XMLInputFactory.SUPPORT_DTD, false);
        factory.setProperty(XMLConstants.ACCESS_EXTERNAL_DTD, "");
        factory.setProperty(XMLConstants.ACCESS_EXTERNAL_SCHEMA, "");
        return factory;
    }

    /**
     * Returns whether the reader is positioned on a start element whose local name matches the supplied name,
     * ignoring case. Namespace prefixes are not considered.
     *
     * @param reader    XML stream reader
     * @param localName Local element name to match
     * @return {@code true} if the current event is a matching start element; otherwise {@code false}
     */
    public static boolean isElement(XMLStreamReader reader, String localName) {
        return reader.getEventType() == XMLStreamConstants.START_ELEMENT
                && reader.getLocalName().equalsIgnoreCase(localName);
    }

    /**
     * Reads the text content of a simple leaf element. The reader must be positioned on its start element and is
     * left positioned on the matching end element. Empty and self-closing elements return an empty string.
     *
     * @param reader XML stream reader positioned on a leaf start element
     * @return Text content of the element, or an empty string if it has no content
     * @throws XMLStreamException If the element contains a child element or the XML stream cannot be read
     */
    public static String readElementText(XMLStreamReader reader) throws XMLStreamException {
        return reader.getElementText();
    }
}
