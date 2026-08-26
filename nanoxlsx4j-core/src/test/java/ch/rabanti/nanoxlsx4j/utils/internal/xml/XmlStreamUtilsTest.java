/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils.internal.xml;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.StringReader;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

class XmlStreamUtilsTest {

    @DisplayName("CreateInputFactory - NamespaceAware should be true")
    @Test
    void createInputFactoryNamespaceAwareShouldBeTrue() {
        XMLInputFactory factory = XmlStreamUtils.createInputFactory();

        assertTrue((Boolean) factory.getProperty(XMLInputFactory.IS_NAMESPACE_AWARE));
    }

    @DisplayName("CreateInputFactory - DTD support should be false")
    @Test
    void createInputFactoryDtdSupportShouldBeFalse() {
        XMLInputFactory factory = XmlStreamUtils.createInputFactory();

        assertFalse((Boolean) factory.getProperty(XMLInputFactory.SUPPORT_DTD));
    }

    @DisplayName("CreateInputFactory - External entity support should be false")
    @Test
    void createInputFactoryExternalEntitiesShouldBeFalse() {
        XMLInputFactory factory = XmlStreamUtils.createInputFactory();

        assertFalse((Boolean) factory.getProperty(XMLInputFactory.IS_SUPPORTING_EXTERNAL_ENTITIES));
    }

    @DisplayName("IsElement - Should return true when positioned on a matching start element (case-insensitive)")
    @ParameterizedTest
    @CsvSource({
            "root, root",
            "Root, root",
            "root, Root",
            "ROOT, root",
            "node, NODE"
    })
    void isElementMatch(String elementName, String localName) throws XMLStreamException {
        XMLStreamReader reader = createReaderAtFirstElement("<" + elementName + "/>");
        try {
            assertTrue(XmlStreamUtils.isElement(reader, localName));
        } finally {
            reader.close();
        }
    }

    @DisplayName("IsElement - Should return false when the local name does not match")
    @ParameterizedTest
    @CsvSource({
            "root, node",
            "child, root",
            "element, elements"
    })
    void isElementNoMatch(String elementName, String localName) throws XMLStreamException {
        XMLStreamReader reader = createReaderAtFirstElement("<" + elementName + "/>");
        try {
            assertFalse(XmlStreamUtils.isElement(reader, localName));
        } finally {
            reader.close();
        }
    }

    @DisplayName("IsElement - Should return false when positioned on an end element")
    @Test
    void isElementEndElement() throws XMLStreamException {
        XMLInputFactory factory = XmlStreamUtils.createInputFactory();
        XMLStreamReader reader = factory.createXMLStreamReader(new StringReader("<root></root>"));
        try {
            reader.next(); // positioned on <root> start element
            reader.next(); // positioned on </root> end element
            assertEquals(XMLStreamConstants.END_ELEMENT, reader.getEventType());
            assertFalse(XmlStreamUtils.isElement(reader, "root"));
        } finally {
            reader.close();
        }
    }

    @DisplayName("IsElement - Should strip namespace prefix and match only local name")
    @Test
    void isElementNamespacePrefix() throws XMLStreamException {
        XMLInputFactory factory = XmlStreamUtils.createInputFactory();
        XMLStreamReader reader = factory.createXMLStreamReader(
                new StringReader("<ns:root xmlns:ns=\"http://example.com\"/>"));
        try {
            reader.next();
            assertTrue(XmlStreamUtils.isElement(reader, "root"));
        } finally {
            reader.close();
        }
    }

    @DisplayName("ReadElementText - Should return empty string for a self-closing (empty) element")
    @Test
    void readElementTextEmptyElement() throws XMLStreamException {
        XMLStreamReader reader = createReaderAtFirstElement("<node/>");
        try {
            String result = XmlStreamUtils.readElementText(reader);
            assertEquals("", result);
        } finally {
            reader.close();
        }
    }

    @DisplayName("ReadElementText - Should return empty string for an open/close element with no content")
    @Test
    void readElementTextOpenCloseEmptyElement() throws XMLStreamException {
        XMLStreamReader reader = createReaderAtFirstElement("<node></node>");
        try {
            String result = XmlStreamUtils.readElementText(reader);
            assertEquals("", result);
        } finally {
            reader.close();
        }
    }

    @DisplayName("ReadElementText - Should return the text content of a leaf element")
    @ParameterizedTest
    @CsvSource({
            "'<node>hello</node>', hello",
            "'<node>123</node>', 123",
            "'<node>  spaces  </node>', '  spaces  '",
            "'<node>line1</node>', line1"
    })
    void readElementTextWithContent(String xml, String expected) throws XMLStreamException {
        XMLStreamReader reader = createReaderAtFirstElement(xml);
        try {
            String result = XmlStreamUtils.readElementText(reader);
            assertEquals(expected, result);
        } finally {
            reader.close();
        }
    }

    private static XMLStreamReader createReaderAtFirstElement(String xml) throws XMLStreamException {
        XMLInputFactory factory = XmlStreamUtils.createInputFactory();
        XMLStreamReader reader = factory.createXMLStreamReader(new StringReader(xml));
        reader.next();
        return reader;
    }
}
