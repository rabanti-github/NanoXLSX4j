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
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.StringReader;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.List;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamWriter;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;
import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.xml.sax.InputSource;

class XmlElementTest {

    @DisplayName("Constructor should correctly set Name and Prefix and leave properties null")
    @ParameterizedTest
    @CsvSource({
            "ElementName, prefix",
            "ElementName, ''",
            "AnotherElement, somePrefix"
    })
    void createXmlElementTest(String name, String prefix) {
        XmlElement element = XmlElement.createElement(name, prefix);

        assertEquals(name, element.getName());
        assertEquals(prefix, element.getPrefix());
        assertNull(element.getChildren());
        assertNull(element.getAttributes());
        assertNull(element.getPrefixNameSpaceMap());
    }

    @DisplayName("Prefix property should be get and set correctly")
    @ParameterizedTest
    @CsvSource({
            "initialPrefix, newPrefix",
            "'', nonEmptyPrefix"
    })
    void prefixPropertyTest(String initialPrefix, String newPrefix) {
        XmlElement element = XmlElement.createElement("TestElement", initialPrefix);
        element.setPrefix(newPrefix);

        assertEquals(newPrefix, element.getPrefix());
    }

    @DisplayName("InnerValue property should set value if non-empty; empty or null resets to null")
    @ParameterizedTest
    @CsvSource(value = {
            "'Some value', 'Some value'",
            "'', NULL",
            "NULL, NULL"
    }, nullValues = "NULL")
    void innerValuePropertyTest(String setValue, String expectedValue) {
        XmlElement element = XmlElement.createElement("TestElement");
        element.setInnerValue(setValue);

        assertEquals(expectedValue, element.getInnerValue());
    }

    @DisplayName("Children property should be null when no children have been added")
    @Test
    void childrenPropertyInitialTest() {
        XmlElement element = XmlElement.createElement("TestElement");

        assertNull(element.getChildren());
    }

    @DisplayName("Attributes property should be null when no attributes have been added")
    @Test
    void attributesPropertyInitialTest() {
        XmlElement element = XmlElement.createElement("TestElement");

        assertNull(element.getAttributes());
    }

    @DisplayName("PrefixNameSpaceMap property should be null when not set")
    @Test
    void prefixNameSpaceMapPropertyInitialTest() {
        XmlElement element = XmlElement.createElement("TestElement");

        assertNull(element.getPrefixNameSpaceMap());
    }

    @DisplayName("AddAttribute(string, string, string) should add a single attribute correctly")
    @ParameterizedTest
    @CsvSource({
            "attr1, value1, prefix1",
            "attr2, value2, ''"
    })
    void addAttributeStringMethodTest(String name, String value, String prefix) {
        XmlElement element = XmlElement.createElement("TestElement");
        element.addAttribute(name, value, prefix);

        assertNotNull(element.getAttributes());
        assertEquals(1, element.getAttributes().size());
        XmlAttribute attr = element.getAttributes().iterator().next();
        assertEquals(name, attr.getName());
        assertEquals(value, attr.getValue());
        assertEquals(prefix, attr.getPrefix());
    }

    @DisplayName("AddAttribute(XmlAttribute?) should add a valid attribute and ignore null values")
    @Test
    void addAttributeNullableAttributeTest() {
        XmlElement element = XmlElement.createElement("TestElement");
        XmlAttribute validAttribute = XmlAttribute.createAttribute("attrValid", "valueValid", "pfx");
        element.addAttribute(validAttribute);
        XmlAttribute nullAttribute = null;
        element.addAttribute(nullAttribute);

        assertNotNull(element.getAttributes());
        assertEquals(1, element.getAttributes().size());
        XmlAttribute attr = element.getAttributes().iterator().next();
        assertEquals("attrValid", attr.getName());
        assertEquals("valueValid", attr.getValue());
        assertEquals("pfx", attr.getPrefix());
    }

    @DisplayName("AddAttributes(IEnumerable<XmlAttribute>) should add multiple attributes, and ignore null/empty collections")
    @Test
    void addAttributesEnumerableTest() {
        XmlElement element = XmlElement.createElement("TestElement");
        List<XmlAttribute> attributesList = List.of(
                XmlAttribute.createAttribute("attrA", "valueA", "pfxA"),
                XmlAttribute.createAttribute("attrB", "valueB"));
        element.addAttributes(attributesList);

        assertNotNull(element.getAttributes());
        assertEquals(attributesList.size(), element.getAttributes().size());

        element.addAttributes(new ArrayList<>());
        assertEquals(attributesList.size(), element.getAttributes().size());
        element.addAttributes(null);
        assertEquals(attributesList.size(), element.getAttributes().size());
    }

    @DisplayName("AddNameSpaceAttribute should add namespace mapping and corresponding attribute when valid")
    @ParameterizedTest
    @CsvSource({
            "ns, xmlns, http://example.com/ns",
            "x, xmlns, http://example.org/x"
    })
    void addNameSpaceAttributeValidInputTest(String prefix, String rootNameSpace, String uri) {
        XmlElement element = XmlElement.createElement("TestElement", "t");
        element.addNameSpaceAttribute(prefix, rootNameSpace, uri);

        assertNotNull(element.getPrefixNameSpaceMap());
        assertTrue(element.getPrefixNameSpaceMap().containsKey(prefix));
        assertEquals(uri, element.getPrefixNameSpaceMap().get(prefix));

        assertNotNull(element.getAttributes());
        XmlAttribute nsAttribute = element.getAttributes().stream()
                .filter(attribute -> attribute.getName().equals(prefix))
                .findFirst()
                .orElse(null);
        assertNotNull(nsAttribute);
        assertEquals(uri, nsAttribute.getValue());
        assertEquals(rootNameSpace, nsAttribute.getPrefix());
    }

    @DisplayName("AddNameSpaceAttribute should ignore empty prefix or URI")
    @ParameterizedTest
    @CsvSource({
            "'', xmlns, http://example.com/ns",
            "ns, xmlns, ''",
            "'', xmlns, ''"
    })
    void addNameSpaceAttributeInvalidInputTest(String prefix, String rootNameSpace, String uri) {
        XmlElement element = XmlElement.createElement("TestElement", "t");
        element.addNameSpaceAttribute(prefix, rootNameSpace, uri);

        assertNull(element.getPrefixNameSpaceMap());
        assertNull(element.getAttributes());
    }

    @DisplayName("AddDefaultXmlNameSpace should set the default XML namespace for the element")
    @ParameterizedTest
    @ValueSource(strings = {"http://example.com/default", "http://example.org/ns"})
    void addDefaultXmlNameSpaceTest(String defaultUri) {
        XmlElement element = XmlElement.createElement("TestElement");
        element.addDefaultXmlNameSpace(defaultUri);
        Document doc = element.transformToDocument();

        assertNotNull(doc.getDocumentElement());
        assertEquals("TestElement", doc.getDocumentElement().getLocalName());
        assertEquals(defaultUri, doc.getDocumentElement().getNamespaceURI());
    }

    @DisplayName("AddChildElementWithAttribute should create a child with one attribute and add it to the parent's children")
    @ParameterizedTest
    @CsvSource("ChildName, attrName, attrValue, childPrefix, attrPrefix")
    void addChildElementWithAttributeTest(
            String childName,
            String attributeName,
            String attributeValue,
            String namePrefix,
            String attributePrefix) {
        XmlElement parent = XmlElement.createElement("Parent");
        XmlElement child = parent.addChildElementWithAttribute(
                childName, attributeName, attributeValue, namePrefix, attributePrefix);

        assertNotNull(child);
        assertNotNull(parent.getChildren());
        assertTrue(parent.getChildren().contains(child));

        assertNotNull(child.getAttributes());
        assertEquals(1, child.getAttributes().size());
        XmlAttribute attr = child.getAttributes().iterator().next();
        assertEquals(attributeName, attr.getName());
        assertEquals(attributeValue, attr.getValue());
        assertEquals(attributePrefix, attr.getPrefix());
    }

    @DisplayName("AddChildElementWithValue should create a child with inner value when provided; returns null for empty inner value")
    @ParameterizedTest
    @CsvSource({
            "ChildName, 'Inner Text', childPrefix, true",
            "ChildName, '', childPrefix, false"
    })
    void addChildElementWithValueTest(
            String childName, String innerValue, String prefix, boolean shouldBeAdded) {
        XmlElement parent = XmlElement.createElement("Parent");
        XmlElement child = parent.addChildElementWithValue(childName, innerValue, prefix);

        if (shouldBeAdded) {
            assertNotNull(child);
            assertNotNull(parent.getChildren());
            assertTrue(parent.getChildren().contains(child));
            assertEquals(innerValue, child.getInnerValue());
        } else {
            assertNull(child);
            assertNull(parent.getChildren());
        }
    }

    @DisplayName("AddChildElement(string, string) should create and add a child element")
    @ParameterizedTest
    @CsvSource({
            "ChildName, childPrefix",
            "AnotherChild, ''"
    })
    void addChildElementStringOverloadTest(String childName, String prefix) {
        XmlElement parent = XmlElement.createElement("Parent");
        XmlElement child = parent.addChildElement(childName, prefix);

        assertNotNull(child);
        assertNotNull(parent.getChildren());
        assertTrue(parent.getChildren().contains(child));
        assertEquals(childName, child.getName());
        assertEquals(prefix, child.getPrefix());
    }

    @DisplayName("AddChildElement(XmlElement) should add a non-null child and ignore null")
    @Test
    void addChildElementXmlElementOverloadTest() {
        XmlElement parent = XmlElement.createElement("Parent");
        XmlElement child = XmlElement.createElement("Child", "c");

        parent.addChildElement(child);

        assertNotNull(parent.getChildren());
        assertTrue(parent.getChildren().contains(child));
        int countAfterValid = parent.getChildren().size();

        parent.addChildElement((XmlElement) null);
        assertEquals(countAfterValid, parent.getChildren().size());
    }

    @DisplayName("AddChildElements(IEnumerable<XmlElement>) should add multiple children and ignore null or empty collections")
    @Test
    void addChildElementsEnumerableTest() {
        XmlElement parent = XmlElement.createElement("Parent");
        XmlElement child1 = XmlElement.createElement("Child1");
        XmlElement child2 = XmlElement.createElement("Child2");
        List<XmlElement> childrenList = List.of(child1, child2);

        parent.addChildElements(childrenList);

        assertNotNull(parent.getChildren());
        assertEquals(childrenList.size(), parent.getChildren().size());
        assertTrue(parent.getChildren().contains(child1));
        assertTrue(parent.getChildren().contains(child2));

        parent.addChildElements(new ArrayList<>());
        assertEquals(childrenList.size(), parent.getChildren().size());

        parent.addChildElements(null);
        assertEquals(childrenList.size(), parent.getChildren().size());
    }

    @DisplayName("AddChildElementBefore should insert before the first occurrence of the first matching ancestor")
    @Test
    void addChildElementBeforeTest() throws IOException {
        XmlElement parent = XmlElement.createElement("Parent");
        XmlElement firstAncestor = parent.addChildElement("Ancestor");
        XmlElement secondAncestor = parent.addChildElement("Ancestor");
        XmlElement child = XmlElement.createElement("Child");

        parent.addChildElementBefore(child, "Ancestor");

        assertEquals(3, parent.getChildren().size());
        assertSame(child, parent.getChildren().get(0));
        assertSame(firstAncestor, parent.getChildren().get(1));
        assertSame(secondAncestor, parent.getChildren().get(2));
    }

    @DisplayName("AddChildElementBefore should use ancestor names as ordered fallbacks")
    @Test
    void addChildElementBeforeFallbackTest() throws IOException {
        XmlElement parent = XmlElement.createElement("Parent");
        XmlElement first = parent.addChildElement("First");
        XmlElement fallbackAncestor = parent.addChildElement("FallbackAncestor");
        XmlElement child = XmlElement.createElement("Child");

        parent.addChildElementBefore(child, "MissingAncestor", "FallbackAncestor");

        assertEquals(3, parent.getChildren().size());
        assertSame(first, parent.getChildren().get(0));
        assertSame(child, parent.getChildren().get(1));
        assertSame(fallbackAncestor, parent.getChildren().get(2));
    }

    @DisplayName("AddChildElementAfter should insert after the last occurrence of the first matching successor")
    @Test
    void addChildElementAfterTest() throws IOException {
        XmlElement parent = XmlElement.createElement("Parent");
        XmlElement firstSuccessor = parent.addChildElement("Successor");
        XmlElement secondSuccessor = parent.addChildElement("Successor");
        XmlElement trailingChild = parent.addChildElement("TrailingChild");
        XmlElement child = XmlElement.createElement("Child");

        parent.addChildElementAfter(child, "Successor");

        assertEquals(4, parent.getChildren().size());
        assertSame(firstSuccessor, parent.getChildren().get(0));
        assertSame(secondSuccessor, parent.getChildren().get(1));
        assertSame(child, parent.getChildren().get(2));
        assertSame(trailingChild, parent.getChildren().get(3));
    }

    @DisplayName("AddChildElementAfter should use successor names as ordered fallbacks")
    @Test
    void addChildElementAfterFallbackTest() throws IOException {
        XmlElement parent = XmlElement.createElement("Parent");
        XmlElement fallbackSuccessor = parent.addChildElement("FallbackSuccessor");
        XmlElement last = parent.addChildElement("Last");
        XmlElement child = XmlElement.createElement("Child");

        parent.addChildElementAfter(child, "MissingSuccessor", "FallbackSuccessor");

        assertEquals(3, parent.getChildren().size());
        assertSame(fallbackSuccessor, parent.getChildren().get(0));
        assertSame(child, parent.getChildren().get(1));
        assertSame(last, parent.getChildren().get(2));
    }

    @DisplayName("Relative child insertion should throw IOException when no named sibling exists")
    @ParameterizedTest
    @CsvSource({"true, false", "true, true", "false, false", "false, true"})
    void addChildElementRelativeMissingSiblingTest(boolean insertBefore, boolean addUnrelatedChild) {
        XmlElement parent = XmlElement.createElement("Parent");
        if (addUnrelatedChild) {
            parent.addChildElement("Unrelated");
        }
        XmlElement child = XmlElement.createElement("Child");

        IOException exception = insertBefore
                ? assertThrows(IOException.class,
                        () -> parent.addChildElementBefore(child, "Missing", "AlsoMissing"))
                : assertThrows(IOException.class,
                        () -> parent.addChildElementAfter(child, "Missing", "AlsoMissing"));

        assertTrue(exception.getMessage().contains(insertBefore ? "ancestor" : "successor"));
        assertFalse(parent.getChildren() != null && parent.getChildren().contains(child));
    }

    @DisplayName("Relative child insertion should throw IOException when no sibling names are supplied")
    @ParameterizedTest
    @ValueSource(booleans = {true, false})
    void addChildElementRelativeMissingNamesTest(boolean insertBefore) {
        XmlElement parent = XmlElement.createElement("Parent");
        XmlElement child = XmlElement.createElement("Child");

        if (insertBefore) {
            assertThrows(IOException.class, () -> parent.addChildElementBefore(child));
            assertThrows(IOException.class, () -> parent.addChildElementBefore(child, (String[]) null));
        } else {
            assertThrows(IOException.class, () -> parent.addChildElementAfter(child));
            assertThrows(IOException.class, () -> parent.addChildElementAfter(child, (String[]) null));
        }
        assertNull(parent.getChildren());
    }

    @DisplayName("Relative child insertion should ignore a null child element")
    @ParameterizedTest
    @ValueSource(booleans = {true, false})
    void addChildElementRelativeNullChildTest(boolean insertBefore) throws IOException {
        XmlElement parent = XmlElement.createElement("Parent");
        XmlElement sibling = parent.addChildElement("Sibling");

        if (insertBefore) {
            parent.addChildElementBefore(null, "Sibling");
        } else {
            parent.addChildElementAfter(null, "Sibling");
        }

        assertEquals(1, parent.getChildren().size());
        assertSame(sibling, parent.getChildren().get(0));
    }

    @DisplayName("CreateElement should instantiate an element with the given name and optional prefix")
    @ParameterizedTest
    @CsvSource({
            "TestElement, prefix",
            "TestElement, ''",
            "AnotherElement, ns"
    })
    void createElementTest(String name, String prefix) {
        XmlElement element = XmlElement.createElement(name, prefix);

        assertNotNull(element);
        assertEquals(name, element.getName());
        assertEquals(prefix, element.getPrefix());
        assertNull(element.getAttributes());
        assertNull(element.getChildren());
        assertNull(element.getPrefixNameSpaceMap());
    }

    @DisplayName("CreateElementWithAttribute should instantiate an element with one attribute")
    @ParameterizedTest
    @CsvSource({
            "ElementWithAttr, attrName, attrValue, elemPrefix, attrPrefix",
            "ElementWithAttr, id, 123, '', ''"
    })
    void createElementWithAttributeTest(
            String name,
            String attributeName,
            String attributeValue,
            String namePrefix,
            String attributePrefix) {
        XmlElement element = XmlElement.createElementWithAttribute(
                name, attributeName, attributeValue, namePrefix, attributePrefix);

        assertNotNull(element);
        assertEquals(name, element.getName());
        assertEquals(namePrefix, element.getPrefix());
        assertNotNull(element.getAttributes());
        assertEquals(1, element.getAttributes().size());
        XmlAttribute attr = element.getAttributes().iterator().next();
        assertEquals(attributeName, attr.getName());
        assertEquals(attributeValue, attr.getValue());
        assertEquals(attributePrefix, attr.getPrefix());
    }

    @DisplayName("TransformToDocument should create an XmlDocument with correct hierarchy, attributes, and inner text, with and without default namespace")
    @ParameterizedTest
    @ValueSource(booleans = {true, false})
    void transformToDocumentTest(boolean useDefaultNamespace) {
        XmlElement root = XmlElement.createElement("Root");
        if (useDefaultNamespace) {
            root.addDefaultXmlNameSpace("http://example.com/ns");
        } else {
            // Namespace-aware JDK DOM rejects the reserved xmlns prefix as an element prefix.
            root.addNameSpaceAttribute("ns", "xmlns", "http://example.com/ns");
        }
        root.addAttribute("version", "1.0");

        XmlElement childWithAttr = useDefaultNamespace
                ? root.addChildElementWithAttribute("Child", "id", "123", "", "")
                : root.addChildElementWithAttribute("Child", "id", "123", "ns", "");
        childWithAttr.setInnerValue("ChildValue");

        Document doc = root.transformToDocument();

        assertNotNull(doc.getDocumentElement());
        assertEquals("Root", doc.getDocumentElement().getLocalName());
        String versionAttr = doc.getDocumentElement().getAttribute("version");
        assertEquals("1.0", versionAttr);

        assertTrue(doc.getDocumentElement().getChildNodes().getLength() >= 1,
                "The root element should have at least one child element.");

        Element childElement = findChildElementByLocalName(doc.getDocumentElement(), "Child");
        assertNotNull(childElement);
        if (useDefaultNamespace) {
            assertEquals("http://example.com/ns", childElement.getNamespaceURI());
        } else {
            assertEquals("http://example.com/ns", childElement.getNamespaceURI());
        }

        assertEquals("ChildValue", childElement.getTextContent());
        String childId = childElement.getAttribute("id");
        assertEquals("123", childId);
    }

    @DisplayName("TransformToDocument should register non-xmlns prefix namespaces so that prefixed child elements resolve to the correct namespace URI")
    @Test
    void transformToDocumentWithPrefixNamespaceTest() {
        XmlElement root = XmlElement.createElement("Root");
        root.addDefaultXmlNameSpace("http://example.com/ns");
        root.addNameSpaceAttribute(
                "x14ac", "xmlns", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
        root.addChildElement("Child", "x14ac");

        Document doc = root.transformToDocument();

        Element childElement = (Element) doc.getDocumentElement().getChildNodes().item(0);
        assertEquals("x14ac", childElement.getPrefix());
        assertEquals("Child", childElement.getLocalName());
        assertEquals(
                "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
                childElement.getNamespaceURI());
    }

    @DisplayName("TransformToDocument should create prefixed attributes with the correct namespace URI and value")
    @Test
    void transformToDocumentWithPrefixedAttributeTest() {
        XmlElement root = XmlElement.createElement("Root");
        root.addDefaultXmlNameSpace("http://example.com/ns");
        root.addNameSpaceAttribute(
                "x14ac", "xmlns", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
        root.addAttribute(XmlAttribute.createAttribute("dyDescent", "0.25", "x14ac"));

        Document doc = root.transformToDocument();

        Attr attr = doc.getDocumentElement().getAttributeNodeNS(
                "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac", "dyDescent");
        assertNotNull(attr);
        assertEquals("0.25", attr.getValue());
        assertEquals("x14ac", attr.getPrefix());
    }

    @DisplayName("WriteTo should emit a default namespace declaration exactly once on the root")
    @Test
    void writeToDefaultNamespaceTest() throws Exception {
        XmlElement root = XmlElement.createElement("Root");
        root.addDefaultXmlNameSpace("http://example.com/ns");
        root.addAttribute("version", "1.0");

        String xml = serializeWriteTo(root);

        Document doc = parseXml(xml);
        assertEquals("Root", doc.getDocumentElement().getLocalName());
        assertEquals("http://example.com/ns", doc.getDocumentElement().getNamespaceURI());
        assertEquals("1.0", doc.getDocumentElement().getAttribute("version"));
        int xmlnsCount = 0;
        NamedNodeMap attributes = doc.getDocumentElement().getAttributes();
        for (int index = 0; index < attributes.getLength(); index++) {
            Node attribute = attributes.item(index);
            if (attribute.getNodeName().equals("xmlns") || "xmlns".equals(attribute.getPrefix())) {
                xmlnsCount++;
            }
        }
        assertEquals(1, xmlnsCount);
    }

    @DisplayName("WriteTo should emit each prefix-namespace declaration exactly once and resolve prefixed attributes")
    @Test
    void writeToPrefixNamespacesTest() throws Exception {
        XmlElement root = XmlElement.createElement("worksheet");
        root.addDefaultXmlNameSpace("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        root.addNameSpaceAttribute(
                "mc", "xmlns", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        root.addNameSpaceAttribute(
                "x14ac", "xmlns", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
        root.addAttribute(XmlAttribute.createAttribute("dyDescent", "0.25", "x14ac"));

        String xml = serializeWriteTo(root);

        Document doc = parseXml(xml);
        assertEquals(
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                doc.getDocumentElement().getNamespaceURI());

        int mcCount = 0;
        int x14acCount = 0;
        NamedNodeMap attributes = doc.getDocumentElement().getAttributes();
        for (int index = 0; index < attributes.getLength(); index++) {
            Node attribute = attributes.item(index);
            if ("xmlns".equals(attribute.getPrefix()) && attribute.getLocalName().equals("mc")) {
                mcCount++;
            }
            if ("xmlns".equals(attribute.getPrefix()) && attribute.getLocalName().equals("x14ac")) {
                x14acCount++;
            }
        }
        assertEquals(1, mcCount);
        assertEquals(1, x14acCount);

        Attr dy = doc.getDocumentElement().getAttributeNodeNS(
                "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac", "dyDescent");
        assertNotNull(dy);
        assertEquals("0.25", dy.getValue());
    }

    @DisplayName("WriteTo should skip xmlns-keyed namespace map entries and not emit an xmlns:xmlns declaration")
    @Test
    void writeToSkipsXmlnsKeyedNamespaceTest() throws Exception {
        XmlElement root = XmlElement.createElement("Root");
        root.addDefaultXmlNameSpace("http://example.com/ns");
        root.addNameSpaceAttribute("xmlns", "xmlns", "http://example.com/other");
        root.addAttribute("id", "42");

        String xml = serializeWriteTo(root);

        Document doc = parseXml(xml);
        assertEquals("Root", doc.getDocumentElement().getLocalName());
        assertEquals("42", doc.getDocumentElement().getAttribute("id"));
        NamedNodeMap attributes = doc.getDocumentElement().getAttributes();
        for (int index = 0; index < attributes.getLength(); index++) {
            Node attribute = attributes.item(index);
            assertFalse("xmlns".equals(attribute.getPrefix()) && "xmlns".equals(attribute.getLocalName()));
        }
    }

    @DisplayName("WriteTo should propagate the default namespace to children without re-declaration or empty xmlns")
    @Test
    void writeToChildDefaultNamespacePropagationTest() throws Exception {
        XmlElement root = XmlElement.createElement("Root");
        root.addDefaultXmlNameSpace("http://example.com/ns");
        XmlElement child = root.addChildElement("Child");
        child.setInnerValue("value");

        String xml = serializeWriteTo(root);

        Document doc = parseXml(xml);
        Element childElement = (Element) doc.getDocumentElement().getChildNodes().item(0);
        assertEquals("http://example.com/ns", childElement.getNamespaceURI());
        assertEquals("value", childElement.getTextContent());
        assertEquals(0, childElement.getAttributes().getLength());
    }

    @DisplayName("WriteTo should escape XML special characters in inner values")
    @Test
    void writeToInnerValueEscapingTest() throws Exception {
        XmlElement root = XmlElement.createElement("Root");
        XmlElement child = root.addChildElement("Child");
        child.setInnerValue("a < b & c > d");

        String xml = serializeWriteTo(root);

        Document doc = parseXml(xml);
        assertEquals("a < b & c > d", doc.getDocumentElement().getChildNodes().item(0).getTextContent());
        assertTrue(xml.contains("&lt;"));
        assertTrue(xml.contains("&amp;"));
    }

    @DisplayName("WriteTo on an empty element should produce a valid empty element")
    @Test
    void writeToEmptyElementTest() throws Exception {
        XmlElement root = XmlElement.createElement("Root");

        String xml = serializeWriteTo(root);

        Document doc = parseXml(xml);
        assertEquals("Root", doc.getDocumentElement().getLocalName());
        assertEquals(0, doc.getDocumentElement().getChildNodes().getLength());
        assertEquals(0, doc.getDocumentElement().getAttributes().getLength());
    }

    @DisplayName("WriteTo on an element with prefix and inner value should preserve both")
    @Test
    void writeToPrefixedElementWithValueTest() throws Exception {
        XmlElement root = XmlElement.createElement("coreProperties", "cp");
        root.addNameSpaceAttribute(
                "cp", "xmlns", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
        root.addNameSpaceAttribute("dc", "xmlns", "http://purl.org/dc/elements/1.1/");
        XmlElement creator = root.addChildElement("creator", "dc");
        creator.setInnerValue("Tester");

        String xml = serializeWriteTo(root);

        Document doc = parseXml(xml);
        assertEquals("cp", doc.getDocumentElement().getPrefix());
        assertEquals("coreProperties", doc.getDocumentElement().getLocalName());
        Element creatorElement = (Element) doc.getDocumentElement().getChildNodes().item(0);
        assertEquals("dc", creatorElement.getPrefix());
        assertEquals("creator", creatorElement.getLocalName());
        assertEquals("http://purl.org/dc/elements/1.1/", creatorElement.getNamespaceURI());
        assertEquals("Tester", creatorElement.getTextContent());
    }

    @DisplayName("FindElementByName should return an IEnumerable with one element, if there is only one matching child")
    @Test
    void findElementByNameTest() {
        XmlElement root = XmlElement.createElement("root");
        root.addChildElementWithValue("node", "test1");
        List<XmlElement> givenResult = toList(root.findChildElementsByName("node"));

        assertEquals(1, givenResult.size());
        assertEquals("test1", givenResult.get(0).getInnerValue());
    }

    @DisplayName("FindElementByName should return an IEnumerable with multiple element, if there more than one matching child")
    @Test
    void findElementByNameTest2() {
        XmlElement root = XmlElement.createElement("root");
        root.addChildElementWithValue("node", "test1");
        root.addChildElementWithValue("node", "test2");
        List<XmlElement> givenResult = toList(root.findChildElementsByName("node"));

        assertEquals(2, givenResult.size());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("test1")).count());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("test2")).count());
    }

    @DisplayName("FindElementByName should return an IEnumerable with multiple element, if there more than one matching child in a complex structure")
    @Test
    void findElementByNameTest3() {
        XmlElement root = XmlElement.createElement("root");
        XmlElement child1 = root.addChildElement("subnode");
        child1.addChildElementWithValue("node", "test1");
        child1.addChildElementWithValue("node", "test2");
        XmlElement child2 = root.addChildElement("subnode2");
        XmlElement child3 = child2.addChildElement("subnode3");
        child3.addChildElementWithValue("node", "test3", "pfx");
        List<XmlElement> givenResult = toList(root.findChildElementsByName("node"));

        assertEquals(3, givenResult.size());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("test1")).count());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("test2")).count());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("test3")).count());
    }

    @DisplayName("FindElementByName should return an IEnumerable with multiple element in the expected order")
    @Test
    void findElementByNameOrderTest() {
        XmlElement root = XmlElement.createElement("root");
        root.addChildElementWithValue("node", "v3");
        root.addChildElementWithValue("unrelated", "v3");
        root.addChildElementWithValue("node", "v2");
        root.addChildElementWithValue("unrelated", "v2");
        root.addChildElementWithValue("node", "v1");
        List<XmlElement> givenResult = toList(root.findChildElementsByName("node"));

        assertEquals(3, givenResult.size());
        assertEquals("node", givenResult.get(0).getName());
        assertEquals("v3", givenResult.get(0).getInnerValue());

        assertEquals("node", givenResult.get(1).getName());
        assertEquals("v2", givenResult.get(1).getInnerValue());

        assertEquals("node", givenResult.get(2).getName());
        assertEquals("v1", givenResult.get(2).getInnerValue());
    }

    @DisplayName("FindElementByName should return an empty IEnumerable, if there is no matching child")
    @ParameterizedTest
    @NullSource
    @ValueSource(strings = {"", " ", "NODE", "node1", "test1"})
    void findElementByNameEmptyTest(String givenName) {
        XmlElement root = XmlElement.createElement("root");
        root.addChildElementWithValue("node", "test1");
        List<XmlElement> givenResult = toList(root.findChildElementsByName(givenName));

        assertTrue(givenResult.isEmpty());
    }

    @DisplayName("FindElementByName should return an empty IEnumerable, if there are no child elements at all")
    @Test
    void findElementByNameEmptyTest2() {
        XmlElement root = XmlElement.createElement("root");
        List<XmlElement> givenResult = toList(root.findChildElementsByName("node"));

        assertTrue(givenResult.isEmpty());
    }

    @DisplayName("FindElementByNameAndAttribute should return an IEnumerable with one element, if there is one matching child")
    @Test
    void findElementByNameAndAttributeTest() {
        XmlElement root = XmlElement.createElement("root");
        root.addChildElementWithAttribute("node", "att1", "test1");
        List<XmlElement> givenResult = toList(root.findChildElementsByNameAndAttribute("node", "att1"));

        assertEquals(1, givenResult.size());
        assertEquals("test1", givenResult.get(0).getAttributes().iterator().next().getValue());
    }

    @DisplayName("FindElementByNameAndAttribute should return an IEnumerable with one element, if there is one matching child with attribute value")
    @Test
    void findElementByNameAndAttributeValueTest() {
        XmlElement root = XmlElement.createElement("root");
        root.addChildElementWithAttribute("node", "att1", "test1");
        List<XmlElement> givenResult = toList(
                root.findChildElementsByNameAndAttribute("node", "att1", "test1"));

        assertEquals(1, givenResult.size());
        assertEquals("test1", givenResult.get(0).getAttributes().iterator().next().getValue());
    }

    @DisplayName("FindElementByNameAndAttribute should return an IEnumerable with multiple elements, if there is more than one matching child")
    @Test
    void findElementByNameAndAttributeTest2() {
        XmlElement root = XmlElement.createElement("root");
        XmlElement child1 = root.addChildElementWithAttribute("node", "att1", "test1");
        child1.setInnerValue("inner-value1");
        XmlElement child2 = root.addChildElementWithAttribute("node", "att1", "other-match");
        child2.setInnerValue("inner-value2");
        XmlElement child3 = root.addChildElementWithAttribute("node", "att1", "test1");
        child3.setInnerValue("inner-value3");
        List<XmlElement> givenResult = toList(root.findChildElementsByNameAndAttribute("node", "att1"));

        assertEquals(3, givenResult.size());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("inner-value1")).count());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("inner-value2")).count());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("inner-value3")).count());
    }

    @DisplayName("FindElementByNameAndAttribute should return an IEnumerable with multiple elements, if there is more than one matching child with attribute value")
    @Test
    void findElementByNameAndAttributeValueTest2() {
        XmlElement root = XmlElement.createElement("root");
        XmlElement child1 = root.addChildElementWithAttribute("node", "att1", "test1");
        child1.setInnerValue("inner-value1");
        XmlElement child2 = root.addChildElementWithAttribute("node", "att1", "no-match");
        child2.setInnerValue("inner-value2");
        XmlElement child3 = root.addChildElementWithAttribute("node", "att1", "test1");
        child3.setInnerValue("inner-value3");
        List<XmlElement> givenResult = toList(
                root.findChildElementsByNameAndAttribute("node", "att1", "test1"));

        assertEquals(2, givenResult.size());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("inner-value1")).count());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("inner-value3")).count());
    }

    @DisplayName("FindElementByName should return an IEnumerable with multiple element, if there more than one matching child in a complex structure")
    @Test
    void findElementByNameAndAttributeTest3() {
        XmlElement root = XmlElement.createElement("root");
        XmlElement child1 = root.addChildElement("subnode");
        XmlElement child1a = child1.addChildElementWithValue("node", "test1");
        child1a.addAttribute("att1", "test1");
        XmlElement child1b = child1.addChildElementWithValue("node", "test2", "pfx");
        child1b.addAttribute("att1", "test1");
        XmlElement child2 = root.addChildElement("subnode2");
        child2.addAttribute("node", "test1");
        XmlElement child3 = child2.addChildElement("subnode3");
        XmlElement child3a = child3.addChildElementWithValue("node", "test3", "pfx");
        child3a.addAttribute("att1", "test1");
        XmlElement child4 = child2.addChildElement("subnode4");
        XmlElement child4a = child3.addChildElementWithValue("node", "test4", "pfx");
        child4a.addAttribute("att1", "other-match");
        List<XmlElement> givenResult = toList(root.findChildElementsByNameAndAttribute("node", "att1"));

        assertEquals(4, givenResult.size());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("test1")).count());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("test2")).count());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("test3")).count());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("test4")).count());
    }

    @DisplayName("FindElementByName should return an IEnumerable with multiple element, if there more than one matching child in a complex structure, with attribute value")
    @Test
    void findElementByNameAndAttributeValueTest3() {
        XmlElement root = XmlElement.createElement("root");
        XmlElement child1 = root.addChildElement("subnode");
        XmlElement child1a = child1.addChildElementWithValue("node", "test1");
        child1a.addAttribute("att1", "test1");
        XmlElement child1b = child1.addChildElementWithValue("node", "test2", "pfx");
        child1b.addAttribute("att1", "test1");
        XmlElement child2 = root.addChildElement("subnode2");
        child2.addAttribute("node", "test1");
        XmlElement child3 = child2.addChildElement("subnode3");
        XmlElement child3a = child3.addChildElementWithValue("node", "test3", "pfx");
        child3a.addAttribute("att1", "test1");
        XmlElement child4 = child2.addChildElement("subnode4");
        XmlElement child4a = child3.addChildElementWithValue("node", "test4", "pfx");
        child4a.addAttribute("att1", "no-match");
        List<XmlElement> givenResult = toList(
                root.findChildElementsByNameAndAttribute("node", "att1", "test1"));

        assertEquals(3, givenResult.size());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("test1")).count());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("test2")).count());
        assertEquals(1, givenResult.stream().filter(node -> node.getInnerValue().equals("test3")).count());
    }

    @DisplayName("FindElementByNameAndAttribute should return an empty IEnumerable, if there is no matching child")
    @ParameterizedTest
    @CsvSource(value = {
            "NULL, att1",
            "'', att1",
            "' ', att1",
            "NULL, att",
            "'', att",
            "' ', att",
            "NODE, att1",
            "node1, att1",
            "test1, att1",
            "NODE, att",
            "node1, att",
            "test1, att",
            "node, att2",
            "node, ATT1",
            "node, att1",
            "node, ATT"
    }, nullValues = "NULL")
    void findElementByNameAndAttributeEmptyTest(String givenTagName, String givenAttributeName) {
        XmlElement root = XmlElement.createElement("root");
        XmlElement child1 = root.addChildElementWithValue("node", "test1");
        child1.addAttribute("att", "test1");
        List<XmlElement> givenResult = toList(
                root.findChildElementsByNameAndAttribute(givenTagName, givenAttributeName));

        assertTrue(givenResult.isEmpty());
    }

    @DisplayName("FindElementByNameAndAttribute should return an empty IEnumerable, if there is no matching child, with attribute value")
    @ParameterizedTest
    @CsvSource(value = {
            "NULL, att1, test1",
            "'', att1, test1",
            "' ', att1, test1",
            "NODE, att1, test1",
            "node1, att1, test1",
            "test1, att1, test1",
            "node, att2, test1",
            "node, ATT1, test1",
            "node, att1, NULL",
            "node, att1, ''",
            "node, att1, ' '",
            "node, att1, TEST1",
            "node, att1, test2"
    }, nullValues = "NULL")
    void findElementByNameAndAttributeEmptyValueTest(
            String givenTagName, String givenAttributeName, String givenAttributeValue) {
        XmlElement root = XmlElement.createElement("root");
        XmlElement child1 = root.addChildElementWithValue("node", "test1");
        child1.addAttribute("att", "test1");
        List<XmlElement> givenResult = toList(root.findChildElementsByNameAndAttribute(
                givenTagName, givenAttributeName, givenAttributeValue));

        assertTrue(givenResult.isEmpty());
    }

    @DisplayName("FindElementByNameAndAttribute should return an empty IEnumerable, if there are no child elements at all")
    @Test
    void findElementByNameAndAttributeEmptyTest2() {
        XmlElement root = XmlElement.createElement("root");
        List<XmlElement> givenResult = toList(
                root.findChildElementsByNameAndAttribute("node", "att1"));

        assertTrue(givenResult.isEmpty());
    }

    @DisplayName("FindElementByNameAndAttribute should return an empty IEnumerable, if there are no child elements at all, with attribute value")
    @Test
    void findElementByNameAndAttributeEmptyValueTest2() {
        XmlElement root = XmlElement.createElement("root");
        List<XmlElement> givenResult = toList(
                root.findChildElementsByNameAndAttribute("node", "att1", "test1"));

        assertTrue(givenResult.isEmpty());
    }

    @DisplayName("FindElementByNameAndAttribute should return an IEnumerable with multiple element in the expected order")
    @Test
    void findElementByNameAndAttributeOrderTest() {
        XmlElement root = XmlElement.createElement("root");
        root.addChildElementWithAttribute("node", "att1", "v3");
        root.addChildElementWithValue("unrelated", "v3");
        root.addChildElementWithAttribute("unrelated", "att1", "v2");
        XmlElement element = root.addChildElementWithAttribute("node", "att1", "v2");
        element.setInnerValue("content");
        root.addChildElementWithAttribute("node", "att2", "unrelated");
        root.addChildElementWithAttribute("node", "att1", "v1", "pfx");
        List<XmlElement> givenResult = toList(root.findChildElementsByNameAndAttribute("node", "att1"));

        assertEquals(3, givenResult.size());
        assertEquals("node", givenResult.get(0).getName());
        assertEquals("v3", givenResult.get(0).getAttributes().iterator().next().getValue());

        assertEquals("node", givenResult.get(1).getName());
        assertEquals("v2", givenResult.get(1).getAttributes().iterator().next().getValue());
        assertEquals("content", givenResult.get(1).getInnerValue());

        assertEquals("node", givenResult.get(2).getName());
        assertEquals("v1", givenResult.get(2).getAttributes().iterator().next().getValue());
    }

    @DisplayName("FindElementByNameAndAttribute with value should return an IEnumerable with multiple element in the expected order")
    @Test
    void findElementByNameAndAttributeValueOrderTest() {
        XmlElement root = XmlElement.createElement("root");
        XmlElement element1 = root.addChildElementWithAttribute("node", "att1", "val");
        element1.setInnerValue("content3");
        root.addChildElementWithValue("unrelated", "v3");
        root.addChildElementWithAttribute("unrelated", "att1", "v2");
        XmlElement unrelated1 = root.addChildElementWithAttribute("node", "att1", "val2");
        unrelated1.setInnerValue("content4");
        XmlElement element2 = root.addChildElementWithAttribute("node", "att1", "val");
        element2.setInnerValue("content2");
        root.addChildElementWithAttribute("node", "att2", "unrelated");
        XmlElement element3 = root.addChildElementWithAttribute("node", "att1", "val", "pfx");
        element3.setInnerValue("content1");
        List<XmlElement> givenResult = toList(
                root.findChildElementsByNameAndAttribute("node", "att1", "val"));

        assertEquals(3, givenResult.size());
        assertEquals("node", givenResult.get(0).getName());
        assertEquals("val", givenResult.get(0).getAttributes().iterator().next().getValue());
        assertEquals("content3", givenResult.get(0).getInnerValue());

        assertEquals("node", givenResult.get(1).getName());
        assertEquals("val", givenResult.get(1).getAttributes().iterator().next().getValue());
        assertEquals("content2", givenResult.get(1).getInnerValue());

        assertEquals("node", givenResult.get(2).getName());
        assertEquals("val", givenResult.get(2).getAttributes().iterator().next().getValue());
        assertEquals("content1", givenResult.get(2).getInnerValue());
    }

    private static String serializeWriteTo(XmlElement root) throws Exception {
        StringWriter stringWriter = new StringWriter();
        XMLStreamWriter writer = XMLOutputFactory.newDefaultFactory().createXMLStreamWriter(stringWriter);
        try {
            root.writeTo(writer);
            writer.flush();
        } finally {
            writer.close();
        }
        return stringWriter.toString();
    }

    private static Document parseXml(String xml) throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newDefaultInstance();
        factory.setNamespaceAware(true);
        factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, "");
        factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_SCHEMA, "");
        return factory.newDocumentBuilder().parse(new InputSource(new StringReader(xml)));
    }

    private static Element findChildElementByLocalName(Element parent, String localName) {
        for (int index = 0; index < parent.getChildNodes().getLength(); index++) {
            Node child = parent.getChildNodes().item(index);
            if (child instanceof Element element && localName.equals(element.getLocalName())) {
                return element;
            }
        }
        return null;
    }

    private static List<XmlElement> toList(Iterable<XmlElement> elements) {
        List<XmlElement> result = new ArrayList<>();
        elements.forEach(result::add);
        return result;
    }
}
