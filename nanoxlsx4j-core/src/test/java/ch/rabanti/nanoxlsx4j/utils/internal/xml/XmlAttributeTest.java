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

import java.util.HashSet;
import java.util.Optional;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

class XmlAttributeTest {

    @DisplayName("CreateXmlAttributeTest: Should initialize properties correctly")
    @ParameterizedTest
    @CsvSource({
            "id, 123, ns",
            "name, value, ''"
    })
    void createXmlAttributeTest(String name, String value, String prefix) {
        XmlAttribute attribute = XmlAttribute.createAttribute(name, value, prefix);

        assertEquals(name, attribute.getName());
        assertEquals(value, attribute.getValue());
        assertEquals(prefix, attribute.getPrefix());
        assertEquals(!prefix.isEmpty(), attribute.isHasPrefix());
    }

    @DisplayName("CreateEmptyAttributeTest: Should create attribute with empty value")
    @ParameterizedTest
    @CsvSource({
            "empty, ''",
            "test, ''"
    })
    void createEmptyAttributeTest(String name, String expectedValue) {
        XmlAttribute attribute = XmlAttribute.createEmptyAttribute(name);

        assertEquals(name, attribute.getName());
        assertEquals(expectedValue, attribute.getValue());
        assertEquals("", attribute.getPrefix());
        assertFalse(attribute.isHasPrefix());
    }

    @DisplayName("EqualsTest: Two attributes with same properties should be equal")
    @ParameterizedTest
    @CsvSource({
            "id, 123, ns",
            "name, value, ''"
    })
    void equalsTest(String name, String value, String prefix) {
        XmlAttribute attribute1 = XmlAttribute.createAttribute(name, value, prefix);
        XmlAttribute attribute2 = XmlAttribute.createAttribute(name, value, prefix);

        assertTrue(attribute1.equals(attribute2));
    }

    @DisplayName("NotEqualsTest: Attributes with different properties should not be equal")
    @Test
    void notEqualsTest() {
        XmlAttribute attribute1 = XmlAttribute.createAttribute("id", "123", "ns");
        XmlAttribute attribute2 = XmlAttribute.createAttribute("id", "456", "ns");

        assertFalse(attribute1.equals(attribute2));
    }

    @DisplayName("GetHashCodeTest: Equal attributes should have the same hash code")
    @ParameterizedTest
    @CsvSource({
            "id, 123, ns",
            "name, value, ''"
    })
    void getHashCodeTest(String name, String value, String prefix) {
        XmlAttribute attribute1 = XmlAttribute.createAttribute(name, value, prefix);
        XmlAttribute attribute2 = XmlAttribute.createAttribute(name, value, prefix);
        int hash1 = attribute1.hashCode();
        int hash2 = attribute2.hashCode();

        assertEquals(hash1, hash2);
    }

    @DisplayName("FindAttribute - Matching when exactly one matching attribute is passed")
    @Test
    void findAttributeTestOneMatching() {
        XmlAttribute attribute = XmlAttribute.createAttribute("test", "value");
        HashSet<XmlAttribute> attributes = new HashSet<>();
        attributes.add(attribute);
        Optional<XmlAttribute> result = XmlAttribute.findAttribute("test", attributes);

        assertTrue(result.isPresent());
        assertEquals(attribute, result.get());
    }

    @DisplayName("FindAttribute - Matching when a HashSet with multiple attributes contains a matching one")
    @Test
    void findAttributeTestMatchingInSet() {
        XmlAttribute matchingAttribute = XmlAttribute.createAttribute("match", "value");
        HashSet<XmlAttribute> attributes = new HashSet<>();
        attributes.add(XmlAttribute.createAttribute("other", "value"));
        attributes.add(matchingAttribute);
        attributes.add(XmlAttribute.createAttribute("another", "value"));
        Optional<XmlAttribute> result = XmlAttribute.findAttribute("match", attributes);

        assertTrue(result.isPresent());
        assertEquals(matchingAttribute, result.get());
    }

    @DisplayName("FindAttribute - Non-matching when null is passed as HashSet")
    @Test
    void findAttributeTestNullSet() {
        Optional<XmlAttribute> result = XmlAttribute.findAttribute("test", null);

        assertTrue(result.isEmpty());
    }

    @DisplayName("FindAttribute - Non-matching when an empty HashSet is passed")
    @Test
    void findAttributeTestEmptySet() {
        HashSet<XmlAttribute> attributes = new HashSet<>();
        Optional<XmlAttribute> result = XmlAttribute.findAttribute("test", attributes);

        assertTrue(result.isEmpty());
    }

    @DisplayName("FindAttribute - Non-matching when no attribute in the HashSet matches the name")
    @Test
    void findAttributeTestNoMatch() {
        HashSet<XmlAttribute> attributes = new HashSet<>();
        attributes.add(XmlAttribute.createAttribute("other", "value"));
        attributes.add(XmlAttribute.createAttribute("another", "value"));
        Optional<XmlAttribute> result = XmlAttribute.findAttribute("test", attributes);

        assertTrue(result.isEmpty());
    }
}
