/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.misc;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;
import ch.rabanti.nanoxlsx4j.Metadata;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

public class MetadataTest {

    @Test
    @DisplayName("Test of the getApplication and setApplication methods")
    public void applicationTest() {
        Metadata metadata = new Metadata();
        assertNotNull(metadata.getApplication());
        assertFalse(metadata.getApplication().isEmpty());
        metadata.setApplication("test");
        assertEquals("test", metadata.getApplication());
    }

    @ParameterizedTest
    @NullSource
    @ValueSource(
            strings = {
                    "",
                    "0.1",
                    "99999.99999"}
    )
    @DisplayName("Test of the getApplicationVersion and setApplicationVersion methods")
    public void applicationVersionTest(String version) {
        Metadata metadata = new Metadata();
        assertNotNull(metadata.getApplicationVersion());
        assertFalse(metadata.getApplicationVersion().isEmpty());
        metadata.setApplicationVersion(version);
        assertEquals(version, metadata.getApplicationVersion());
    }

    @ParameterizedTest
    @ValueSource(
            strings = {
                    "1",
                    "1.2.3",
                    " ",
                    "xyz",
                    "111111.1",
                    "1.222222",
                    "333333.333333"}
    )
    @DisplayName("Test of failing setApplicationVersion method on invalid versions")
    public void applicationVersionFailTest(String version) {
        Metadata metadata = new Metadata();
        assertThrows(FormatException.class, () -> metadata.setApplicationVersion(version));
    }

    @Test
    @DisplayName("Test of the getCategory and setCategory methods")
    public void categoryTest() {
        Metadata metadata = new Metadata();
        assertNull(metadata.getCategory());
        metadata.setCategory("test");
        assertEquals("test", metadata.getCategory());
    }

    @Test
    @DisplayName("Test of the getCompany and setCompany methods")
    public void companyTest() {
        Metadata metadata = new Metadata();
        assertNull(metadata.getCompany());
        metadata.setCompany("test");
        assertEquals("test", metadata.getCompany());
    }

    @Test
    @DisplayName("Test of the getContentStatus and setContentStatus methods")
    public void contentStatusTest() {
        Metadata metadata = new Metadata();
        assertNull(metadata.getContentStatus());
        metadata.setContentStatus("test");
        assertEquals("test", metadata.getContentStatus());
    }

    @Test
    @DisplayName("Test of the getCreator and setCreator methods")
    public void creatorTest() {
        Metadata metadata = new Metadata();
        assertNull(metadata.getCreator());
        metadata.setCreator("test");
        assertEquals("test", metadata.getCreator());
    }

    @Test
    @DisplayName("Test of the getDescription and setDescription methods")
    public void descriptionTest() {
        Metadata metadata = new Metadata();
        assertNull(metadata.getDescription());
        metadata.setDescription("test");
        assertEquals("test", metadata.getDescription());
    }

    @Test
    @DisplayName("Test of the getHyperlinkBase and setHyperlinkBase methods")
    public void hyperlinkBaseTest() {
        Metadata metadata = new Metadata();
        assertNull(metadata.getHyperlinkBase());
        metadata.setHyperlinkBase("test");
        assertEquals("test", metadata.getHyperlinkBase());
    }

    @Test
    @DisplayName("Test of the getKeywords and setKeywords methods")
    public void keywordsTest() {
        Metadata metadata = new Metadata();
        assertNull(metadata.getKeywords());
        metadata.setKeywords("test");
        assertEquals("test", metadata.getKeywords());
    }

    @Test
    @DisplayName("Test of the getManager and setManager methods")
    public void managerTest() {
        Metadata metadata = new Metadata();
        assertNull(metadata.getManager());
        metadata.setManager("test");
        assertEquals("test", metadata.getManager());
    }

    @Test
    @DisplayName("Test of the getSubject and setSubject methods")
    public void subjectTest() {
        Metadata metadata = new Metadata();
        assertNull(metadata.getSubject());
        metadata.setSubject("test");
        assertEquals("test", metadata.getSubject());
    }

    @Test
    @DisplayName("Test of the getTitle and setTitle methods")
    public void titleTest() {
        Metadata metadata = new Metadata();
        assertNull(metadata.getTitle());
        metadata.setTitle("test");
        assertEquals("test", metadata.getTitle());
    }

    @Test
    @DisplayName("Test of the Constructor")
    public void constructorTest() {
        Metadata metadata = new Metadata();
        assertNotNull(metadata);
        assertFalse(metadata.getApplication().isEmpty());
        assertFalse(metadata.getApplicationVersion().isEmpty());
        assertEquals(Metadata.DEFAULT_APPLICATION_NAME, metadata.getApplication());
        assertEquals(Metadata.DEFAULT_APPLICATION_VERSION, metadata.getApplicationVersion());
        assertNotEquals("0.0", Metadata.DEFAULT_APPLICATION_VERSION);
    }

    @ParameterizedTest
    @CsvSource(
            {
                    "1, 2, 2, 5, 1.225",
                    "4, 2, 2, 0, 4.22",
                    "11, 2, 0, 0, 11.2",
                    "112, 0, 0, 0, 112.0",
                    "0, 0, 0, 0, 0.0",
                    "0, 4, 5, 1, 0.451",
                    "0, 0, 2, 1, 0.021",
                    "0, 0, 0, 1, 0.001",
                    "9999, 666, 555, 444, 9999.66655",
                    "99999, 0, 0, 1234567, 99999.00123"
            }
    )
    @DisplayName("Test of the parseVersion method")
    public void parseVersionTest(int major, int minor, int build, int revision, String expectedVersion) {
        String version = Metadata.parseVersion(major, minor, build, revision);
        assertEquals(expectedVersion, version);
    }

    @ParameterizedTest
    @CsvSource(
            {
                    "111111, 1, 1, 1",
                    "-1, 1, 1, 1",
                    "1, -1, 1, 1",
                    "1, 1, -1, 1",
                    "1, 1, 1, -1"
            }
    )
    @DisplayName("Test of the failing parseVersion method")
    public void parseVersionFailTest(int major, int minor, int build, int revision) {
        assertThrows(FormatException.class, () -> Metadata.parseVersion(major, minor, build, revision));
    }
}
