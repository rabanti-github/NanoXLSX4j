package ch.rabanti.nanoxlsx4j.misc;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import ch.rabanti.nanoxlsx4j.Metadata;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

public class MetadataTest {

	@DisplayName("Test of the get and set function of the application field ")
	@Test()
	void applicationTest() {
		Metadata metadata = new Metadata();
		assertNotNull(metadata.getApplication());
		assertNotEquals("", metadata.getApplication());
		metadata.setApplication("test");
		assertEquals("test", metadata.getApplication());
	}

	@DisplayName("Test of the get and set function of the applicationVersion field ")
	@ParameterizedTest(name = "Given value {1} should lead to a valid application version")
	@CsvSource({
			"NULL, ''",
			"STRING, ''",
			"STRING, '0.1'",
			"STRING, '99999.99999'", })
	void applicationVersionTest(String sourceType, String sourceValue) {
		String version = (String) TestUtils.createInstance(sourceType, sourceValue);
		Metadata metadata = new Metadata();
		assertNotNull(metadata.getApplicationVersion());
		assertNotEquals("", metadata.getApplicationVersion());
		metadata.setApplicationVersion(version);
		assertEquals(version, metadata.getApplicationVersion());
	}

	@DisplayName("Test of failing set function of the applicationVersion field on invalid versions")
	@ParameterizedTest(name = "Given value {0} should lead to an exception")
	@CsvSource({
			"'1'",
			"'1.2.3'",
			"' '",
			"'xyz'",
			"'111111.1'",
			"'1.222222'",
			"'333333.333333'", })
	void applicationVersionFailTest(String version) {
		Metadata metadata = new Metadata();
		assertThrows(FormatException.class, () -> metadata.setApplicationVersion(version));
	}

	@DisplayName("Test of the get and set function of the category field ")
	@Test()
	void categoryTest() {
		Metadata metadata = new Metadata();
		assertNull(metadata.getCategory());
		metadata.setCategory("test");
		assertEquals("test", metadata.getCategory());
	}

	@DisplayName("Test of the get and set function of the company field ")
	@Test()
	void companyTest() {
		Metadata metadata = new Metadata();
		assertNull(metadata.getCompany());
		metadata.setCompany("test");
		assertEquals("test", metadata.getCompany());
	}

	@DisplayName("Test of the get and set function of the contentStatus field ")
	@Test()
	void contentStatusTest() {
		Metadata metadata = new Metadata();
		assertNull(metadata.getContentStatus());
		metadata.setContentStatus("test");
		assertEquals("test", metadata.getContentStatus());
	}

	@DisplayName("Test of the get and set function of the creator field ")
	@Test()
	void creatorTest() {
		Metadata metadata = new Metadata();
		assertNull(metadata.getCreator());
		metadata.setCreator("test");
		assertEquals("test", metadata.getCreator());
	}

	@DisplayName("Test of the get and set function of the description field ")
	@Test()
	void descriptionTest() {
		Metadata metadata = new Metadata();
		assertNull(metadata.getDescription());
		metadata.setDescription("test");
		assertEquals("test", metadata.getDescription());
	}

	@DisplayName("Test of the get and set function of the hyperlinkBase field ")
	@Test()
	void hyperlinkBaseTest() {
		Metadata metadata = new Metadata();
		assertNull(metadata.getHyperlinkBase());
		metadata.setHyperlinkBase("test");
		assertEquals("test", metadata.getHyperlinkBase());
	}

	@DisplayName("Test of the get and set function of the keywords field ")
	@Test()
	void keywordsTest() {
		Metadata metadata = new Metadata();
		assertNull(metadata.getKeywords());
		metadata.setKeywords("test");
		assertEquals("test", metadata.getKeywords());
	}

	@DisplayName("Test of the get and set function of the manager field ")
	@Test()
	void managerTest() {
		Metadata metadata = new Metadata();
		assertNull(metadata.getManager());
		metadata.setManager("test");
		assertEquals("test", metadata.getManager());
	}

	@DisplayName("Test of the get and set function of the subject field ")
	@Test()
	void subjectTest() {
		Metadata metadata = new Metadata();
		assertNull(metadata.getSubject());
		metadata.setSubject("test");
		assertEquals("test", metadata.getSubject());
	}

	@DisplayName("Test of the get and set function of the title field ")
	@Test()
	void titleTest() {
		Metadata metadata = new Metadata();
		assertNull(metadata.getTitle());
		metadata.setTitle("test");
		assertEquals("test", metadata.getTitle());
	}

	@DisplayName("Test of the Constructor")
	@Test()
	void constructorTest() {
		Metadata metadata = new Metadata();
		assertNotNull(metadata);
		assertNotEquals("", metadata.getApplication());
		assertNotEquals("", metadata.getApplicationVersion());
	}

	@DisplayName("Test of the ParseVersion function")
	@ParameterizedTest(name = "Given major {0}, minor {1}, build {2} and revision {3} number should lead to the version string {4}")
	@CsvSource({
			"1, 2, 2, 5, 1.225",
			"4, 2, 2, 0, 4.22",
			"11, 2, 0, 0, 11.2",
			"112, 0, 0, 0, 112.0",
			"0, 0, 0, 0, 0.0",
			"0, 4, 5, 1, 0.451",
			"0, 0, 2, 1, 0.021",
			"0, 0, 0, 1, 0.001",
			"9999, 666, 555, 444, 9999.66655",
			"99999, 0, 0, 1234567, 99999.00123", })
	void parseVersionTest(int major, int minor, int build, int revision, String expectedVersion) {
		String version = Metadata.parseVersion(major, minor, build, revision);
		assertEquals(expectedVersion, version);
	}

	@DisplayName("Test of the failingParseVersion function")
	@ParameterizedTest(name = "Given major {0}, minor {1}, build {2} or revision {3} number should lead to an exception")
	@CsvSource({
			"111111, 1, 1, 1",
			"-1, 1, 1, 1",
			"1, -1, 1, 1",
			"1, 1, -1, 1",
			"1, 1, 1, -1", })
	void parseVersionFailTest(int major, int minor, int build, int revision) {
		assertThrows(FormatException.class, () -> Metadata.parseVersion(major, minor, build, revision));
	}

}
