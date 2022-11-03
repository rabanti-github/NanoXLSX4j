package ch.rabanti.nanoxlsx4j.misc;

import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import ch.rabanti.nanoxlsx4j.Version;

public class VersionTest {
	// For code coverage
	@DisplayName("Test that the Version class returns non-null and non-empty values")
	@Test
	public void versionTest() {
		Assertions.assertNotEquals("", Version.VERSION);
		assertNotEquals("", Version.APPLICATION_NAME);
		assertNotNull(Version.VERSION);
		assertNotNull(Version.APPLICATION_NAME);
	}
}
