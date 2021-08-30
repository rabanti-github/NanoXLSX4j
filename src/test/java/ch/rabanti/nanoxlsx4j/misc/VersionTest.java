package ch.rabanti.nanoxlsx4j.misc;

import ch.rabanti.nanoxlsx4j.Version;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

public class VersionTest {
    // For code coverage
    @DisplayName("Test that the Version class returns non-null and non-empty values")
    @Test
    public void versionTest(){
        Assertions.assertNotEquals("", Version.VERSION);
        assertNotEquals("", Version.APPLICATIONNAME);
        assertNotNull( Version.VERSION);
        assertNotNull(Version.APPLICATIONNAME);
    }
}
