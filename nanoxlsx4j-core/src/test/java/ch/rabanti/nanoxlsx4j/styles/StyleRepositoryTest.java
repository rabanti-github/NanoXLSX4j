/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class StyleRepositoryTest {

    @BeforeEach
    void flush() {
        StyleRepository.getInstance().getStyles().clear();
    }

    @Test
    @DisplayName("Test of the AddStyle method")
    void addStyleTest() {
        StyleRepository repository = StyleRepository.getInstance();
        assertTrue(repository.getStyles().isEmpty());
        Style style = new Style();
        style.getCurrentFont().setName("Arial");
        Style result = repository.addStyle(style);
        assertEquals(1, repository.getStyles().size());
        assertEquals(style.hashCode(), result.hashCode());
        assertEquals(style.hashCode(), repository.getStyles().get(style.hashCode()).hashCode());
    }

    @Test
    @DisplayName("Test of the AddStyle method on a null object")
    void addStyleTest2() {
        StyleRepository repository = StyleRepository.getInstance();
        assertTrue(repository.getStyles().isEmpty());
        Style result = repository.addStyle(null);
        assertTrue(repository.getStyles().isEmpty());
        assertNull(result);
    }

    @Test
    @DisplayName("Test of the Flush method")
    void flushTest() {
        StyleRepository repository = StyleRepository.getInstance();
        assertTrue(repository.getStyles().isEmpty());
        Style style = new Style();
        style.getCurrentFont().setName("Arial");
        repository.addStyle(style);
        assertEquals(1, repository.getStyles().size());
        repository.flushStyles();
        assertTrue(repository.getStyles().isEmpty());
    }
}
