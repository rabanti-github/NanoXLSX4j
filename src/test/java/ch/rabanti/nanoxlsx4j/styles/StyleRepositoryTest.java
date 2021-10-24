package ch.rabanti.nanoxlsx4j.styles;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

public class StyleRepositoryTest {

    @BeforeEach
    private void flush() {
        StyleRepository.getInstance().getStyles().clear();
    }

    @DisplayName("Test of the addStyle method")
    @Test()
    void addStyleTest() {
        StyleRepository repository = StyleRepository.getInstance();
        assertEquals(0, repository.getStyles().size());
        Style style = new Style();
        style.getFont().setName("Arial");
        Style result = repository.addStyle(style);
        assertEquals(1, repository.getStyles().size());
        assertEquals(style.hashCode(), result.hashCode());
        assertEquals(style.hashCode(), repository.getStyles().get(style.hashCode()).hashCode());
    }


    @DisplayName("Test of the AddStyle method on a null object")
    @Test()
    void addStyleTest2() {
        StyleRepository repository = StyleRepository.getInstance();
        assertEquals(0, repository.getStyles().size());
        Style result = repository.addStyle(null);
        assertEquals(0, repository.getStyles().size());
        assertNull(result);
    }


    @DisplayName("Test of the flush method")
    @Test()
    void flushTest() {
        StyleRepository repository = StyleRepository.getInstance();
        assertEquals(0, repository.getStyles().size());
        Style style = new Style();
        style.getFont().setName("Arial");
        Style result = repository.addStyle(style);
        assertEquals(1, repository.getStyles().size());
        repository.flushStyles();
        assertEquals(0, repository.getStyles().size());
    }

}
