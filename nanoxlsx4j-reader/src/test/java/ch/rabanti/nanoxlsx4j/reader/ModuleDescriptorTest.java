package ch.rabanti.nanoxlsx4j.reader;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.module.ModuleDescriptor;
import java.util.Set;
import java.util.stream.Collectors;
import org.junit.jupiter.api.Test;

class ModuleDescriptorTest {

    @Test
    void hasExpectedModuleBoundary() {
        Module module = getClass().getModule();
        Module coreModule = ModuleLayer.boot()
            .findModule("ch.rabanti.nanoxlsx4j.core")
            .orElseThrow();

        assertTrue(module.isNamed());
        assertEquals("ch.rabanti.nanoxlsx4j.reader", module.getName());
        assertTrue(module.canRead(coreModule));

        Set<String> requiredModules = requiredModules(module);
        assertTrue(requiredModules.contains("ch.rabanti.nanoxlsx4j.core"));
        assertFalse(requiredModules.contains("ch.rabanti.nanoxlsx4j.writer"));
    }

    private static Set<String> requiredModules(Module module) {
        return module.getDescriptor().requires().stream()
            .map(ModuleDescriptor.Requires::name)
            .collect(Collectors.toUnmodifiableSet());
    }
}
