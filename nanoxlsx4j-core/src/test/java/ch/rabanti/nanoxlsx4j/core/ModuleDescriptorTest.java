package ch.rabanti.nanoxlsx4j.core;

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

        assertTrue(module.isNamed());
        assertEquals("ch.rabanti.nanoxlsx4j.core", module.getName());

        Set<String> requiredModules = requiredModules(module);
        assertFalse(requiredModules.contains("ch.rabanti.nanoxlsx4j.reader"));
        assertFalse(requiredModules.contains("ch.rabanti.nanoxlsx4j.writer"));
        assertTrue(module.getDescriptor().exports().stream()
            .anyMatch(export -> export.source().equals("ch.rabanti.nanoxlsx4j.colors") && !export.isQualified()));
    }

    private static Set<String> requiredModules(Module module) {
        return module.getDescriptor().requires().stream()
            .map(ModuleDescriptor.Requires::name)
            .collect(Collectors.toUnmodifiableSet());
    }
}
