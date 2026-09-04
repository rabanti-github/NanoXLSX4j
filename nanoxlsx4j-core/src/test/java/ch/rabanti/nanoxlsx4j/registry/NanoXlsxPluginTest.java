/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.registry;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class NanoXlsxPluginTest {

    @Test
    @DisplayName("Default PlugInOrder should be 0")
    void defaultPluginOrderTest() {
        NanoXlsxPlugin annotation = ReplacementPlugin.class.getAnnotation(NanoXlsxPlugin.class);
        assertEquals(0, annotation.order());
    }

    @Test
    @DisplayName("PlugInUUID is a required annotation element")
    void pluginUuidRequiredTest() throws NoSuchMethodException {
        assertNull(NanoXlsxPlugin.class.getDeclaredMethod("pluginUuid").getDefaultValue());
    }

    @Test
    void hasRuntimeTypeMetadataTest() {
        Retention retention = NanoXlsxPlugin.class.getAnnotation(Retention.class);
        Target target = NanoXlsxPlugin.class.getAnnotation(Target.class);

        assertEquals(RetentionPolicy.RUNTIME, retention.value());
        assertEquals(1, target.value().length);
        assertEquals(ElementType.TYPE, target.value()[0]);
    }

    @Test
    @DisplayName("PlugInOrder Get Test")
    void pluginOrderGetTest() {
        NanoXlsxPlugin annotation = OrderedReplacementPlugin.class.getAnnotation(NanoXlsxPlugin.class);
        assertEquals(10, annotation.order());
        assertEquals("UUID-123", annotation.pluginUuid());
    }

    @NanoXlsxPlugin(pluginUuid = "DEFAULT")
    private static class ReplacementPlugin {
    }

    @NanoXlsxPlugin(pluginUuid = "UUID-123", order = 10)
    private static class OrderedReplacementPlugin {
    }
}
