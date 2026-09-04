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

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class NanoXlsxQueuePluginTest {

    @Test
    @DisplayName("PlugInUUID and QueueUUID are required annotation elements")
    void requiredUuidTest() throws NoSuchMethodException {
        assertNull(NanoXlsxQueuePlugin.class.getDeclaredMethod("pluginUuid").getDefaultValue());
        assertNull(NanoXlsxQueuePlugin.class.getDeclaredMethod("queueUuid").getDefaultValue());
    }

    @Test
    @DisplayName("Default PlugInOrder should be 0")
    void defaultPluginOrderTest() {
        NanoXlsxQueuePlugin annotation = QueuePlugin.class.getAnnotation(NanoXlsxQueuePlugin.class);
        assertEquals(0, annotation.order());
    }

    @Test
    @DisplayName("PlugInOrder Get Test")
    void pluginOrderGetTest() {
        NanoXlsxQueuePlugin[] annotations = RepeatedQueuePlugin.class
            .getAnnotationsByType(NanoXlsxQueuePlugin.class);

        assertEquals(2, annotations.length);
        assertEquals("PLUGIN-1", annotations[0].pluginUuid());
        assertEquals("QUEUE-1", annotations[0].queueUuid());
        assertEquals(-1, annotations[0].order());
        assertEquals("PLUGIN-2", annotations[1].pluginUuid());
        assertEquals("QUEUE-2", annotations[1].queueUuid());
        assertEquals(10, annotations[1].order());
    }

    @Test
    void hasRuntimeTypeMetadataTest() {
        Retention retention = NanoXlsxQueuePlugin.class.getAnnotation(Retention.class);
        Target target = NanoXlsxQueuePlugin.class.getAnnotation(Target.class);

        assertEquals(RetentionPolicy.RUNTIME, retention.value());
        assertEquals(1, target.value().length);
        assertEquals(ElementType.TYPE, target.value()[0]);
    }

    @NanoXlsxQueuePlugin(pluginUuid = "PLUGIN", queueUuid = "QUEUE")
    private static class QueuePlugin {
    }

    @NanoXlsxQueuePlugin(pluginUuid = "PLUGIN-1", queueUuid = "QUEUE-1", order = -1)
    @NanoXlsxQueuePlugin(pluginUuid = "PLUGIN-2", queueUuid = "QUEUE-2", order = 10)
    private static class RepeatedQueuePlugin {
    }
}
