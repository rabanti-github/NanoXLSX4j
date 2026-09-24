package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.math.BigDecimal;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

final class WorksheetTestUtils {

    private WorksheetTestUtils() {
    }

    static Object createInstance(String sourceType, String stringValue) {
        return switch (sourceType.toUpperCase()) {
            case "BIGDECIMAL" -> new BigDecimal(stringValue);
            case "INTEGER" -> Integer.parseInt(stringValue);
            case "LONG" -> Long.parseLong(stringValue);
            case "BYTE" -> Byte.parseByte(stringValue);
            case "DOUBLE" -> Double.parseDouble(stringValue);
            case "FLOAT" -> Float.parseFloat(stringValue);
            case "BOOLEAN" -> Boolean.parseBoolean(stringValue);
            case "STRING" -> stringValue;
            case "NULL" -> null;
            default -> throw new IllegalArgumentException("Not implemented source type: " + sourceType);
        };
    }

    static List<String> splitValuesAsList(String valueString) {
        if (valueString == null || valueString.isEmpty()) {
            return new ArrayList<>();
        }
        List<String> output = new ArrayList<>();
        for (String value : valueString.split(",|\\s")) {
            if (!value.isEmpty()) {
                output.add(value);
            }
        }
        return output;
    }

    static <K, V> void assertMapEntry(K expectedKey, V expectedValue, Map<K, V> map) {
        assertNotNull(map);
        assertFalse(map.isEmpty());
        assertTrue(map.containsKey(expectedKey));
        assertEquals(expectedValue, map.get(expectedKey));
    }

    static <K, MV, V> void assertMapEntry(K expectedKey, V expectedValue,
            Map<K, MV> map, Function<MV, V> method) {
        assertNotNull(map);
        assertFalse(map.isEmpty());
        assertTrue(map.containsKey(expectedKey));
        V actualValue = method.apply(map.get(expectedKey));
        if (actualValue == null) {
            assertNull(expectedValue);
        } else {
            assertEquals(expectedValue, actualValue);
        }
    }

    static <T> void assertListEntry(T expectedEntry, List<T> list) {
        assertNotNull(list);
        assertFalse(list.isEmpty());
        assertTrue(list.contains(expectedEntry));
    }

    static Date buildDate(int year, int month, int day, int hour, int minute, int second) {
        Calendar calendar = new GregorianCalendar();
        calendar.set(year, month, day, hour, minute, second);
        calendar.set(Calendar.MILLISECOND, 0);
        return calendar.getTime();
    }

    static Duration buildTime(int hour, int minute, int second) {
        return Duration.ofHours(hour).plusMinutes(minute).plusSeconds(second);
    }

    @FunctionalInterface
    interface TriConsumer<T1, T2, T3> {
        void accept(T1 t1, T2 t2, T3 t3);
    }

    @FunctionalInterface
    interface QuadConsumer<T1, T2, T3, T4> {
        void accept(T1 t1, T2 t2, T3 t3, T4 t4);
    }
}
