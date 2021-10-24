package ch.rabanti.nanoxlsx4j;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.*;

public class TestUtils {

    public static Object createInstance(String sourceType, String stringValue) {
        switch (sourceType.toUpperCase()) {
            case "INTEGER":
                return Integer.parseInt(stringValue);
            case "LONG":
                return Long.parseLong(stringValue);
            case "BYTE":
                return Byte.parseByte(stringValue);
            case "DOUBLE":
                return Double.parseDouble(stringValue);
            case "FLOAT":
                return Float.parseFloat(stringValue);
            case "BOOLEAN":
                return Boolean.parseBoolean(stringValue);
            case "STRING":
                return stringValue;
            case "NULL":
                return null;
            default:
                throw new IllegalArgumentException("Not implemented source type: " + sourceType);
        }
    }

    public static void assertCellRange(String expectedAddresses, List<Address> addresses) {
        String[] addressStrings = splitValues(expectedAddresses);
        List<Address> expected = new ArrayList<Address>();
        for (String address : addressStrings) {
            expected.add(new Address(address));
        }
        assertEquals(expected.size(), addresses.size());
        for (int i = 0; i < expected.size(); i++) {
            assertEquals(expected.get(i), addresses.get(i));
        }
    }

    public static String[] splitValues(String valueString) {
        List<String> list = splitValuesAsList(valueString);
        String[] output = new String[list.size()];
        list.toArray(output);
        return output;
    }

    public static List<String> splitValuesAsList(String valueString) {
        if (valueString == null || valueString.isEmpty()) {
            return new ArrayList<>();
        }
        String[] valueStrings = valueString.split(",|\\s");
        List<String> output = new ArrayList<>();
        for (String value : valueStrings) {
            if (value != null && !value.isEmpty()) {
                output.add(value);
            }
        }
        return output;
    }

    public static <T> void assertCellsEqual(T value1, T value2, T inequalValue, Address cellAddress) {
        Cell cell1 = new Cell(value1, Cell.CellType.DEFAULT, cellAddress);
        Cell cell2 = new Cell(value2, Cell.CellType.DEFAULT, cellAddress);
        Cell cell3 = new Cell(inequalValue, Cell.CellType.DEFAULT, cellAddress);
        assertTrue(cell1.equals(cell2));
        assertFalse(cell1.equals(cell3));
    }

    public static <K, V> void assertMapEntry(K expectedKey, V expectedValue, Map<K, V> map) {
        assertNotNull(map);
        assertNotEquals(0, map.size());
        assertTrue(map.containsKey(expectedKey));
        assertTrue(map.get(expectedKey).equals(expectedValue));
    }

    public static <K, MV, V> void assertMapEntry(K expectedKey, V expectedValue, Map<K, MV> map, Function<MV, V> method) {
        assertNotNull(map);
        assertNotEquals(0, map.size());
        assertTrue(map.containsKey(expectedKey));
        MV mapValue = map.get(expectedKey);
        V actualValue = method.apply(mapValue);
        if (actualValue == null) {
            assertNull(expectedValue);
        } else {
            assertTrue(actualValue.equals(expectedValue));
        }
    }

    public static <T> void assertListEntry(T expectedEntry, List<T> list) {
        assertNotNull(list);
        assertNotEquals(0, list.size());
        assertTrue(list.contains(expectedEntry));
    }

    public interface TriConsumer<T1, T2, T3> {
        void accept(T1 t1, T2 t2, T3 t3);

        default TriConsumer<T1, T2, T3> andThen(TriConsumer<? super T1, ? super T2, ? super T3> after) {
            Objects.requireNonNull(after);
            return (l, r, s) -> {
                accept(l, r, s);
                after.accept(l, r, s);
            };
        }
    }

    public static InputStream getResource(String resourceName) {
        ClassLoader classLoader = TestUtils.class.getClassLoader();
        return ClassLoader.getSystemResourceAsStream(resourceName);
    }

    public interface QuadConsumer<T1, T2, T3, T4> {
        void accept(T1 t1, T2 t2, T3 t3, T4 t4);

        default QuadConsumer<T1, T2, T3, T4> andThen(QuadConsumer<? super T1, ? super T2, ? super T3, ? super T4> after) {
            Objects.requireNonNull(after);
            return (l, r, s, t) -> {
                accept(l, r, s, t);
                after.accept(l, r, s, t);
            };
        }
    }
}
