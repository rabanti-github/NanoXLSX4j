package ch.rabanti.nanoxlsx4j.misc;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import ch.rabanti.nanoxlsx4j.internal.AuxiliaryData;

class AuxiliaryDataTest {

    @ParameterizedTest
    @CsvSource(
            value = {
                    "plugin1|0|test_value|STRING",
                    "plugin2|42|12345|INTEGER",
                    "plugin3|999|true|BOOLEAN"
            },
            delimiter = '|'
    )
    @DisplayName("Test of setData and getData with integer valueId")
    void setDataGetDataWithIntValueIdTest(
            String pluginId, int valueId, String rawValue, TestValueType valueType
    ) {
        Object value = valueType.convert(rawValue);
        AuxiliaryData data = new AuxiliaryData();
        data.setData(pluginId, valueId, value);

        Object result = data.getData(pluginId, valueId);
        assertEquals(value, result);
    }

    @ParameterizedTest
    @CsvSource(
            value = {
                    "plugin1|key1|test_value|STRING",
                    "plugin2|A1|12345|INTEGER",
                    "plugin3|cell_ref|false|BOOLEAN"
            },
            delimiter = '|'
    )
    @DisplayName("Test of setData and getData with string valueId")
    void setDataGetDataWithStringValueIdTest(
            String pluginId, String valueId, String rawValue, TestValueType valueType
    ) {
        Object value = valueType.convert(rawValue);
        AuxiliaryData data = new AuxiliaryData();
        data.setData(pluginId, valueId, value);

        Object result = data.getData(pluginId, valueId);
        assertEquals(value, result);
    }

    @ParameterizedTest
    @CsvSource(
            value = {
                    "plugin1|entity1|0|test_value|STRING",
                    "plugin2|worksheet1|42|12345|INTEGER",
                    "plugin3|sheet_A|999|true|BOOLEAN"
            },
            delimiter = '|'
    )
    @DisplayName("Test of setData and getData with entityId and integer valueId")
    void setDataGetDataWithEntityIdAndIntValueIdTest(
            String pluginId, String entityId, int valueId, String rawValue, TestValueType valueType
    ) {
        Object value = valueType.convert(rawValue);
        AuxiliaryData data = new AuxiliaryData();
        data.setData(pluginId, entityId, valueId, value);

        Object result = data.getData(pluginId, entityId, valueId);
        assertEquals(value, result);
    }

    @ParameterizedTest
    @CsvSource(
            value = {
                    "plugin1|entity1|key1|test_value|STRING",
                    "plugin2|worksheet1|A1|12345|INTEGER",
                    "plugin3|sheet_A|cell_ref|false|BOOLEAN"
            },
            delimiter = '|'
    )
    @DisplayName("Test of setData and getData with entityId and string valueId")
    void setDataGetDataWithEntityIdAndStringValueIdTest(
            String pluginId, String entityId, String valueId, String rawValue, TestValueType valueType
    ) {
        Object value = valueType.convert(rawValue);
        AuxiliaryData data = new AuxiliaryData();
        data.setData(pluginId, entityId, valueId, value);

        Object result = data.getData(pluginId, entityId, valueId);
        assertEquals(value, result);
    }

    @ParameterizedTest
    @CsvSource(
            value = {
                    "plugin1|0|test_value|STRING",
                    "plugin2|42|12345|INTEGER",
                    "plugin3|999|true|BOOLEAN"
            },
            delimiter = '|'
    )
    @DisplayName("Test of generic getData<T> with integer valueId")
    void getDataGenericWithIntValueIdTest(
            String pluginId, int valueId, String rawValue, TestValueType valueType
    ) {
        Object value = valueType.convert(rawValue);
        AuxiliaryData data = new AuxiliaryData();
        data.setData(pluginId, valueId, value);

        Object result = data.getData(pluginId, valueId, value.getClass());
        assertEquals(value, result);
    }

    @ParameterizedTest
    @CsvSource(
            value = {
                    "plugin1|key1|test_value|STRING",
                    "plugin2|A1|12345|INTEGER",
                    "plugin3|cell_ref|true|BOOLEAN"
            },
            delimiter = '|'
    )
    @DisplayName("Test of generic getData<T> with string valueId")
    void getDataGenericWithStringValueIdTest(
            String pluginId, String valueId, String rawValue, TestValueType valueType
    ) {
        Object value = valueType.convert(rawValue);
        AuxiliaryData data = new AuxiliaryData();
        data.setData(pluginId, valueId, value);

        Object result = data.getData(pluginId, valueId, value.getClass());
        assertEquals(value, result);
    }

    @ParameterizedTest
    @CsvSource(
            value = {
                    "plugin1|entity1|0|test_value|STRING",
                    "plugin2|worksheet1|42|12345|INTEGER",
                    "plugin3|sheet_A|999|true|BOOLEAN"
            },
            delimiter = '|'
    )
    @DisplayName("Test of generic getData<T> with entityId and integer valueId")
    void getDataGenericWithEntityIdAndIntValueIdTest(
            String pluginId, String entityId, int valueId, String rawValue, TestValueType valueType
    ) {
        Object value = valueType.convert(rawValue);
        AuxiliaryData data = new AuxiliaryData();
        data.setData(pluginId, entityId, valueId, value);

        Object result = data.getData(pluginId, entityId, valueId, value.getClass());
        assertEquals(value, result);
    }

    @ParameterizedTest
    @CsvSource(
            value = {
                    "plugin1|entity1|key1|test_value|STRING",
                    "plugin2|worksheet1|A1|12345|INTEGER",
                    "plugin3|sheet_A|cell_ref|false|BOOLEAN"
            },
            delimiter = '|'
    )
    @DisplayName("Test of generic getData<T> with entityId and string valueId")
    void getDataGenericWithEntityIdAndStringValueIdTest(
            String pluginId, String entityId, String valueId, String rawValue, TestValueType valueType
    ) {
        Object value = valueType.convert(rawValue);
        AuxiliaryData data = new AuxiliaryData();
        data.setData(pluginId, entityId, valueId, value);

        Object result = data.getData(pluginId, entityId, valueId, value.getClass());
        assertEquals(value, result);
    }

    @Test
    @DisplayName("Test of getData returning null for non-existent data")
    void getDataNonExistentTest() {
        AuxiliaryData data = new AuxiliaryData();

        assertNull(data.getData("plugin1", 0));
        assertNull(data.getData("plugin1", "key1"));
        assertNull(data.getData("plugin1", "entity1", 0));
        assertNull(data.getData("plugin1", "entity1", "key1"));
    }

    @Test
    @DisplayName("Test of generic getData<T> returning default for non-existent data")
    void getDataGenericNonExistentTest() {
        AuxiliaryData data = new AuxiliaryData();

        // Java generic return types are references, so their default is null for every target type.
        assertNull(data.getData("plugin1", 0, String.class));
        assertNull(data.getData("plugin1", "key1", Integer.class));
        assertNull(data.getData("plugin1", "entity1", 0, Boolean.class));
        assertNull(data.getData("plugin1", "entity1", "key1", String.class));
    }

    @Test
    @DisplayName("Test of generic getData<T> returning default for wrong type")
    void getDataGenericWrongTypeTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", 0, "string_value");

        assertNull(data.getData("plugin1", 0, Integer.class));
    }

    @Test
    @DisplayName("Test of setData update behavior")
    void setDataUpdateTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", 0, "initial_value");

        assertEquals("initial_value", data.getData("plugin1", 0));

        data.setData("plugin1", 0, "updated_value");
        assertEquals("updated_value", data.getData("plugin1", 0));
    }

    @Test
    @DisplayName("Test of getDataList with default entity")
    void getDataListDefaultEntityTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", 0, "value1");
        data.setData("plugin1", 1, "value2");
        data.setData("plugin1", 2, "value3");

        List<String> result = data.getDataList("plugin1", String.class);
        assertEquals(3, result.size());
        assertTrue(result.contains("value1"));
        assertTrue(result.contains("value2"));
        assertTrue(result.contains("value3"));
    }

    @Test
    @DisplayName("Test of getDataList with specific entity")
    void getDataListSpecificEntityTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", "entity1", 0, "value1");
        data.setData("plugin1", "entity1", 1, "value2");
        data.setData("plugin1", "entity2", 0, "value3");

        List<String> result = data.getDataList("plugin1", "entity1", String.class);
        assertEquals(2, result.size());
        assertTrue(result.contains("value1"));
        assertTrue(result.contains("value2"));
        assertFalse(result.contains("value3"));
    }

    @Test
    @DisplayName("Test of getDataList with mixed types")
    void getDataListMixedTypesTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", 0, "string_value");
        data.setData("plugin1", 1, 42);
        data.setData("plugin1", 2, "another_string");

        List<String> result = data.getDataList("plugin1", String.class);
        assertEquals(2, result.size());
        assertTrue(result.contains("string_value"));
        assertTrue(result.contains("another_string"));
    }

    @Test
    @DisplayName("Test of getDataList returning empty list for non-existent plugin")
    void getDataListNonExistentPluginTest() {
        AuxiliaryData data = new AuxiliaryData();

        List<String> result = data.getDataList("nonexistent_plugin", String.class);
        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    @Test
    @DisplayName("Test of removeData with string valueId")
    void removeDataWithStringValueIdTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", "entity1", "key1", "value1");
        data.setData("plugin1", "entity1", "key2", "value2");

        assertTrue(data.removeEntityData("plugin1", "entity1", "key1"));
        assertNull(data.getData("plugin1", "entity1", "key1"));
        assertEquals("value2", data.getData("plugin1", "entity1", "key2"));
        assertFalse(data.removeEntityData("plugin1", "entity1", "missing"));
        assertFalse(data.removeEntityData("missing", "entity1", "key1"));
    }

    @Test
    @DisplayName("Test of removeEntityData with integer valueId")
    void removeEntityDataWithIntValueIdTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", "entity1", 42, "value1");

        assertTrue(data.removeEntityData("plugin1", "entity1", 42));
        assertNull(data.getData("plugin1", "entity1", 42));
        assertFalse(data.removeEntityData("plugin1", "entity1", 42));
    }

    @Test
    @DisplayName("Test of removeEntityData with object reference")
    void removeEntityDataWithObjectReferenceTest() {
        AuxiliaryData data = new AuxiliaryData();
        Object value = new Object();
        Object otherValue = new Object();
        String equalValue = new String(new char[]{
                'v',
                'a',
                'l',
                'u',
                'e'});
        String equalButDifferentReference = new String(new char[]{
                'v',
                'a',
                'l',
                'u',
                'e'});
        data.setData("plugin1", "entity1", "key1", value);
        data.setData("plugin1", "entity1", "key2", otherValue);
        data.setData("plugin1", "entity1", "key3", equalValue);

        assertFalse(data.removeEntityData("plugin1", "entity1", (Object) equalButDifferentReference));
        assertTrue(data.removeEntityData("plugin1", "entity1", value));
        assertNull(data.getData("plugin1", "entity1", "key1"));
        assertSame(otherValue, data.getData("plugin1", "entity1", "key2"));
        assertSame(equalValue, data.getData("plugin1", "entity1", "key3"));
        assertFalse(data.removeEntityData("plugin1", "entity1", value));
    }

    @Test
    @DisplayName("Test of clearEntityData")
    void clearEntityDataTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", "entity1", "key1", "value1");
        data.setData("plugin1", "entity2", "key1", "value2");
        data.setData("plugin2", "entity1", "key1", "value3");

        data.clearEntityData("plugin1", "entity1");

        assertNull(data.getData("plugin1", "entity1", "key1"));
        assertEquals("value2", data.getData("plugin1", "entity2", "key1"));
        assertEquals("value3", data.getData("plugin2", "entity1", "key1"));
        data.clearEntityData("plugin1", "missing");
        data.clearEntityData("missing", "entity1");
    }

    @Test
    @DisplayName("Test of clearEntityData removing the last plug-in entity")
    void clearEntityDataRemovesLastPluginEntityTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", "entity1", "key1", "value1");

        data.clearEntityData("plugin1", "entity1");

        assertNull(data.getData("plugin1", "entity1", "key1"));
        assertTrue(data.getDataList("plugin1", "entity1", String.class).isEmpty());
    }

    @Test
    @DisplayName("Test of persistent data with clearTemporaryData")
    void persistentDataTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", 0, "temporary_value", false);
        data.setData("plugin1", 1, "persistent_value", true);

        data.clearTemporaryData();

        assertNull(data.getData("plugin1", 0));
        assertEquals("persistent_value", data.getData("plugin1", 1));
    }

    @Test
    @DisplayName("Test of clearTemporaryData with multiple plugins and entities")
    void clearTemporaryDataComplexTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", "entity1", 0, "temp1", false);
        data.setData("plugin1", "entity1", 1, "persist1", true);
        data.setData("plugin1", "entity2", 0, "temp2", false);
        data.setData("plugin2", "entity1", 0, "temp3", false);
        data.setData("plugin2", "entity1", 1, "persist2", true);

        data.clearTemporaryData();

        assertNull(data.getData("plugin1", "entity1", 0));
        assertEquals("persist1", data.getData("plugin1", "entity1", 1));
        assertNull(data.getData("plugin1", "entity2", 0));
        assertNull(data.getData("plugin2", "entity1", 0));
        assertEquals("persist2", data.getData("plugin2", "entity1", 1));
    }

    @Test
    @DisplayName("Test of clearTemporaryData removing empty structures")
    void clearTemporaryDataRemovesEmptyStructuresTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", "entity1", 0, "temp1", false);
        data.setData("plugin1", "entity2", 0, "temp2", false);

        data.clearTemporaryData();

        assertTrue(data.getDataList("plugin1", "entity1", String.class).isEmpty());
        assertTrue(data.getDataList("plugin1", "entity2", String.class).isEmpty());
    }

    @Test
    @DisplayName("Test of clearData")
    void clearDataTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", 0, "value1", false);
        data.setData("plugin1", 1, "value2", true);
        data.setData("plugin2", "entity1", 0, "value3", false);

        data.clearData();

        assertNull(data.getData("plugin1", 0));
        assertNull(data.getData("plugin1", 1));
        assertNull(data.getData("plugin2", "entity1", 0));
        assertTrue(data.getDataList("plugin1", String.class).isEmpty());
    }

    @Test
    @DisplayName("Test of data isolation between different plugins")
    void dataIsolationBetweenPluginsTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", 0, "value1");
        data.setData("plugin2", 0, "value2");

        String result1 = data.getData("plugin1", 0, String.class);
        String result2 = data.getData("plugin2", 0, String.class);

        assertEquals("value1", result1);
        assertEquals("value2", result2);
        assertNotEquals(result1, result2);
    }

    @Test
    @DisplayName("Test of data isolation between different entities")
    void dataIsolationBetweenEntitiesTest() {
        AuxiliaryData data = new AuxiliaryData();
        data.setData("plugin1", "entity1", 0, "value1");
        data.setData("plugin1", "entity2", 0, "value2");

        String result1 = data.getData("plugin1", "entity1", 0, String.class);
        String result2 = data.getData("plugin1", "entity2", 0, String.class);

        assertEquals("value1", result1);
        assertEquals("value2", result2);
        assertNotEquals(result1, result2);
    }

    @ParameterizedTest
    @CsvSource(
            value = {
                    "plugin1|0|key1|test_value|STRING",
                    "plugin2|42|A1|12345|INTEGER",
                    "plugin3|999|cell_ref|false|BOOLEAN"
            },
            delimiter = '|'
    )
    @DisplayName("Test of getData with integer entityId and string valueId")
    void getDataWithIntEntityIdAndStringValueIdTest(
            String pluginId, int entityId, String valueId, String rawValue, TestValueType valueType
    ) {
        Object value = valueType.convert(rawValue);
        AuxiliaryData data = new AuxiliaryData();
        data.setData(pluginId, entityId, valueId, value);

        Object result = data.getData(pluginId, entityId, valueId);
        assertEquals(value, result);
    }

    /**
     * Test enum for data types
     */
    private enum TestValueType {
        STRING,
        INTEGER,
        BOOLEAN;

        /**
         * Helper method to map a test value
         * @param value Value to map
         * @return Mapped value in the expected type
         */
        private Object convert(String value) {
            return switch (this) {
                case STRING -> value;
                case INTEGER -> Integer.valueOf(value);
                case BOOLEAN -> Boolean.valueOf(value);
            };
        }
    }
}
