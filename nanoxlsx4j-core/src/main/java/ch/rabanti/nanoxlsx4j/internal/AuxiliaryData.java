package ch.rabanti.nanoxlsx4j.internal;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;

/**
 * Holds non-core data, for example data produced by a loader plug-in, for later processing by plug-ins.
 *
 * <p>The data is kept in memory and is not written to the workbook file.</p>
 */
@InternalApi
public final class AuxiliaryData {

    /** Entity identifier used by overloads that do not specify an entity. */
    private static final String DEFAULT_ENTITY_ID = "";

    private final Map<String, Map<String, Map<String, DataEntry>>> data = new HashMap<>();

    /** Creates an empty auxiliary data store. */
    public AuxiliaryData() {
    }

    /**
     * Registers or updates a temporary value for a plug-in using the default entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param valueId numeric value identifier, such as an index
     * @param value value to store
     */
    public void setData(String pluginId, int valueId, Object value) {
        setData(pluginId, valueId, value, false);
    }

    /**
     * Registers or updates a value for a plug-in using the default entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param valueId numeric value identifier, such as an index
     * @param value value to store
     * @param persistent {@code true} to retain the value when temporary data is cleared
     */
    public void setData(String pluginId, int valueId, Object value, boolean persistent) {
        setData(pluginId, Integer.toString(valueId), value, persistent);
    }

    /**
     * Registers or updates a temporary value for a plug-in using the default entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param valueId value identifier, such as a cell address
     * @param value value to store
     */
    public void setData(String pluginId, String valueId, Object value) {
        setData(pluginId, valueId, value, false);
    }

    /**
     * Registers or updates a value for a plug-in using the default entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param valueId value identifier, such as a cell address
     * @param value value to store
     * @param persistent {@code true} to retain the value when temporary data is cleared
     */
    public void setData(String pluginId, String valueId, Object value, boolean persistent) {
        setData(pluginId, DEFAULT_ENTITY_ID, valueId, value, persistent);
    }

    /**
     * Registers or updates a temporary value for a plug-in and entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     * @param valueId numeric value identifier, such as an index
     * @param value value to store
     */
    public void setData(String pluginId, String entityId, int valueId, Object value) {
        setData(pluginId, entityId, valueId, value, false);
    }

    /**
     * Registers or updates a value for a plug-in and entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     * @param valueId numeric value identifier, such as an index
     * @param value value to store
     * @param persistent {@code true} to retain the value when temporary data is cleared
     */
    public void setData(String pluginId, String entityId, int valueId, Object value, boolean persistent) {
        setData(pluginId, entityId, Integer.toString(valueId), value, persistent);
    }

    /**
     * Registers or updates a temporary value for a plug-in and a numerically identified entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId numeric entity identifier, such as a worksheet index
     * @param valueId value identifier, such as a cell address
     * @param value value to store
     */
    public void setData(String pluginId, int entityId, String valueId, Object value) {
        setData(pluginId, entityId, valueId, value, false);
    }

    /**
     * Registers or updates a value for a plug-in and a numerically identified entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId numeric entity identifier, such as a worksheet index
     * @param valueId value identifier, such as a cell address
     * @param value value to store
     * @param persistent {@code true} to retain the value when temporary data is cleared
     */
    public void setData(String pluginId, int entityId, String valueId, Object value, boolean persistent) {
        setData(pluginId, Integer.toString(entityId), valueId, value, persistent);
    }

    /**
     * Registers or updates a temporary value for a plug-in and entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     * @param valueId value identifier, such as a cell address
     * @param value value to store
     */
    public void setData(String pluginId, String entityId, String valueId, Object value) {
        setData(pluginId, entityId, valueId, value, false);
    }

    /**
     * Registers or updates a value for a plug-in and entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     * @param valueId value identifier, such as a cell address
     * @param value value to store
     * @param persistent {@code true} to retain the value when temporary data is cleared
     */
    public void setData(String pluginId, String entityId, String valueId, Object value, boolean persistent) {
        data.computeIfAbsent(pluginId, ignored -> new HashMap<>())
                .computeIfAbsent(entityId, ignored -> new HashMap<>())
                .put(valueId, new DataEntry(value, persistent));
    }

    /**
     * Retrieves a value from the default entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param valueId numeric value identifier, such as an index
     * @return stored value, or {@code null} if it was not found
     */
    public Object getData(String pluginId, int valueId) {
        return getData(pluginId, Integer.toString(valueId));
    }

    /**
     * Retrieves a value from the default entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param valueId value identifier, such as a cell address
     * @return stored value, or {@code null} if it was not found
     */
    public Object getData(String pluginId, String valueId) {
        return getData(pluginId, DEFAULT_ENTITY_ID, valueId);
    }

    /**
     * Retrieves a value for a plug-in and entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     * @param valueId numeric value identifier, such as an index
     * @return stored value, or {@code null} if it was not found
     */
    public Object getData(String pluginId, String entityId, int valueId) {
        return getData(pluginId, entityId, Integer.toString(valueId));
    }

    /**
     * Retrieves a value for a plug-in and a numerically identified entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId numeric entity identifier, such as a worksheet index
     * @param valueId value identifier, such as a cell address
     * @return stored value, or {@code null} if it was not found
     */
    public Object getData(String pluginId, int entityId, String valueId) {
        return getData(pluginId, Integer.toString(entityId), valueId);
    }

    /**
     * Retrieves a value for a plug-in and entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     * @param valueId value identifier, such as a cell address
     * @return stored value, or {@code null} if it was not found
     */
    public Object getData(String pluginId, String entityId, String valueId) {
        Map<String, Map<String, DataEntry>> pluginData = data.get(pluginId);
        if (pluginData == null) {
            return null;
        }
        Map<String, DataEntry> entityData = pluginData.get(entityId);
        if (entityData == null) {
            return null;
        }
        DataEntry entry = entityData.get(valueId);
        return entry == null ? null : entry.value();
    }

    /**
     * Retrieves a typed value from the default entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param valueId numeric value identifier, such as an index
     * @param type requested value type
     * @param <T> requested value type
     * @return typed value, or {@code null} if it was not found or has another type
     */
    public <T> T getData(String pluginId, int valueId, Class<T> type) {
        return castValue(getData(pluginId, valueId), type);
    }

    /**
     * Retrieves a typed value from the default entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param valueId value identifier, such as a cell address
     * @param type requested value type
     * @param <T> requested value type
     * @return typed value, or {@code null} if it was not found or has another type
     */
    public <T> T getData(String pluginId, String valueId, Class<T> type) {
        return castValue(getData(pluginId, valueId), type);
    }

    /**
     * Retrieves a typed value for a plug-in and entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     * @param valueId numeric value identifier, such as an index
     * @param type requested value type
     * @param <T> requested value type
     * @return typed value, or {@code null} if it was not found or has another type
     */
    public <T> T getData(String pluginId, String entityId, int valueId, Class<T> type) {
        return castValue(getData(pluginId, entityId, valueId), type);
    }

    /**
     * Retrieves a typed value for a plug-in and entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     * @param valueId value identifier, such as a cell address
     * @param type requested value type
     * @param <T> requested value type
     * @return typed value, or {@code null} if it was not found or has another type
     */
    public <T> T getData(String pluginId, String entityId, String valueId, Class<T> type) {
        return castValue(getData(pluginId, entityId, valueId), type);
    }

    /**
     * Retrieves all values of a requested type from the default entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param type requested value type
     * @param <T> requested value type
     * @return matching values, or an empty list if no values match
     */
    public <T> List<T> getDataList(String pluginId, Class<T> type) {
        return getDataList(pluginId, DEFAULT_ENTITY_ID, type);
    }

    /**
     * Retrieves all values of a requested type for a plug-in and entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     * @param type requested value type
     * @param <T> requested value type
     * @return matching values, or an empty list if no values match
     */
    public <T> List<T> getDataList(String pluginId, String entityId, Class<T> type) {
        List<T> result = new ArrayList<>();
        Map<String, Map<String, DataEntry>> pluginData = data.get(pluginId);
        if (pluginData == null) {
            return result;
        }
        Map<String, DataEntry> entityData = pluginData.get(entityId);
        if (entityData == null) {
            return result;
        }
        for (DataEntry entry : entityData.values()) {
            T value = castValue(entry.value(), type);
            if (value != null) {
                result.add(value);
            }
        }
        return result;
    }

    /**
     * Removes a value for a plug-in and entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     * @param valueId value identifier, such as a cell address
     * @return {@code true} if the value was removed; otherwise {@code false}
     */
    public boolean removeEntityData(String pluginId, String entityId, String valueId) {
        Map<String, Map<String, DataEntry>> pluginData = data.get(pluginId);
        if (pluginData == null) {
            return false;
        }
        Map<String, DataEntry> entityData = pluginData.get(entityId);
        if (entityData == null || entityData.remove(valueId) == null) {
            return false;
        }
        removeEmptyContainers(pluginId, entityId, pluginData, entityData);
        return true;
    }

    /**
     * Removes a value for a plug-in and entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     * @param valueId numeric value identifier, such as an index
     * @return {@code true} if the value was removed; otherwise {@code false}
     */
    public boolean removeEntityData(String pluginId, String entityId, int valueId) {
        return removeEntityData(pluginId, entityId, Integer.toString(valueId));
    }

    /**
     * Removes the first value that is the same object reference as the supplied value.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     * @param value object reference to remove
     * @return {@code true} if a matching object reference was removed; otherwise {@code false}
     */
    public boolean removeEntityData(String pluginId, String entityId, Object value) {
        Map<String, Map<String, DataEntry>> pluginData = data.get(pluginId);
        if (pluginData == null) {
            return false;
        }
        Map<String, DataEntry> entityData = pluginData.get(entityId);
        if (entityData == null) {
            return false;
        }
        for (Map.Entry<String, DataEntry> entry : entityData.entrySet()) {
            if (entry.getValue().value() == value) {
                return removeEntityData(pluginId, entityId, entry.getKey());
            }
        }
        return false;
    }

    /**
     * Clears all values for a plug-in and entity.
     *
     * @param pluginId plug-in UUID or other identifier
     * @param entityId entity identifier, such as a worksheet identifier
     */
    public void clearEntityData(String pluginId, String entityId) {
        Map<String, Map<String, DataEntry>> pluginData = data.get(pluginId);
        if (pluginData == null) {
            return;
        }
        pluginData.remove(entityId);
        if (pluginData.isEmpty()) {
            data.remove(pluginId);
        }
    }

    /** Clears every temporary entry while retaining persistent entries. */
    public void clearTemporaryData() {
        Iterator<Map.Entry<String, Map<String, Map<String, DataEntry>>>> pluginIterator =
                data.entrySet().iterator();
        while (pluginIterator.hasNext()) {
            Map<String, Map<String, DataEntry>> pluginData = pluginIterator.next().getValue();
            Iterator<Map.Entry<String, Map<String, DataEntry>>> entityIterator =
                    pluginData.entrySet().iterator();
            while (entityIterator.hasNext()) {
                Map<String, DataEntry> entityData = entityIterator.next().getValue();
                entityData.entrySet().removeIf(entry -> !entry.getValue().persistent());
                if (entityData.isEmpty()) {
                    entityIterator.remove();
                }
            }
            if (pluginData.isEmpty()) {
                pluginIterator.remove();
            }
        }
    }

    /** Clears every stored entry, including persistent entries. */
    public void clearData() {
        data.clear();
    }

    /**
     * Casts a stored value if it has the requested type.
     *
     * @param value stored value
     * @param type requested value type
     * @param <T> requested value type
     * @return typed value, or {@code null} if the value has another type or is {@code null}
     */
    private static <T> T castValue(Object value, Class<T> type) {
        return type.isInstance(value) ? type.cast(value) : null;
    }

    /**
     * Removes empty entity and plug-in maps after an entry was removed.
     *
     * @param pluginId plug-in identifier
     * @param entityId entity identifier
     * @param pluginData data belonging to the plug-in
     * @param entityData data belonging to the entity
     */
    private void removeEmptyContainers(
            String pluginId,
            String entityId,
            Map<String, Map<String, DataEntry>> pluginData,
            Map<String, DataEntry> entityData
    ) {
        if (entityData.isEmpty()) {
            pluginData.remove(entityId);
        }
        if (pluginData.isEmpty()) {
            data.remove(pluginId);
        }
    }

    /**
     * Stored value and its persistence state.
     *
     * @param value stored value
     * @param persistent whether the value survives {@link #clearTemporaryData()}
     */
    private record DataEntry(Object value, boolean persistent) {
    }
}
