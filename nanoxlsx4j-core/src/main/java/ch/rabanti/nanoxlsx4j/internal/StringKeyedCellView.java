/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal;

import java.util.AbstractMap;
import java.util.AbstractSet;
import java.util.Collection;
import java.util.Collections;
import java.util.Iterator;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.BiFunction;
import java.util.function.Function;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Worksheet;

/**
 * Non-materializing read-only view over the internal cell dictionary that exposes cells keyed by their rendered address
 * string (e.g. "A1"). All read paths delegate directly to the backing dictionary, translating keys on-the-fly without
 * allocating a snapshot copy. Mutation is intentionally unsupported; use the corresponding {@link Worksheet} operations
 * instead.
 * <p>This class is for internal use only.</p>
 */
public final class StringKeyedCellView extends AbstractMap<String, Cell> {

    private final Map<CellKey, Cell> store;
    private final Set<Entry<String, Cell>> entries = Collections.unmodifiableSet(new EntrySet());
    private Set<String> keys;
    private Collection<Cell> values;

    /**
     * Creates a read-only, live view of the supplied cell store.
     *
     * @param store backing cell store
     */
    public StringKeyedCellView(Map<CellKey, Cell> store) {
        this.store = Objects.requireNonNull(store, "store");
    }

    /**
     * Gets the cell associated with the supplied address.
     *
     * @param key cell address, such as {@code A1}
     * @return associated cell, or {@code null} when the address is invalid or absent
     */
    @Override
    public Cell get(Object key) {
        CellKey resolvedKey = resolveKey(key);
        return resolvedKey == null ? null : store.get(resolvedKey);
    }

    /**
     * Determines whether the backing store contains a cell at the supplied address.
     *
     * @param key cell address, such as {@code A1}
     * @return {@code true} when the address is valid and has an associated cell
     */
    @Override
    public boolean containsKey(Object key) {
        CellKey resolvedKey = resolveKey(key);
        return resolvedKey != null && store.containsKey(resolvedKey);
    }

    /**
     * Determines whether the backing store contains the supplied cell value.
     *
     * @param value cell value to find
     * @return {@code true} when the backing store contains the value
     */
    @Override
    public boolean containsValue(Object value) {
        return store.containsValue(value);
    }

    /**
     * Gets the number of cells in the backing store.
     *
     * @return number of cells
     */
    @Override
    public int size() {
        return store.size();
    }

    /**
     * Determines whether the backing store contains no cells.
     *
     * @return {@code true} when the backing store is empty
     */
    @Override
    public boolean isEmpty() {
        return store.isEmpty();
    }

    /**
     * Gets a read-only, live set of string-addressed cell entries.
     *
     * @return read-only entry set
     */
    @Override
    public Set<Entry<String, Cell>> entrySet() {
        return entries;
    }

    /**
     * Gets a read-only, live set of rendered cell addresses.
     *
     * @return read-only key set
     */
    @Override
    public Set<String> keySet() {
        if (keys == null) {
            keys = Collections.unmodifiableSet(super.keySet());
        }
        return keys;
    }

    /**
     * Gets a read-only, live collection of cells without rendering their addresses.
     *
     * @return read-only cell collection
     */
    @Override
    public Collection<Cell> values() {
        if (values == null) {
            values = Collections.unmodifiableCollection(store.values());
        }
        return values;
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @param key   cell address
     * @param value cell value
     * @return never returns normally
     * @throws UnsupportedOperationException always
     */
    @Override
    public Cell put(String key, Cell value) {
        throw readOnly();
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @param key cell address
     * @return never returns normally
     * @throws UnsupportedOperationException always
     */
    @Override
    public Cell remove(Object key) {
        throw readOnly();
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @param map mappings to add
     * @throws UnsupportedOperationException always
     */
    @Override
    public void putAll(Map<? extends String, ? extends Cell> map) {
        throw readOnly();
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @throws UnsupportedOperationException always
     */
    @Override
    public void clear() {
        throw readOnly();
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @param function replacement function
     * @throws UnsupportedOperationException always
     */
    @Override
    public void replaceAll(BiFunction<? super String, ? super Cell, ? extends Cell> function) {
        throw readOnly();
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @param key   cell address
     * @param value cell value
     * @return never returns normally
     * @throws UnsupportedOperationException always
     */
    @Override
    public Cell putIfAbsent(String key, Cell value) {
        throw readOnly();
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @param key   cell address
     * @param value expected cell value
     * @return never returns normally
     * @throws UnsupportedOperationException always
     */
    @Override
    public boolean remove(Object key, Object value) {
        throw readOnly();
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @param key      cell address
     * @param oldValue expected cell value
     * @param newValue replacement cell value
     * @return never returns normally
     * @throws UnsupportedOperationException always
     */
    @Override
    public boolean replace(String key, Cell oldValue, Cell newValue) {
        throw readOnly();
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @param key   cell address
     * @param value replacement cell value
     * @return never returns normally
     * @throws UnsupportedOperationException always
     */
    @Override
    public Cell replace(String key, Cell value) {
        throw readOnly();
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @param key             cell address
     * @param mappingFunction function used to create a cell
     * @return never returns normally
     * @throws UnsupportedOperationException always
     */
    @Override
    public Cell computeIfAbsent(String key, Function<? super String, ? extends Cell> mappingFunction) {
        throw readOnly();
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @param key               cell address
     * @param remappingFunction function used to replace a cell
     * @return never returns normally
     * @throws UnsupportedOperationException always
     */
    @Override
    public Cell computeIfPresent(
            String key,
            BiFunction<? super String, ? super Cell, ? extends Cell> remappingFunction
    ) {
        throw readOnly();
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @param key               cell address
     * @param remappingFunction function used to compute a cell
     * @return never returns normally
     * @throws UnsupportedOperationException always
     */
    @Override
    public Cell compute(String key, BiFunction<? super String, ? super Cell, ? extends Cell> remappingFunction) {
        throw readOnly();
    }

    /**
     * Mutation is unsupported by this read-only view.
     *
     * @param key               cell address
     * @param value             cell value to merge
     * @param remappingFunction function used to merge cell values
     * @return never returns normally
     * @throws UnsupportedOperationException always
     */
    @Override
    public Cell merge(
            String key, Cell value,
            BiFunction<? super Cell, ? super Cell, ? extends Cell> remappingFunction
    ) {
        throw readOnly();
    }

    private static CellKey resolveKey(Object key) {
        if (!(key instanceof String address) || address.isEmpty()) {
            return null;
        }

        int index = 0;
        int length = address.length();
        if (address.charAt(index) == '$') {
            index++;
        }

        int column = 0;
        int columnStart = index;
        while (index < length) {
            char character = address.charAt(index);
            int letter;
            if (character >= 'A' && character <= 'Z') {
                letter = character - 'A' + 1;
            } else if (character >= 'a' && character <= 'z') {
                letter = character - 'a' + 1;
            } else {
                break;
            }
            if (column > (Worksheet.MAX_COLUM_NUMBER + 1 - letter) / 26) {
                return null;
            }
            column = column * 26 + letter;
            index++;
        }
        if (index == columnStart || column > Worksheet.MAX_COLUM_NUMBER + 1) {
            return null;
        }

        if (index < length && address.charAt(index) == '$') {
            index++;
        }

        int row = 0;
        int rowStart = index;
        while (index < length) {
            char character = address.charAt(index);
            if (character < '0' || character > '9') {
                return null;
            }
            int digit = character - '0';
            if (row > (Worksheet.MAX_ROW_NUMBER + 1 - digit) / 10) {
                return null;
            }
            row = row * 10 + digit;
            index++;
        }
        if (index == rowStart || row < 1 || row > Worksheet.MAX_ROW_NUMBER + 1) {
            return null;
        }

        return new CellKey(column - 1, row - 1);
    }

    private static UnsupportedOperationException readOnly() {
        return new UnsupportedOperationException("StringKeyedCellView is read-only");
    }

    private final class EntrySet extends AbstractSet<Entry<String, Cell>> {

        /** {@inheritDoc} */
        @Override
        public Iterator<Entry<String, Cell>> iterator() {
            Iterator<Entry<CellKey, Cell>> iterator = store.entrySet().iterator();
            return new Iterator<>() {
                /** {@inheritDoc} */
                @Override
                public boolean hasNext() {
                    return iterator.hasNext();
                }

                /** {@inheritDoc} */
                @Override
                public Entry<String, Cell> next() {
                    Entry<CellKey, Cell> entry = iterator.next();
                    CellKey key = entry.getKey();
                    return new SimpleImmutableEntry<>(Cell.resolveCellAddress(key.column, key.row), entry.getValue());
                }
            };
        }

        /** {@inheritDoc} */
        @Override
        public int size() {
            return store.size();
        }

        /**
         * Mutation is unsupported by this read-only entry set.
         *
         * @throws UnsupportedOperationException always
         */
        @Override
        public void clear() {
            throw readOnly();
        }

        /**
         * Mutation is unsupported by this read-only entry set.
         *
         * @param object entry to remove
         * @return never returns normally
         * @throws UnsupportedOperationException always
         */
        @Override
        public boolean remove(Object object) {
            throw readOnly();
        }
    }
}
