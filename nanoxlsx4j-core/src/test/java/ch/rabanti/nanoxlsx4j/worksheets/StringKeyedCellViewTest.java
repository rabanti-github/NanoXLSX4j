/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.internal.CellKey;
import ch.rabanti.nanoxlsx4j.internal.StringKeyedCellView;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;

import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

public class StringKeyedCellViewTest {

    @Test
    @DisplayName("ContainsKey: null key returns false without throwing")
    public void containsKeyNullKeyReturnsFalse() {
        StringKeyedCellView view = buildView(0, 0);
        assertFalse(view.containsKey(null));
    }

    @ParameterizedTest
    @DisplayName("ContainsKey: empty/mull string key returns false without throwing")
    @ValueSource(strings = "")
    @NullSource
    public void containsKeyEmptyKeyReturnsFalse(String key) {
        StringKeyedCellView view = buildView(0, 0);
        assertFalse(view.containsKey(key));
    }

    @ParameterizedTest
    @DisplayName("ContainsKey: malformed address returns false without throwing")
    @ValueSource(strings = {"123", "!!!", "AAAAA1", "A0", "A", " A7"})
    public void containsKeyInvalidAddressReturnsFalse(String key) {
        StringKeyedCellView view = buildView(0, 0);
        assertFalse(view.containsKey(key));
    }

    @ParameterizedTest
    @DisplayName("TryGetValue: empty/null key returns false and null cell")
    @ValueSource(strings = "")
    @NullSource
    public void tryGetValueNullKeyReturnsFalse(String key) {
        StringKeyedCellView view = buildView(0, 0);
        Cell cell = view.get(key);
        assertFalse(view.containsKey(key));
        assertNull(cell);
    }

    @ParameterizedTest
    @DisplayName("TryGetValue: malformed address returns false and null cell")
    @ValueSource(strings = {"123", "!!!", "AAAAA1", "A0", "A", "ZZZZZ9999", " A7"})
    public void tryGetValueInvalidAddressReturnsFalseAndNullCell(String key) {
        StringKeyedCellView view = buildView(0, 0);
        Cell cell = view.get(key);
        assertFalse(view.containsKey(key));
        assertNull(cell);
    }

    @Disabled("TODO: Requires the Cell value and coordinate API port")
    @Test
    @DisplayName("TryGetValue: valid address for existing cell returns true and correct cell")
    public void tryGetValueExistingKeyReturnsTrueAndCell() {
        // TODO: C# reference: TryGetValue_ExistingKey_ReturnsTrueAndCell.
    }

    @Test
    @DisplayName("TryGetValue: valid address for absent cell returns false and null cell")
    public void tryGetValueAbsentKeyReturnsFalseAndNullCell() {
        StringKeyedCellView view = buildView(0, 0);
        Cell cell = view.get("Z99");
        assertFalse(view.containsKey("Z99"));
        assertNull(cell);
    }

    @Test
    @DisplayName("IEnumerable.GetEnumerator: non-generic enumerator yields all entries")
    public void iEnumerableGetEnumeratorYieldsAllEntries() {
        StringKeyedCellView view = buildView(2, 4, 0, 0);
        int count = 0;
        for (Object item : view.entrySet()) {
            assertInstanceOf(Map.Entry.class, item);
            count++;
        }
        assertEquals(2, count);
    }

    @Test
    public void viewIsLiveAndReadOnly() {
        Map<CellKey, Cell> store = new HashMap<>();
        StringKeyedCellView view = new StringKeyedCellView(store);
        Cell cell = new Cell();

        store.put(new CellKey(0, 0), cell);

        assertSame(cell, view.get("A1"));
        assertThrows(UnsupportedOperationException.class, () -> view.put("B1", new Cell()));
        assertThrows(UnsupportedOperationException.class, () -> view.remove("A1"));
        assertThrows(UnsupportedOperationException.class, view::clear);
        assertThrows(UnsupportedOperationException.class, () -> view.keySet().remove("A1"));
        assertThrows(UnsupportedOperationException.class, () -> view.values().remove(cell));
        assertThrows(UnsupportedOperationException.class,
                () -> view.entrySet().iterator().next().setValue(new Cell()));
    }

    private static StringKeyedCellView buildView(int... coordinates) {
        Map<CellKey, Cell> store = new HashMap<>();
        for (int i = 0; i < coordinates.length; i += 2) {
            store.put(new CellKey(coordinates[i], coordinates[i + 1]), new Cell());
        }
        return new StringKeyedCellView(store);
    }
}
