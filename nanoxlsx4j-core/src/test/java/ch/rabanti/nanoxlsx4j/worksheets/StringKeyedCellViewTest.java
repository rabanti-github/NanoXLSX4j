/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.HashMap;
import java.util.Map;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.internal.CellKey;
import ch.rabanti.nanoxlsx4j.internal.StringKeyedCellView;

public class StringKeyedCellViewTest {

    @Test
    @DisplayName("ContainsKey: null key returns false without throwing")
    public void containsKeyNullKeyReturnsFalse() {
        StringKeyedCellView view = buildView(new CellEntry(0, 0, "x"));
        assertFalse(view.containsKey(null));
    }

    @ParameterizedTest
    @DisplayName("ContainsKey: empty/mull string key returns false without throwing")
    @ValueSource(strings = "")
    @NullSource
    public void containsKeyEmptyKeyReturnsFalse(String key) {
        StringKeyedCellView view = buildView(new CellEntry(0, 0, "x"));
        assertFalse(view.containsKey(key));
    }

    @ParameterizedTest
    @DisplayName("ContainsKey: malformed address returns false without throwing")
    @ValueSource(
            strings = {
                    "123",
                    // digits only — no column letters
                    "!!!",
                    // non-alphanumeric
                    "AAAAA1",
                    // column part too long (> 3 letters)
                    "A0",
                    // row 0 is out of range (1-based in Excel)
                    "A",
                    // no row number
                    " A7"
                    // leading space
            }
    )
    public void containsKeyInvalidAddressReturnsFalse(String key) {
        StringKeyedCellView view = buildView(new CellEntry(0, 0, "x"));
        assertFalse(view.containsKey(key));
    }

    @ParameterizedTest
    @DisplayName("TryGetValue: empty/null key returns false and null cell")
    @ValueSource(strings = "")
    @NullSource
    public void tryGetValueNullKeyReturnsFalse(String key) {
        StringKeyedCellView view = buildView(new CellEntry(0, 0, "x"));
        Cell cell = view.get(key);
        assertFalse(view.containsKey(key));
        assertNull(cell);
    }

    @ParameterizedTest
    @DisplayName("TryGetValue: malformed address returns false and null cell")
    @ValueSource(
            strings = {
                    "123",
                    "!!!",
                    "AAAAA1",
                    "A0",
                    "A",
                    "ZZZZZ9999",
                    " A7"}
    )
    public void tryGetValueInvalidAddressReturnsFalseAndNullCell(String key) {
        StringKeyedCellView view = buildView(new CellEntry(0, 0, "x"));
        Cell cell = view.get(key);
        assertFalse(view.containsKey(key));
        assertNull(cell);
    }

    @Test
    @DisplayName("TryGetValue: valid address for existing cell returns true and correct cell")
    public void tryGetValueExistingKeyReturnsTrueAndCell() {
        StringKeyedCellView view = buildView(new CellEntry(2, 4, "hello"));

        Cell cell = view.get("C5");

        assertTrue(view.containsKey("C5"));
        assertNotNull(cell);
        assertEquals("hello", cell.getValue());
    }

    @Test
    @DisplayName("TryGetValue: valid address for absent cell returns false and null cell")
    public void tryGetValueAbsentKeyReturnsFalseAndNullCell() {
        StringKeyedCellView view = buildView(new CellEntry(0, 0, "x"));
        Cell cell = view.get("Z99");
        assertFalse(view.containsKey("Z99"));
        assertNull(cell);
    }

    @Test
    @DisplayName("IEnumerable.GetEnumerator: non-generic enumerator yields all entries")
    public void iEnumerableGetEnumeratorYieldsAllEntries() {
        StringKeyedCellView view = buildView(
                new CellEntry(2, 4, "hello"),
                new CellEntry(0, 0, "world")
        );
        int count = 0;
        for (Object item : view.entrySet()) {
            assertInstanceOf(Map.Entry.class, item);
            count++;
        }
        assertEquals(2, count);
    }

    @Test
    @DisplayName("View: backing store changes are visible and mutation is unsupported")
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
        assertThrows(
                UnsupportedOperationException.class,
                () -> view.entrySet().iterator().next().setValue(new Cell())
        );
    }

    private static StringKeyedCellView buildView(CellEntry... entries) {
        Map<CellKey, Cell> store = new HashMap<>();
        for (CellEntry entry : entries) {
            store.put(
                    new CellKey(entry.col, entry.row),
                    new Cell(entry.value, Cell.CellType.DEFAULT, entry.col, entry.row)
            );
        }
        return new StringKeyedCellView(store);
    }

    /**
     * Helper construct
     *
     * @param col   Column number
     * @param row   Row number
     * @param value Value object
     */
    private record CellEntry(int col, int row, Object value) {

    }
}
