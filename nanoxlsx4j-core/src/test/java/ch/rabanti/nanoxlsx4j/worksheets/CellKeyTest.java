/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.internal.CellKey;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class CellKeyTest {

    // ── Equals(CellKey) ─────────────────────────────────

    @Test
    @DisplayName("Equals(CellKey): same column and row returns true")
    public void equalsTypedSameColRowReturnsTrue() {
        CellKey a = new CellKey(3, 7);
        CellKey b = new CellKey(3, 7);
        assertTrue(a.equals(b));
    }

    @Test
    @DisplayName("Equals(CellKey): different column, same row returns false")
    public void equalsTypedDifferentColReturnsFalse() {
        CellKey a = new CellKey(3, 7);
        CellKey b = new CellKey(4, 7);
        assertFalse(a.equals(b));
    }

    @Test
    @DisplayName("Equals(CellKey): same column, different row returns false")
    public void equalsTypedDifferentRowReturnsFalse() {
        CellKey a = new CellKey(3, 7);
        CellKey b = new CellKey(3, 8);
        assertFalse(a.equals(b));
    }

    @Test
    @DisplayName("Equals(CellKey): different column and row returns false")
    public void equalsTypedDifferentColAndRowReturnsFalse() {
        CellKey a = new CellKey(0, 0);
        CellKey b = new CellKey(1, 1);
        assertFalse(a.equals(b));
    }

    // ── Equals(object) ──────────────────────────────────

    @Test
    @DisplayName("Equals(object): boxed CellKey with same values returns true")
    public void equalsObjectBoxedSameValuesReturnsTrue() {
        CellKey a = new CellKey(5, 2);
        Object b = new CellKey(5, 2);
        assertTrue(a.equals(b));
    }

    @Test
    @DisplayName("Equals(object): boxed CellKey with different values returns false")
    public void equalsObjectBoxedDifferentValuesReturnsFalse() {
        CellKey a = new CellKey(5, 2);
        Object b = new CellKey(5, 3);
        assertFalse(a.equals(b));
    }

    @Test
    @DisplayName("Equals(object): null returns false")
    public void equalsObjectNullReturnsFalse() {
        CellKey a = new CellKey(1, 1);
        // The cast selects equals(Object); unlike a C# value type, Java's CellKey overload can accept null.
        assertFalse(a.equals((Object) null));
    }

    @Test
    @DisplayName("Equals(object): unrelated type returns false")
    public void equalsObjectUnrelatedTypeReturnsFalse() {
        CellKey a = new CellKey(1, 1);
        assertFalse(a.equals("A2"));
    }

    // ── GetHashCode ──────────────────────────────────────

    @Test
    @DisplayName("GetHashCode: equal keys produce equal hash codes")
    public void getHashCodeEqualKeysEqualHashes() {
        CellKey a = new CellKey(10, 20);
        CellKey b = new CellKey(10, 20);
        assertEquals(a.hashCode(), b.hashCode());
    }

    @ParameterizedTest
    @DisplayName("GetHashCode: no collision at boundary coordinates")
    @CsvSource({
        "0, 0",
        "16383, 0",
        "0, 1048575",
        "16383, 1048575"
    })
    public void getHashCodeBoundaryCoordinatesStable(int col, int row) {
        CellKey key = new CellKey(col, row);
        // calling twice must return the same value (deterministic)
        assertEquals(key.hashCode(), key.hashCode());
    }

    @Test
    @DisplayName("GetHashCode: keys that differ only in column produce different hashes")
    public void getHashCodeDifferentColumnDifferentHash() {
        // (row*16384)^col — for two keys with the same row, hashes differ iff columns differ
        CellKey a = new CellKey(0, 5);
        CellKey b = new CellKey(1, 5);
        assertNotEquals(a.hashCode(), b.hashCode());
    }

    @Test
    @DisplayName("GetHashCode: keys that differ only in row produce different hashes")
    public void getHashCodeDifferentRowDifferentHash() {
        // (row*16384)^col — for two keys with the same col=0, hashes differ iff rows differ
        CellKey a = new CellKey(0, 0);
        CellKey b = new CellKey(0, 1);
        assertNotEquals(a.hashCode(), b.hashCode());
    }

    // ── ToString ─────────────────────────────────────────

    @ParameterizedTest
    @DisplayName("ToString: renders the Excel address string")
    @CsvSource({
        "0, 0, A1",
        "1, 0, B1",
        "25, 0, Z1",
        "26, 0, AA1",
        "0, 9, A10",
        "2, 4, C5",
        "16383, 1048575, XFD1048576"
    })
    public void toStringReturnsExcelAddress(int col, int row, String expected) {
        CellKey key = new CellKey(col, row);
        assertEquals(expected, key.toString());
    }
}
