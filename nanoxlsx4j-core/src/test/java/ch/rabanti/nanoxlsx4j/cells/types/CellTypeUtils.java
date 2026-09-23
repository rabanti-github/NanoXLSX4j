package ch.rabanti.nanoxlsx4j.cells.types;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.function.BiPredicate;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.styles.Style;

class CellTypeUtils {
    private final Address cellAddress = new Address(0, 0);
    private final Workbook workbook = new Workbook(true);
    private final Worksheet worksheet = workbook.getCurrentWorksheet();

    Address getCellAddress() {
        return cellAddress;
    }

    <T> void assertCellCreation(
            T initialValue, T expectedValue, Cell.CellType expectedType,
            BiPredicate<T, T> comparer
    ) {
        assertCellCreation(initialValue, expectedValue, expectedType, comparer, null);
    }

    <T> void assertStyledCellCreation(
            T initialValue, T expectedValue, Cell.CellType expectedType,
            BiPredicate<T, T> comparer, Style style
    ) {
        assertCellCreation(initialValue, expectedValue, expectedType, comparer, style);
    }

    @SuppressWarnings("unchecked")
    private <T> void assertCellCreation(
            T initialValue, T expectedValue, Cell.CellType expectedType,
            BiPredicate<T, T> comparer, Style style
    ) {
        Cell actualCell = new Cell(initialValue, Cell.CellType.DEFAULT, cellAddress);
        if (style != null) {
            actualCell.setStyle(style);
        }
        assertTrue(comparer.test(initialValue, (T) actualCell.getValue()));
        assertEquals(initialValue.getClass(), actualCell.getValue().getClass());
        assertEquals(expectedType, actualCell.getDataType());
        actualCell.setValue(expectedValue);
        assertTrue(comparer.test(expectedValue, (T) actualCell.getValue()));
        if (style != null) {
            // Note: Date and Time styles are set internally and are not asserted if style is null.
            // The same applies to merged styles. These must be asserted separately
            // Assert.Equals may fail here because of object reference comparison and not value comparison
            assertEquals(style, actualCell.getCellStyle());
        }
    }

    Cell createVariantCell(Object value, Address address) {
        return new Cell(value, Cell.CellType.DEFAULT, address);
    }
}
