package ch.rabanti.nanoxlsx4j.cells.types;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.styles.Style;

import java.util.function.BiFunction;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class CellTypeUtils {

    Address cellAddress;
    Workbook workbook;
    Worksheet worksheet;
    Class<?> cls;

    public CellTypeUtils(Class<?> cls) {
        this.cellAddress = new Address(0, 0);
        this.workbook = new Workbook(true);
        this.worksheet = workbook.getCurrentWorksheet();
        this.cls = cls;
    }

    public Address getCellAddress() {
        return this.cellAddress;
    }

    public <T> void assertCellCreation(T initialValue, T expectedValue, Cell.CellType expectedType, BiFunction<T, T, Boolean> comparer) {
        assertCellCreation(initialValue, expectedValue, expectedType, comparer, null);
    }

    public <T> void assertStyledCellCreation(T initialValue, T expectedValue, Cell.CellType expectedType, BiFunction<T, T, Boolean> comparer, Style style) {
        assertCellCreation(initialValue, expectedValue, expectedType, comparer, style);
    }

    private <T> void assertCellCreation(T initialValue, T expectedValue, Cell.CellType expectedType, BiFunction<T, T, Boolean> comparer, Style style) {
        Cell actualCell = new Cell(initialValue, Cell.CellType.DEFAULT, this.cellAddress);
        if (style != null) {
            actualCell.setStyle(style);
        }
        assertTrue(comparer.apply(initialValue, (T) actualCell.getValue()));
        assertEquals(cls, actualCell.getValue().getClass());
        assertEquals(expectedType, actualCell.getDataType());


        actualCell.setValue(expectedValue);
        assertTrue(comparer.apply(expectedValue, (T) actualCell.getValue()));
        if (style != null) {
            // Note: Date and Time styles are set internally and are not asserted if style is null.
            // The same applies to merged styles. These must be asserted separately
            assertEquals(style, actualCell.getCellStyle());
        }

    }

    public <T> Cell createVariantCell(T value, Address cellAddress) {
        return createVariantCell(value, cellAddress, null);
    }

    public <T> Cell createVariantCell(T value, Address cellAddress, Style style) {
        return createVariantCell(value, cellAddress, true, style);
    }

    public <T> Cell createVariantCell(T value, Address cellAddress, boolean referenceWorksheet, Style style) {
        Cell givenCell = new Cell(value, Cell.CellType.DEFAULT, cellAddress);
        if (style != null) {
            givenCell.setStyle(style);
        }
        return givenCell;
    }

}
