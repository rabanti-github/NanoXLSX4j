package ch.rabanti.nanoxlsx4j.cells.types;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.enums.FormulaError;

class ErrorCellTest {

    @ParameterizedTest
    @DisplayName("FormulaError values resolve to standalone error cells")
    @EnumSource(
            value = FormulaError.class,
            names = {
                    "NULL",
                    "DIVISION_BY_ZERO",
                    "VALUE",
                    "REFERENCE",
                    "NAME",
                    "NUMBER",
                    "NOT_AVAILABLE",
                    "GETTING_DATA"
            }
    )
    void errorValueTest(FormulaError error) {
        Cell cell = new Cell(error, Cell.CellType.DEFAULT);

        assertEquals(Cell.CellType.ERROR, cell.getDataType());
        assertEquals(error, cell.getValue());
        assertNull(cell.getFormula());
    }

    @Test
    @DisplayName("Changing an error value resolves the new ordinary cell type")
    void errorToStringTest() {
        Cell cell = new Cell(FormulaError.REFERENCE, Cell.CellType.ERROR);

        cell.setValue("text");

        assertEquals(Cell.CellType.STRING, cell.getDataType());
        assertEquals("text", cell.getValue());
        assertNull(cell.getFormula());
    }

    @Test
    @DisplayName("Forcing an ordinary value to Error preserves the value and creates no formula metadata")
    void explicitErrorTypeTest() {
        Cell cell = new Cell("#DIV/0!", Cell.CellType.STRING);

        cell.setDataType(Cell.CellType.ERROR);

        assertEquals(Cell.CellType.ERROR, cell.getDataType());
        assertEquals("#DIV/0!", cell.getValue());
        assertNull(cell.getFormula());
    }
}
