package ch.rabanti.nanoxlsx4j.cells.types;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.Objects;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;
import ch.rabanti.nanoxlsx4j.Cell;

class StringCellTest {
    private final CellTypeUtils utils = new CellTypeUtils();

    @ParameterizedTest
    @DisplayName("String value cell test: Test of the cell values, as well as proper modification")
    @NullSource // provides null as value (empty cell)
    @ValueSource(
            strings = {
                    "",
                    "Text",
                    " ",
                    "start\tend"}
    )
    void stringsCellTest(String value) {
        utils.assertCellCreation("Initial Value", value, Cell.CellType.STRING, Objects::equals);
    }

    @ParameterizedTest
    @DisplayName("Test of the string comparison method on cells")
    @CsvSource(
            value = {
                    "NULL|NULL|0",
                    "NULL|X|-1",
                    "x|NULL|1",
                    "''|''|0",
                    "' '|' '|0",
                    "a|b|-1",
                    "9|8|1"
            },
            delimiter = '|',
            nullValues = "NULL"
    )
    void stringCellComparisonTest(String value1, String value2, int expectedResult) {
        Cell cell1 = utils.createVariantCell(value1, utils.getCellAddress());
        Cell cell2 = utils.createVariantCell(value2, utils.getCellAddress());
        String left = (String) cell1.getValue();
        String right = (String) cell2.getValue();
        int comparison = left == null ? (right == null ? 0 : -1) : (right == null ? 1 : left.compareTo(right));
        assertEquals(expectedResult, comparison);
    }
}
