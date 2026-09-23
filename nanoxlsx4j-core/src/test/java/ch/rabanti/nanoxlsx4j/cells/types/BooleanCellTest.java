package ch.rabanti.nanoxlsx4j.cells.types;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;

class BooleanCellTest {
    private final CellTypeUtils utils = new CellTypeUtils();

    @Test
    @DisplayName("Bool value cell test: Test of the cell values, as well as proper modification")
    void boolCellTest() {
        utils.assertCellCreation(true, true, Cell.CellType.BOOL, Boolean::equals);
        utils.assertCellCreation(true, false, Cell.CellType.BOOL, Boolean::equals);
    }

    @Test
    @DisplayName("Bool value cell test with style")
    void boolCellTest2() {
        Style style = BasicStyles.getBold();
        utils.assertStyledCellCreation(true, true, Cell.CellType.BOOL, Boolean::equals, style);
        utils.assertStyledCellCreation(true, false, Cell.CellType.BOOL, Boolean::equals, style);
    }

    @ParameterizedTest
    @DisplayName("Test of the bool comparison method on cells")
    @CsvSource(
            {
                    "true, true, 0",
                    "false, false, 0",
                    "true, false, 1",
                    "false, true, -1"}
    )
    void boolCellComparisonTest(boolean value1, boolean value2, int expectedResult) {
        Cell cell1 = utils.createVariantCell(value1, utils.getCellAddress());
        Cell cell2 = utils.createVariantCell(value2, utils.getCellAddress());
        int comparison = Boolean.compare((boolean) cell1.getValue(), (boolean) cell2.getValue());
        assertEquals(expectedResult, comparison);
    }
}
