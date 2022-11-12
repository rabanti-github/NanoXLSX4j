package ch.rabanti.nanoxlsx4j.cells.types;

import ch.rabanti.nanoxlsx4j.Cell;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class StringCellTest {

    CellTypeUtils utils;

    public StringCellTest() {
        utils = new CellTypeUtils(String.class);
    }

    @DisplayName("String value cell test: Test of the cell values, as well as proper modification")
    @ParameterizedTest(name = "Given value {0} should lead to a valid cell")
    @CsvSource(
            {
                    "''",
                    "NULL",
                    "Text",
                    "' '",
                    "start\tend",}
    )
    void StringsCellTest(String value) {
        String actualValue = resolveString(value);
        utils.assertCellCreation("Initial Value", actualValue, Cell.CellType.STRING, StringCellTest::compareString);
    }

    @DisplayName("Test of the string comparison method on cells")
    @ParameterizedTest(name = "Given value {0} compared to {1} should lead to {2}")
    @CsvSource(
            {
                    "NULL, NULL, 0",
                    "NULL, X, -1",
                    "x, NULL, 1",
                    ", , 0",
                    " ,  , 0",
                    "a, b, -1",
                    "9, 8, 1",}
    )
    void StringCellComparisonTest(String value1, String value2, int expectedResult) {
        String v1 = resolveString(value1);
        String v2 = resolveString(value2);
        Cell cell1 = utils.createVariantCell(v1, utils.getCellAddress());
        Cell cell2 = utils.createVariantCell(v2, utils.getCellAddress());
        if (cell1.getValue() == null && cell2.getValue() == null) {
            assertEquals(expectedResult, 0);
        } else if (cell1.getValue() == null && cell2.getValue() != null) {
            assertEquals(expectedResult, -1);
        } else if (cell1.getValue() != null && cell2.getValue() == null) {
            assertEquals(expectedResult, 1);
        } else {
            int comparison = ((String) cell1.getValue()).compareTo((String) cell2.getValue());
            assertEquals(comparison, expectedResult);
        }

    }

    private static String resolveString(String input) {
        if (input == null || input.equalsIgnoreCase("NULL")) {
            return null;
        } else {
            return input;
        }
    }

    private static boolean compareString(String current, String other) {
        if (current == null && other == null) {
            return true;
        } else if (current == null) {
            return false;
        }
        return current.equals(other);
    }
}
