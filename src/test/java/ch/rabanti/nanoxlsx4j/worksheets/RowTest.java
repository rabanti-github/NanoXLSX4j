package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

public class RowTest {

    public enum RowProperty {
        Hidden,
        Height
    }

    @DisplayName("Test of the addHiddenRow function with a row number")
    @Test()
    void addHiddenRowTest() {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getHiddenRows().size());
        worksheet.addHiddenRow(2);
        assertEquals(1, worksheet.getHiddenRows().size());
        assertTrue(worksheet.getHiddenRows().containsKey(2));
        assertTrue(worksheet.getHiddenRows().get(2));
        worksheet.addHiddenRow(2); // Should not add an additional entry
        assertEquals(1, worksheet.getHiddenRows().size());
    }

    @DisplayName("Test of the failing addHiddenRow function with an invalid column number")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource({
            "-1",
            "-100",
            "1048576",
    })
    void addHiddenRowFailTest(int value) {
        Worksheet worksheet = new Worksheet();
        assertThrows(RangeException.class, () -> worksheet.addHiddenRow(value));
    }

    @DisplayName("Test of the getLastRowNumber function with an empty worksheet")
    @Test()
    void getLastRowNumberTest() {
        Worksheet worksheet = new Worksheet();
        int row = worksheet.getLastRowNumber();
        assertEquals(-1, row);
    }

    @DisplayName("Test of the getLastRowNumber function with defined rows on an empty worksheet")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastRowNumberTest2(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(2, 44.4f);
        }
        int row = worksheet.getLastRowNumber();
        assertEquals(2, row);
    }

    @DisplayName("Test of the getLastRowNumber function with defined rows on an empty worksheet, where the row definition has gaps")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastRowNumberTest3(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(10);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        int row = worksheet.getLastRowNumber();
        assertEquals(10, row);
    }

    @DisplayName("Test of the getLastRowNumber function with defined rows where cells are defined below the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastRowNumberTest4(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(10);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        worksheet.addCell("test", "E5");
        int row = worksheet.getLastRowNumber();
        assertEquals(10, row);
    }

    @DisplayName("Test of the getLastRowNumber function with defined rows where cells are defined above the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastRowNumberTest5(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(2, 44.4f);
        }
        worksheet.addCell("test", "F5");
        int row = worksheet.getLastRowNumber();
        assertEquals(4, row);
    }

    @DisplayName("Test of the GetLastDataRowNumber function with an empty worksheet")
    @Test()
    void getLastDataRowNumberTest() {
        Worksheet worksheet = new Worksheet();
        int row = worksheet.getLastDataRowNumber();
        assertEquals(-1, row);
    }

    @DisplayName("Test of the GetLastDataRowNumber function with defined rows on an empty worksheet")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastDataRowNumberTest2(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(2, 44.4f);
        }
        int row = worksheet.getLastDataRowNumber();
        assertEquals(-1, row);
    }

    @DisplayName("Test of the GetLastDataRowNumber function with defined rows where cells are defined below the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastDataRowNumberTest3(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(10);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        worksheet.addCell("test", "E5");
        int row = worksheet.getLastDataRowNumber();
        assertEquals(4, row);
    }

    @DisplayName("Test of the GetLastDataRowNumber function with defined rows where cells are defined above the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastDataRowNumberTest4(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(3, 44.4f);
        }
        worksheet.addCell("test", "F5");
        int row = worksheet.getLastDataRowNumber();
        assertEquals(4, row);
    }

    @DisplayName("Test of the getCurrentRowNumber function")
    @Test()
    void getCurrentRowNumberTest() {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCurrentRowNumber());
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
        worksheet.addNextCell("test");
        worksheet.addNextCell("test");
        assertEquals(2, worksheet.getCurrentRowNumber());
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.ColumnToColumn);
        worksheet.addNextCell("test");
        worksheet.addNextCell("test");
        assertEquals(2, worksheet.getCurrentRowNumber()); // should not change
        worksheet.goToNextRow();
        assertEquals(3, worksheet.getCurrentRowNumber());
        worksheet.goToNextRow(2);
        assertEquals(5, worksheet.getCurrentRowNumber());
        worksheet.goToNextColumn(2);
        assertEquals(0, worksheet.getCurrentRowNumber()); // should reset
    }

    @DisplayName("Test of the goToNextRow function")
    @ParameterizedTest(name = "Given initial row number {0} and number {1} should lead to the row {2}")
    @CsvSource({
            "0, 0, 0",
            "0, 1, 1",
            "1, 1, 2",
            "3, 10, 13",
            "3, -1, 2",
            "3, -3, 0",
    })
    void goToNextRowTest(int initialRowNumber, int number, int expectedRowNumber) {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentRowNumber(initialRowNumber);
        worksheet.goToNextRow(number);
        assertEquals(expectedRowNumber, worksheet.getCurrentRowNumber());
    }

    @DisplayName("Test of the goToNextRow function with the option to keep the column")
    @ParameterizedTest(name = "Given start address {0} and number {1} with the option to keep the column: {2} should lead to the address {3}")
    @CsvSource({
            "A1, 0, false, A1",
            "A1, 0, true, A1",
            "A1, 1, false, A2",
            "A1, 1, true, A2",
            "C10, 1, false, A11",
            "C10, 1, true, C11",
            "R5, 5, false, A10",
            "R5, 5, true, R10",
            "F5, -3, false, A2",
            "F5, -3, true, F2",
            "F5, -4, false, A1",
            "F5, -4, true, F1",
    })
    void goToNextRowTest2(String initialAddress, int number, boolean keepColumnPosition, String expectedAddress) {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentCellAddress(initialAddress);
        worksheet.goToNextRow(number, keepColumnPosition);
        Address expected = new Address(expectedAddress);
        assertEquals(expected.Column, worksheet.getCurrentColumnNumber());
        assertEquals(expected.Row, worksheet.getCurrentRowNumber());
    }

    @DisplayName("Test of the failing goToNextRow function on invalid values")
    @ParameterizedTest(name = "Given incremental value {1} an basis {0} should lead to an exception")
    @CsvSource({
            "0, -1",
            "10, -12",
            "0, 1048576",
            "0, 1248575",
    })
    void goToNextRowFailTest(int initialValue, int value) {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentRowNumber(initialValue);
        assertEquals(initialValue, worksheet.getCurrentRowNumber());
        assertThrows(RangeException.class, () -> worksheet.goToNextRow(value));
    }

    @DisplayName("Test of the removeRowHeight function")
    @Test()
    void removeRowHeightTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.setRowHeight(2, 22.2f);
        worksheet.setRowHeight(4, 33.3f);
        assertEquals(2, worksheet.getRowHeights().size());
        worksheet.removeRowHeight(2);
        assertEquals(1, worksheet.getRowHeights().size());
        worksheet.removeRowHeight(3); // Should not cause anything
        worksheet.removeRowHeight(-1); // Should not cause anything
        assertEquals(1, worksheet.getRowHeights().size());
    }

    @DisplayName("Test of the setCurrentRowNumber function")
    @ParameterizedTest(name = "Given row number {0} should lead to the same position on the worksheet")
    @CsvSource({
            "0",
            "3",
            "1048575",
    })
    void setCurrentRowNumberTest(int row) {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCurrentRowNumber());
        worksheet.goToNextRow();
        worksheet.setCurrentRowNumber(row);
        assertEquals(row, worksheet.getCurrentRowNumber());
    }

    @DisplayName("Test of the failing setCurrentRowNumber function")
    @ParameterizedTest(name = "Given column number {0} should lead to an exception")
    @CsvSource({
            "-1",
            "-10",
            "1048576",
    })
    void setCurrentRowNumberFailTest(int row) {
        Worksheet worksheet = new Worksheet();
        assertThrows(RangeException.class, () -> worksheet.setCurrentRowNumber(row));
    }

    @DisplayName("Test of the setRowHeight function")
    @ParameterizedTest(name = "Given height {0} should lead to do a valid rowHeight definition")
    @CsvSource({
            "0f",
            "0.1f",
            "10f",
            "255f",
    })
    void setRowHeightTest(float height) {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getRowHeights().size());
        worksheet.setRowHeight(0, height);
        assertEquals(1, worksheet.getRowHeights().size());
        assertEquals(height, worksheet.getRowHeights().get(0));
        worksheet.setRowHeight(0, Worksheet.DEFAULT_ROW_HEIGHT);
        assertEquals(1, worksheet.getRowHeights().size()); // No removal so far
        assertEquals(Worksheet.DEFAULT_ROW_HEIGHT, worksheet.getRowHeights().get(0));
    }

    @DisplayName("Test of the failing setRowHeight function")
    @ParameterizedTest(name = "Given row number {0} or width {1} should lead to an exception")
    @CsvSource({
            "-1, 0f",
            "1048576, 0.0f",
            "0, -10f",
            "0, 409.51f",
            "0, 500f",
    })
    void setRowHeightFailTest(int rowNumber, float height) {
        Worksheet worksheet = new Worksheet();
        assertThrows(RangeException.class, () -> worksheet.setRowHeight(rowNumber, height));
    }

    @DisplayName("Test of the getRow function")
    @Test()
    void getRowTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(22, "B1");
        worksheet.addCell(23, "B2");
        worksheet.addCell("test", "C2");
        worksheet.addCell(true, "D2");
        worksheet.addCell(false, "B3");
        List<Cell> row = worksheet.getRow(1);
        assertEquals(3, row.size());
        assertEquals(23, row.get(0).getValue());
        assertEquals("test", row.get(1).getValue());
        assertEquals(true, row.get(2).getValue());
    }

    @DisplayName("Test of the getRow function when no values are applying")
    @Test()
    void getRowTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(22, "B1");
        worksheet.addCell(false, "B3");
        List<Cell> row = worksheet.getRow(1);
        assertEquals(0, row.size());
    }

}
